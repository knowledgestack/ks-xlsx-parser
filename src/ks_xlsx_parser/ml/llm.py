"""Singleton llama.cpp wrapper + on-disk decision cache.

The four pretrained-model classifiers in this package (boundary,
header, role, self_check) all funnel through this one wrapper so the
~1 GB Qwen2.5-1.5B model is loaded into memory exactly once per
process, and so the (workbook_hash, prompt_version, input_hash)
decision cache is shared across them.

Why a cache: a full SpreadsheetBench 912 pass triggers ~40–60k
classification calls. At 100 ms/call on Apple Metal that's about
70 minutes per bench run. Iterating on the segmenter without a cache
means re-paying that 70 minutes per code change. With the cache,
structural changes re-validate in ~14 minutes against frozen LLM
output — the cache invalidates only when the prompt version bumps
or the input changes.

Free-to-run guarantees:
* Weights: Qwen2.5-1.5B-Instruct, Apache 2.0
* Runtime: llama-cpp-python, MIT
* No API calls at runtime. One-time HF Hub download into
  ``~/.cache/ks_xlsx_parser/models/`` (~1 GB).
* Graceful fallback: if ``llama_cpp`` or the model file is missing,
  ``classify`` raises ``LLMUnavailable`` and callers fall back to
  the deterministic heuristic path.
"""
from __future__ import annotations

import hashlib
import json
import logging
import os
import threading
from dataclasses import dataclass
from pathlib import Path
from typing import Any

logger = logging.getLogger(__name__)

# Cache location — XDG-style; users can override via env.
_DEFAULT_CACHE_HOME = Path(
    os.environ.get("KS_ML_CACHE_DIR")
    or os.environ.get("XDG_CACHE_HOME", str(Path.home() / ".cache"))
) / "ks_xlsx_parser"

MODEL_DIR = _DEFAULT_CACHE_HOME / "models"
DECISIONS_DIR = _DEFAULT_CACHE_HOME / "decisions"

# Canonical model — small enough to ship the download story (~1 GB Q4),
# capable enough for zero-shot table-structure classification.
MODEL_REPO = "bartowski/Qwen2.5-1.5B-Instruct-GGUF"
MODEL_FILE = "Qwen2.5-1.5B-Instruct-Q4_K_M.gguf"
MODEL_VERSION = "qwen2_5-1_5b-q4_k_m@v1"


class LLMUnavailable(RuntimeError):
    """Raised when llama-cpp-python or the model file aren't usable.

    Callers MUST catch this and fall back to the deterministic
    heuristic path. The library's `pip install` story is "ML is
    optional"; nothing in the parser is allowed to hard-depend on it.
    """


@dataclass
class LLMConfig:
    """Inference knobs. Defaults tuned for Apple Metal + small batch.

    n_ctx is generous (4096) so the boundary classifier can include
    sample rows from both regions plus features without truncation.
    n_gpu_layers=-1 offloads everything to Metal when the build has
    Metal support; on a Metal-less build it harmlessly maps to CPU.
    """

    n_ctx: int = 4096
    n_threads: int = max(2, (os.cpu_count() or 4) - 1)
    n_gpu_layers: int = -1
    seed: int = 1337
    temperature: float = 0.0
    max_tokens: int = 64


_lock = threading.Lock()
_singleton: "LLMSingleton | None" = None


class LLMSingleton:
    """Wraps a single llama.cpp `Llama` instance + a JSON decision cache.

    Public surface: ``classify(prompt_version, input_text)`` returns the
    cached or freshly-generated text completion. Callers parse the
    response into their own task-specific label.
    """

    def __init__(self, config: LLMConfig | None = None) -> None:
        self.config = config or LLMConfig()
        self._llama = None  # Lazy — don't pay model-load cost on import.
        self._cache: dict[str, str] = {}
        self._cache_path: Path | None = None

    def _ensure_loaded(self) -> None:
        if self._llama is not None:
            return
        try:
            from llama_cpp import Llama
        except ImportError as exc:
            raise LLMUnavailable(
                "llama-cpp-python is not installed. "
                "Run `pip install ks-xlsx-parser[ml]` or "
                "`CMAKE_ARGS='-DGGML_METAL=on' pip install llama-cpp-python`."
            ) from exc

        model_path = MODEL_DIR / MODEL_FILE
        if not model_path.exists():
            try:
                from huggingface_hub import hf_hub_download
            except ImportError as exc:
                raise LLMUnavailable(
                    "huggingface_hub not installed; can't auto-download "
                    "the model weights. Either install it or place "
                    f"{MODEL_FILE} manually at {MODEL_DIR}."
                ) from exc
            MODEL_DIR.mkdir(parents=True, exist_ok=True)
            logger.warning(
                "ks-xlsx-parser ML: downloading %s (~1 GB) on first use; "
                "cached at %s",
                MODEL_FILE, MODEL_DIR,
            )
            hf_hub_download(
                repo_id=MODEL_REPO, filename=MODEL_FILE,
                local_dir=MODEL_DIR,
            )

        # The Llama ctor downloads-on-import only when from_pretrained is
        # used; we already have the file, so a direct path keeps load
        # silent + offline-safe.
        self._llama = Llama(
            model_path=str(model_path),
            n_ctx=self.config.n_ctx,
            n_threads=self.config.n_threads,
            n_gpu_layers=self.config.n_gpu_layers,
            seed=self.config.seed,
            verbose=False,
        )

    def bind_cache(self, cache_path: Path) -> None:
        """Use ``cache_path`` (a JSON file) as the persistent decision store.

        Calling without ``bind_cache`` runs an in-memory cache only.
        Bench harnesses should always bind — that's the whole point of
        the cache (cross-run reuse of expensive LLM calls).
        """
        self._cache_path = cache_path
        if cache_path.exists():
            try:
                self._cache = json.loads(cache_path.read_text())
            except json.JSONDecodeError:
                logger.warning("Decision cache %s corrupt; starting fresh.",
                               cache_path)
                self._cache = {}

    def _persist(self) -> None:
        if self._cache_path is None:
            return
        self._cache_path.parent.mkdir(parents=True, exist_ok=True)
        tmp = self._cache_path.with_suffix(".tmp")
        tmp.write_text(json.dumps(self._cache, separators=(",", ":")))
        tmp.replace(self._cache_path)

    @staticmethod
    def _cache_key(prompt_version: str, input_text: str) -> str:
        # Hash everything that could change the LLM's output. The model
        # version is baked in so a model swap invalidates the cache.
        h = hashlib.sha256()
        h.update(MODEL_VERSION.encode())
        h.update(b"\x00")
        h.update(prompt_version.encode())
        h.update(b"\x00")
        h.update(input_text.encode())
        return h.hexdigest()[:32]

    def classify(self, prompt_version: str, prompt: str) -> str:
        """Run a single classification call; cached.

        ``prompt`` is the full text sent to the model (system + user +
        examples — caller's responsibility to format). ``prompt_version``
        is a short stable label like ``boundary_v1`` so that prompt
        edits invalidate stale cache entries.
        """
        key = self._cache_key(prompt_version, prompt)
        if key in self._cache:
            return self._cache[key]

        self._ensure_loaded()
        assert self._llama is not None

        out = self._llama.create_completion(
            prompt=prompt,
            max_tokens=self.config.max_tokens,
            temperature=self.config.temperature,
            seed=self.config.seed,
        )
        text = out["choices"][0]["text"].strip()
        self._cache[key] = text
        # Flush every 50 new entries to bound data loss on crash.
        if len(self._cache) % 50 == 0:
            self._persist()
        return text

    def flush(self) -> None:
        """Force-write the cache. Call at end of a bench run."""
        self._persist()

    def stats(self) -> dict[str, Any]:
        return {
            "model_version": MODEL_VERSION,
            "cache_entries": len(self._cache),
            "cache_path": str(self._cache_path) if self._cache_path else None,
        }


def get_llm(config: LLMConfig | None = None) -> LLMSingleton:
    """Return the process-wide singleton; create lazily."""
    global _singleton
    with _lock:
        if _singleton is None:
            _singleton = LLMSingleton(config=config)
    return _singleton


def is_enabled() -> bool:
    """True iff the caller should attempt to use the LLM.

    Two off-switches:
      - ``KS_ML=0`` env var explicitly disables (testing, CI).
      - The llama_cpp import or model file being missing is treated
        as "off" at the call site via the LLMUnavailable exception.

    Default is ON when both are present — agents and benches should
    set ``KS_ML=1`` explicitly to make the choice visible in logs.
    """
    return os.environ.get("KS_ML", "0") == "1"
