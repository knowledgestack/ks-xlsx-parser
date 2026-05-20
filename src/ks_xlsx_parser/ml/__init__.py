"""ML-augmented helpers for the segmentation / chunking pipeline.

Each module under here provides a small **pretrained** model wrapper —
nothing in this package needs training, by design. The library stays
free-to-run for anyone: the LLM weights are Apache 2.0 (Qwen2.5),
inference runs on CPU or Apple Metal via `llama-cpp-python` (MIT),
the model file is downloaded once and cached locally.

Modules:
    llm           Singleton llama.cpp wrapper + decision cache.
                  Backs all four pretrained-model classifiers below.

    boundary      "Are these two adjacent regions the same table?"
                  Hooks into analysis.table_grouper.

    header        "Which rows of this region are the column headers?"
                  Hooks into analysis.light_block_detector / block_splitter.

    role          "What kind of block is this — data_table / key_value /
                  notes / totals / template / dashboard?"
                  Hooks into chunking.chunker for role-aware rendering.

    self_check    "Is this chunk self-contained, or does it need
                  neighbour context?"
                  Hooks into chunking.chunker post-emit feedback.

All four classifiers gate behind an env var + graceful fallback:
when llama-cpp-python is unavailable or the model weights aren't
present, the wrapper short-circuits and the segmenter uses the
deterministic heuristic path. The default `pip install ks-xlsx-parser`
ships with everything OFF — opt in via `pip install
ks-xlsx-parser[ml]` and `KS_ML=1`.
"""
