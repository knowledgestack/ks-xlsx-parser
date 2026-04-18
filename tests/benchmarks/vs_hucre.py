"""
Entry point: head-to-head benchmark of ks-xlsx-parser vs hucre (TypeScript).

Usage (from repo root, with venv active):

    python -m tests.benchmarks.vs_hucre \\
        --corpus testBench \\
        --out tests/benchmarks/reports \\
        [--subset real_world,enterprise] \\
        [--sample 50] \\
        [--per-file-timeout 120]

Outputs (under `--out`/<ISO>_<git-sha>/):
    results.csv   — one row per (file, parser)
    raw.ndjson    — full per-row records (nullable fields preserved)
    failures.jsonl — rows where status != ok
    summary.md    — aggregate counts, capability matrix, perf percentiles
    drift.md      — per-feature disagreement between parsers
    manifest.json — run metadata
"""



import argparse
import subprocess
import sys
from datetime import UTC, datetime
from pathlib import Path

from ._driver import enumerate_corpus, run_benchmark
from ._runner import hucre_runner, ks_runner


def main(argv: list[str] | None = None) -> int:
    parser = argparse.ArgumentParser(description=__doc__.splitlines()[1] if __doc__ else "")
    parser.add_argument("--corpus", type=Path, default=Path("testBench"),
                        help="Corpus directory containing .xlsx/.xlsm files.")
    parser.add_argument("--out", type=Path, default=Path("tests/benchmarks/reports"),
                        help="Root directory for reports; a timestamped subdir is created.")
    parser.add_argument("--subset", type=str, default=None,
                        help="Comma-separated top-level sub-dirs under corpus (e.g. real_world,enterprise)")
    parser.add_argument("--sample", type=int, default=None,
                        help="If set, randomly sample N files (seeded).")
    parser.add_argument("--seed", type=int, default=1337)
    parser.add_argument("--per-file-timeout", type=float, default=120.0)
    parser.add_argument("--parsers", type=str, default="ks,hucre",
                        help="Comma-separated subset of parsers to run.")
    parser.add_argument("--ks-python", type=str, default=sys.executable,
                        help="Python binary for the ks adapter (default: current).")
    parser.add_argument("--batch-size", type=int, default=50,
                        help="Files per worker before respawn.")
    args = parser.parse_args(argv)

    corpus = args.corpus.resolve()
    if not corpus.exists():
        sys.stderr.write(f"corpus not found: {corpus}\n")
        return 2

    subset = [s.strip() for s in args.subset.split(",")] if args.subset else None
    files = enumerate_corpus(corpus, subset, args.sample, args.seed)
    if not files:
        sys.stderr.write("no files matched the corpus filter\n")
        return 2

    # Build runners
    runners = {}
    selected = {s.strip() for s in args.parsers.split(",")}
    if "ks" in selected:
        r = ks_runner(python_bin=args.ks_python, timeout_s=args.per_file_timeout)
        r.cfg.batch_size = args.batch_size
        runners["ks-xlsx-parser"] = r
    if "hucre" in selected:
        r = hucre_runner(timeout_s=args.per_file_timeout)
        r.cfg.batch_size = args.batch_size
        runners["hucre"] = r
    if not runners:
        sys.stderr.write("no parsers selected\n")
        return 2

    # Resolve git sha for the run subdir name
    try:
        sha = subprocess.check_output(
            ["git", "rev-parse", "--short", "HEAD"],
            text=True, stderr=subprocess.DEVNULL).strip()
    except Exception:  # noqa: BLE001
        sha = "nogit"
    stamp = datetime.now(UTC).strftime("%Y%m%dT%H%M%S")
    out_dir = args.out.resolve() / f"{stamp}_{sha}"

    sys.stderr.write(
        f"benchmark: {len(files)} files × {len(runners)} parsers = "
        f"{len(files) * len(runners)} runs\n"
        f"  parsers: {sorted(runners.keys())}\n"
        f"  timeout: {args.per_file_timeout:.0f}s per file\n"
        f"  out:     {out_dir}\n\n"
    )

    try:
        run_benchmark(files=files, runners=runners, out_dir=out_dir, seed=args.seed)
    finally:
        for r in runners.values():
            r.stop()

    sys.stderr.write(f"\nreports written to: {out_dir}\n")
    sys.stderr.write(f"  - {out_dir}/summary.md\n")
    sys.stderr.write(f"  - {out_dir}/drift.md\n")
    sys.stderr.write(f"  - {out_dir}/results.csv\n")
    return 0


if __name__ == "__main__":
    sys.exit(main())
