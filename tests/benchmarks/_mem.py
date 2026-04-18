"""
Cross-platform peak RSS helper for the Python adapter side.

macOS: `ru_maxrss` is bytes.
Linux: `ru_maxrss` is kilobytes.

Both adapters report peak memory with ±30% noise — not precise. This is
an intentional limitation documented in the benchmark README. For tight
memory measurement, wrap the whole worker in `/usr/bin/time -v`.
"""



import resource
import sys


def peak_rss_mb() -> float:
    """Peak resident set size of the current process, in megabytes."""
    raw = resource.getrusage(resource.RUSAGE_SELF).ru_maxrss
    if sys.platform == "darwin":
        return raw / (1024 * 1024)
    return raw / 1024
