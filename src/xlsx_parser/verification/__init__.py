"""
Stage verification system for the Excellent Spreadsheet Analysis Algorithm.

Maps the XLSXParser pipeline output to the 11-stage Excellent algorithm
and produces a verification report showing implementation status, metrics,
gaps, and recommendations.
"""

from .stage_verifier import (
    ExcellentStage,
    ImplementationStatus,
    StageResult,
    StageVerifier,
    VerificationReport,
)

__all__ = [
    "StageVerifier",
    "VerificationReport",
    "StageResult",
    "ExcellentStage",
    "ImplementationStatus",
]
