"""xlayer: a Python transaction layer for ``.xlsx`` workbook changes.

Experimental alpha: bounded stored facts and approval-bound existing-cell
scalar edits. No calculation, rendering or business-outcome verification.
Supported imports are listed here; underscore implementation paths are private.
"""

__version__ = "0.1.0a1"

from xlayer._approval import Approval, VerifiedApproval
from xlayer._dependencies import ImpactLimits
from xlayer._errors import ClosedWorkbookError, Refusal, WorkbookOpenError, XlayerError
from xlayer._inspection import Inspection, ReadLimits
from xlayer._ooxml.archive import ArchiveLimits
from xlayer._preview import Preview
from xlayer._public_edits import SetValue
from xlayer._public_proposal import Proposal
from xlayer._public_workbook import Workbook
from xlayer._receipt import Receipt

__all__ = [
    "Approval",
    "ArchiveLimits",
    "ClosedWorkbookError",
    "ImpactLimits",
    "Inspection",
    "Preview",
    "Proposal",
    "ReadLimits",
    "Receipt",
    "Refusal",
    "SetValue",
    "VerifiedApproval",
    "Workbook",
    "WorkbookOpenError",
    "XlayerError",
    "__version__",
]
