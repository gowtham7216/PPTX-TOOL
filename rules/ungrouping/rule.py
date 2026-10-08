"""
SlideGuard Selective Ungrouping Rule.
"""

from rules.ungrouping.ungroup_rule import SelectiveUngroupingRule

# Backward-compatible alias
UngroupingRule = SelectiveUngroupingRule

__all__ = ["SelectiveUngroupingRule", "UngroupingRule"]
