"""Built-in rules used by the Excess summary candidate selection."""

from __future__ import annotations

from typing import Any


# Source: Item_fix.xlsx / Type!A2:A23.
# Keep this list in the application so report generation does not depend on a
# copy of Item_fix.xlsx being present in the user's Downloads folder.
SUMMARY_ALLOWED_ITEM_TYPES = frozenset(
    {
        "DP",
        "Film",
        "IC",
        "Interconnection",
        "LGP",
        "Light Bar",
        "Metal",
        "Module",
        "OC",
        "PCB",
        "PCBA",
        "PD2-1",
        "PD2-2",
        "PD2-3",
        "PD2-O",
        "Plastic",
        "Power",
        "Remote",
        "SMT-I-1",
        "SMT-O-1",
        "SOC",
        "SPK",
    }
)
SUMMARY_ALLOWED_ITEM_TYPE_KEYS = frozenset(
    value.casefold() for value in SUMMARY_ALLOWED_ITEM_TYPES
)


def is_allowed_summary_item_type(value: Any) -> bool:
    """Return whether an Item TYPE is allowed for summary ranking."""
    if value in (None, ""):
        return False
    return str(value).strip().casefold() in SUMMARY_ALLOWED_ITEM_TYPE_KEYS
