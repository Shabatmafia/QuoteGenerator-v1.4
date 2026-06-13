"""Excel loading and validation for JotForm quote submissions.

Responsible for:
- Reading the .xlsx export
- Detecting WHICH quote type it is (Auto, HO3, ...) from its columns
- Resolving expected columns through the matching profile's aliases
- Returning one canonical record (pd.Series) per submission row
"""

from __future__ import annotations

from pathlib import Path
from typing import Any, List, Optional, Tuple

import pandas as pd

import quote_profiles

# ---------------------------------------------------------------------------
# Per-quote-type column aliases now live in quote_profiles.py.
# ---------------------------------------------------------------------------

# Optional identifying columns that JotForm exports sometimes include.
# Shown in the table if present, but never required.
OPTIONAL_ID_ALIASES = {
    "Submission Date": ["Submission Date", "Submission Date and Time", "Created At"],
    "Submission ID": ["Submission ID", "SubmissionID", "Quote ID", "ID"],
    "Company Name": ["Company Name", "Company"],
}

# Fields that must contain a value for a quote to be generated.
REQUIRED_ROW_FIELDS = ["Client Name", "Proposal Date"]


class ExcelLoadError(Exception):
    """Raised when the Excel file cannot be loaded or is invalid."""


def load_submissions(
    excel_path: Path,
    expected_profile: Optional[str] = None,
) -> Tuple[List[pd.Series], str]:
    """Read a JotForm Excel export.

    Returns (rows, profile_key) where profile_key identifies the quote
    type ("AUTO", "HO3", ...). If expected_profile is given (because the
    user already chose Auto or Homeowners in the app), the file is checked
    against that choice. Raises ExcelLoadError with a user-friendly
    message on any problem.
    """
    excel_path = Path(excel_path)

    # --- File-level validation (friendly messages only, no raw errors) ----
    if not excel_path.exists():
        raise ExcelLoadError(
            "That file could not be found.\n\nPlease try selecting it again."
        )
    if excel_path.suffix.lower() not in (".xlsx", ".xlsm"):
        raise ExcelLoadError(
            "That doesn't look like an Excel file.\n\n"
            "Please select the .xlsx file exported from JotForm."
        )

    try:
        df = pd.read_excel(excel_path, engine="openpyxl")
    except Exception as exc:  # corrupt file, password-protected, etc.
        raise ExcelLoadError(
            "This file could not be opened.\n\n"
            "Wrong file selected — please export from JotForm again."
        ) from exc

    if df.empty:
        raise ExcelLoadError(
            "This file contains no submissions.\n\n"
            "Please export from JotForm again."
        )

    # --- Which quote type is this export? --------------------------------
    detected = quote_profiles.detect_profile(df.columns)
    if expected_profile is not None:
        if detected is not None and detected != expected_profile:
            detected_label = quote_profiles.PROFILES[detected]["label"]
            wanted_label = quote_profiles.PROFILES[expected_profile]["label"]
            raise ExcelLoadError(
                f"This file contains {detected_label} quotes, but you chose "
                f"{wanted_label}.\n\nPlease pick the matching file."
            )
        profile_key = expected_profile
    else:
        if detected is None:
            raise ExcelLoadError(
                "Wrong file selected — please export from JotForm again.\n\n"
                "(This file doesn't look like an Auto or Home quote export.)"
            )
        profile_key = detected
    field_aliases = quote_profiles.PROFILES[profile_key]["field_aliases"]

    # --- Column resolution -------------------------------------------------
    # Tolerant: headers are matched ignoring case and extra spaces, and a
    # missing non-essential column just leaves that part of the quote blank
    # instead of rejecting the whole file. Only the essentials are mandatory.
    norm_map = {}
    for col in df.columns:
        norm_map.setdefault(quote_profiles.norm_header(col), col)

    resolved_columns = {}
    for field, aliases in field_aliases.items():
        for alias in aliases:
            col = norm_map.get(quote_profiles.norm_header(alias))
            if col is not None:
                resolved_columns[field] = col
                break

    for essential in ("Client Name", "Proposal Date"):
        if essential not in resolved_columns:
            raise ExcelLoadError(
                "Wrong file selected — please export from JotForm again.\n\n"
                f"(No '{essential}' column was found in this file.)"
            )

    # Optional identifying columns (not required)
    optional_columns = {}
    for field, aliases in OPTIONAL_ID_ALIASES.items():
        found = next((alias for alias in aliases if alias in df.columns), None)
        if found is not None:
            optional_columns[field] = found

    # --- Build one canonical record per row ------------------------------
    rows: List[pd.Series] = []
    for idx in range(len(df)):
        canonical = {field: df.iloc[idx][col] for field, col in resolved_columns.items()}
        for field, col in optional_columns.items():
            canonical[field] = df.iloc[idx][col]
        rows.append(pd.Series(canonical))
    return rows, profile_key


def display_value(value: Any) -> str:
    """Human-friendly rendering of a cell value for the table / preview."""
    if pd.isna(value):
        return ""
    if isinstance(value, pd.Timestamp):
        return str(value.date())
    return str(value).strip()


def missing_required_fields(row: pd.Series) -> List[str]:
    """Return the names of required fields that are empty in this row."""
    return [f for f in REQUIRED_ROW_FIELDS if not display_value(row.get(f))]


def get_optional(row: pd.Series, field: str) -> Optional[str]:
    """Return an optional field's display value, or None if absent."""
    if field in row.index:
        return display_value(row[field])
    return None


# ---------------------------------------------------------------------------
# Auto-detection of the latest JotForm export, so the user can skip the
# file dialog entirely. Only files whose names look like the JotForm export
# are considered (we never silently open unrelated spreadsheets).
# ---------------------------------------------------------------------------
EXPORT_NAME_HINTS = tuple(
    set(quote_profiles.COMMON_NAME_HINTS).union(
        *(p["name_hints"] for p in quote_profiles.PROFILES.values())
    )
)


def find_export_candidates(limit: int = 5) -> List[Path]:
    """Newest JotForm-looking .xlsx files from Downloads and Desktop."""
    candidates: List[Path] = []
    for folder in (Path.home() / "Downloads", Path.home() / "Desktop"):
        try:
            if folder.exists():
                candidates.extend(
                    p for p in folder.glob("*.xlsx")
                    if not p.name.startswith("~$")
                    and any(h in p.name.lower() for h in EXPORT_NAME_HINTS)
                )
        except Exception:
            continue  # an unreadable folder must never crash the app
    candidates.sort(key=lambda p: p.stat().st_mtime, reverse=True)
    return candidates[:limit]
