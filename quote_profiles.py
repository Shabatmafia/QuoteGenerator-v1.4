"""Quote-type profiles. One app, many proposal templates.

Each profile describes one line of business:
  label              friendly name shown (subtly) in the UI
  template_glob      how to find the PowerPoint template file
  field_aliases      canonical field -> acceptable Excel column headers
  required_fields    must be non-empty for a quote to be generated
  signature_columns  columns that identify an Excel as THIS quote type
  name_hints         filename fragments for auto-detecting exports

To add a new quote type later: add a profile here, drop the template
next to the exe (or embed it at build time), and — if it isn't handled
by the generic text-replacement path — add a mapping JSON like
ho3_mapping.json.
"""

from __future__ import annotations

# ---------------------------------------------------------------------------
# AUTO (personal auto) — the original, fully working profile
# ---------------------------------------------------------------------------
AUTO_FIELD_ALIASES = {
    "Proposal Date": ["Proposal Date"],
    "Client Name": ["Client Name"],
    "Lead Producer": ["Lead Producer", "Producer"],
    "Secondary Producer": ["Secondary Producer", "Producer Role"],
    "Named Insureds": ["Named Insureds"],
    "Street": ["Street"],
    "City": ["City"],
    "State": ["State"],
    "ZIP": ["ZIP"],
    "Driver 1 Name": ["Driver 1 Name"],
    "Driver 1 DL#": ["Driver 1 DL#"],
    "Driver 1 State": ["Driver 1 State"],
    "Driver 2 Name": ["Driver 2 Name"],
    "Driver 2 DL#": ["Driver 2 DL#"],
    "Driver 2 State": ["Driver 2 State"],
    "Vehicle 1 VIN": ["Vehicle 1 VIN"],
    "Vehicle 1 Year": ["Vehicle 1 Year"],
    "Vehicle 1 Make": ["Vehicle 1 Make"],
    "Vehicle 1 Model": ["Vehicle 1 Model"],
    "Vehicle 2 VIN": ["Vehicle 2 VIN"],
    "Vehicle 2 Year": ["Vehicle 2 Year"],
    "Vehicle 2 Make": ["Vehicle 2 Make"],
    "Vehicle 2 Model": ["Vehicle 2 Model"],
    "Expiring Carrier": ["Expiring Carrier"],
    "Current BI Limit": ["Current BI Limit"],
    "Current PD Limit": ["Current PD Limit"],
    "Current UIM Limit": ["Current UIM Limit"],
    "Current MED Limit": ["Current MED Limit"],
    "Current TE Limit": ["Current TE Limit"],
    "Current Annual Premium": ["Current Annual Premium"],
    "New Carrier": ["New Carrier"],
    "New Renewal BI Limit": ["New Renewal BI Limit", "New BI Limit"],
    "New Renewal PD Limit": ["New Renewal PD Limit", "New PD Limit"],
    "New Renewal UIM Limit": ["New Renewal UIM Limit", "New UIM Limit"],
    "New Renewal MED Limit": ["New Renewal MED Limit", "New MED Limit"],
    "New Renewal TE Limit": ["New Renewal TE Limit", "New TE Limit"],
    "New Premium": ["New Premium"],
    "A.M. Best Rating": ["A.M. Best Rating"],
    "Policy Date": ["Policy Date"],
    "LOB": ["LOB"],
    "Current PIP Limit": ["Current PIP Limit"],
    "Current Collision Deductible": ["Current Collision Deductible"],
    "Current Comp Deductible": ["Current Comp Deductible"],
    "New PIP Limit": ["New PIP Limit"],
    "New Collision Deductible": ["New Collision Deductible"],
    "New Comp Deductible": ["New Comp Deductible"],
    "Billing Method": ["Billing Method"],
    # Optional image uploads (JotForm exports these as links)
    "Cover Photo": ["Cover Photo", "Cover Photo (optional)", "Photo"],
}

# ---------------------------------------------------------------------------
# HO3 (homeowners) — columns the new JotForm form will produce.
# These names match the JotForm AI prompt shipped with this project
# (jotform_ho3_prompt.txt); keep them in sync.
# ---------------------------------------------------------------------------
HO3_FIELD_ALIASES = {
    "Proposal Date": ["Proposal Date"],
    "Client Name": ["Client Name"],
    "Lead Producer": ["Lead Producer", "Producer"],
    "Secondary Producer": ["Secondary Producer"],
    "Named Insureds": ["Named Insureds"],
    # JotForm address widgets export under their own header names — accept all
    "Street": ["Street", "Property Street", "Street Address", "Address",
               "Street Address Line 1"],
    "City": ["City", "Property City"],
    "State": ["State", "Property State", "State / Province", "State/Province"],
    "ZIP": ["ZIP", "Property ZIP", "Zip", "Zip Code", "ZIP Code",
            "Postal / Zip Code", "Postal/Zip Code"],
    "Expiring Carrier": ["Expiring Carrier"],
    "Current Dwelling Limit": ["Current Dwelling Limit", "Current Coverage A"],
    "Current Other Structures Limit": ["Current Other Structures Limit", "Current Coverage B"],
    "Current Personal Property Limit": ["Current Personal Property Limit", "Current Coverage C"],
    "Current Loss of Use Limit": ["Current Loss of Use Limit", "Current Coverage D"],
    "Current Personal Liability Limit": ["Current Personal Liability Limit", "Current Coverage E"],
    "Current Medical Payments Limit": ["Current Medical Payments Limit", "Current Coverage F"],
    "Current AOP Deductible": ["Current AOP Deductible", "Current All Other Perils Deductible"],
    "Current Hurricane Deductible": ["Current Hurricane Deductible", "Current Wind Deductible"],
    "Current Annual Premium": ["Current Annual Premium"],
    "New Carrier": ["New Carrier"],
    "New Dwelling Limit": ["New Dwelling Limit", "New Coverage A"],
    "New Other Structures Limit": ["New Other Structures Limit", "New Coverage B"],
    "New Personal Property Limit": ["New Personal Property Limit", "New Coverage C"],
    "New Loss of Use Limit": ["New Loss of Use Limit", "New Coverage D"],
    "New Personal Liability Limit": ["New Personal Liability Limit", "New Coverage E"],
    "New Medical Payments Limit": ["New Medical Payments Limit", "New Coverage F"],
    "New AOP Deductible": ["New AOP Deductible", "New All Other Perils Deductible"],
    "New Hurricane Deductible": ["New Hurricane Deductible", "New Wind Deductible"],
    "New Premium": ["New Premium"],
    "A.M. Best Rating": ["A.M. Best Rating"],
    "Policy Date": ["Policy Date"],
    "LOB": ["LOB"],
    "Billing Method": ["Billing Method"],
    # Wind & flood option pricing shown on the proposal's options page
    "Wind Premium": ["Wind Premium"],
    "Flood Premium": ["Flood Premium"],
    "Flood Contents Premium": ["Flood Contents Premium"],
    "Flood Building Limit": ["Flood Building Limit"],
    "Flood Contents Limit": ["Flood Contents Limit"],
    "Home Replacement Cost": ["Home Replacement Cost"],
    # Optional image uploads (JotForm exports these as links; the app
    # downloads them automatically at generation time)
    "Cover Photo": ["Cover Photo", "Cover Photo (optional)", "Photo"],
    "New Carrier Logo": ["New Carrier Logo", "Carrier Logo", "Logo"],
    "Expiring Carrier Logo": ["Expiring Carrier Logo"],
}

PROFILES = {
    "AUTO": {
        "label": "Auto",
        "template_glob": "Proposal  PERSONAL AUTO*.pptx",
        "field_aliases": AUTO_FIELD_ALIASES,
        "required_fields": ["Client Name", "Proposal Date"],
        # Columns that mark an export as an Auto export
        "signature_columns": ["Vehicle 1 VIN", "Driver 1 Name"],
        "name_hints": ("auto_insurance_proposal", "personal_auto"),
        # Auto uses the original hand-tuned generation path
        "mapping_json": None,
    },
    "HO3": {
        "label": "Home (HO3)",
        "template_glob": "HO3 Proposal*.pptx",
        "field_aliases": HO3_FIELD_ALIASES,
        "required_fields": ["Client Name", "Proposal Date"],
        # Columns that mark an export as a Homeowners export
        "signature_columns": ["Current Dwelling Limit", "Current Hurricane Deductible"],
        "name_hints": ("ho3", "homeowner", "home_insurance_proposal"),
        # HO3 uses the generic path driven by this mapping file...
        "mapping_json": "ho3_mapping.json",
        # ...plus table-aware structural filling (coverage tables, premium
        # summary, locations) because the template reuses identical sample
        # values for different fields — plain text replacement can't
        # distinguish them.
        "structural": "ho3",
    },
}

# Filename fragments shared by all profiles (JotForm exports generally)
COMMON_NAME_HINTS = ("proposal_form", "jotform")


def norm_header(s) -> str:
    """Normalize a column header for tolerant matching (case, spacing)."""
    return " ".join(str(s).strip().lower().split())


def detect_profile(columns) -> str | None:
    """Given the Excel's column headers, decide which quote type it is.

    A profile matches when ANY of its signature columns (or their aliases)
    are present. Matching is case- and spacing-insensitive because JotForm
    sometimes tweaks header text. Auto is checked first for compatibility.
    """
    cols = {norm_header(c) for c in columns}
    for key in ("AUTO", "HO3"):
        profile = PROFILES[key]
        for sig_field in profile["signature_columns"]:
            aliases = profile["field_aliases"].get(sig_field, [sig_field])
            if any(norm_header(a) in cols for a in aliases):
                return key
    return None
