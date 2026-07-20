"""Central configuration for TCR extraction.

This module contains the adjustable parts of the parser: header aliases,
BOM column aliases, confidence weights and output columns. When a real TCR
variant uses another wording, add it here instead of changing processing code.
"""

from __future__ import annotations

# Header aliases are matched with normalized keys that ignore case, spaces and punctuation.
HEADER_ALIASES = {
    "Project": [
        "project", "project name", "program", "program name", "project designation",
        "tcr project", "vehicle project", "customer project",
    ],
    "SOP": [
        "sop", "start of production", "sop date", "sop timing", "sop year",
    ],
    "FOTP": [
        "fotp", "first off tool parts", "first off tool part", "first off tool",
        "fot", "fot date", "fotp date",
    ],
    "MOB Presentation Date": [
        "mob presentation date", "make or buy presentation date", "presentation date",
        "mob date", "decision board date", "mob board date", "make buy date",
    ],
    "Production Plant": [
        "production plant", "plant", "prod plant", "prod. plant", "manufacturing plant",
        "production location", "mfg plant", "plant of production",
    ],
    "EOP": [
        "eop", "end of production", "eop date", "end production",
    ],
    "Volume": [
        "volume", "annual volume", "volumes", "qty", "quantity", "total volume",
        "volume p.a.", "volume pa", "yearly volume",
    ],
}

# BOM aliases are also matched with normalized keys that ignore case, spaces,
# underscores, dashes, dots and other punctuation.
BOM_COLUMN_ALIASES = {
    # Used for BOM detection and diagnostics; not exported as a dedicated column.
    "BOM Pos": [
        "bom pos", "bom position", "bom pos.", "position", "pos", "item", "item no",
        "item number", "bom item", "no", "no.", "nr", "nr.",
    ],
    "Reference Part Number": [
        "reference part number", "ref part number", "reference pn", "ref pn",
        "ref. part no", "ref part no", "reference part no", "reference part no.",
        "referenzteilenummer", "reference material number",
    ],
    "Part Number": [
        "part number", "pn", "part no", "part no.", "part-no", "part_number",
        "material no", "material number", "artikelnummer", "component number",
    ],
    "Complexity Category": [
        "complexity category", "complexity", "category", "complexity class",
        "complexity level", "comp category",
    ],
    "Final Decision Board": [
        "final decision board", "final decision", "decision board", "mob decision",
        "make or buy decision", "make/buy decision", "decision", "board decision",
    ],
    "Part Name": [
        "part name", "part description", "description", "part", "designation",
        "component name", "component description", "name", "benennung",
    ],
    "Part Size X": ["part size x", "size x", "x", "length", "x dimension", "dim x"],
    "Part Size Y": ["part size y", "size y", "y", "width", "y dimension", "dim y"],
    "Part Size Z": ["part size z", "size z", "z", "height", "z dimension", "dim z"],
    "Material": ["material", "raw material", "material type", "resin", "plastic material"],
    "Material Key Number": [
        "material key number", "material key", "mat key", "material key no",
        "material key no.", "material number", "mat no", "sap material",
    ],
    "Surface": ["surface", "surface treatment", "surface finish", "finish"],
    "Metallized": ["metallized", "metallised", "metalized", "metalised", "metallization"],
    "Special Surface": [
        "special surface", "special process surface", "special treatment", "special finish",
    ],
    "Weight": ["weight", "part weight", "weight g", "weight [g]", "weight kg", "mass"],
    "Machine/Tonnage": [
        "machine/tonnage", "machine", "tonnage", "press", "injection machine",
        "machine tonnage", "clamping force", "press tonnage",
    ],
    "Process (1K/2K/3K)": [
        "process", "1k/2k/3k", "1k", "2k", "3k", "process 1k/2k/3k",
        "technology", "injection process", "k process",
    ],
}

# Weighted confidence for BOM header detection. Strong identifiers are worth more
# than generic columns such as size/weight/material.
BOM_CONFIDENCE_WEIGHTS = {
    "BOM Pos": 18,
    "Reference Part Number": 30,
    "Part Number": 25,
    "Part Name": 25,
    "Complexity Category": 20,
    "Final Decision Board": 22,
    "Material": 8,
    "Material Key Number": 8,
    "Process (1K/2K/3K)": 6,
    "Machine/Tonnage": 6,
    "Weight": 4,
    "Part Size X": 3,
    "Part Size Y": 3,
    "Part Size Z": 3,
}

# Bonus rules: the real BOM normally has these combinations on the same header row.
BOM_COMBINATION_BONUSES = [
    ({"Part Name", "Part Number"}, 20),
    ({"Part Name", "Reference Part Number"}, 24),
    ({"Part Name", "Complexity Category"}, 14),
    ({"Part Name", "Final Decision Board"}, 14),
    ({"BOM Pos", "Part Name"}, 10),
    ({"BOM Pos", "Part Number"}, 10),
    ({"BOM Pos", "Reference Part Number"}, 12),
    ({"Complexity Category", "Final Decision Board"}, 10),
]

# Minimum confidence needed to accept a BOM table.
MIN_BOM_CONFIDENCE_SCORE = 65
BOM_SEARCH_MAX_ROWS = 180
HEADER_SEARCH_MAX_ROWS = 80
HEADER_SEARCH_MAX_COLUMNS = 30

MANDATORY_FIELDS = ["Part Number", "Material", "Final Decision Board"]

# Expected BOM columns for debug/missing-column reporting only.
# These names must match canonical keys from BOM_COLUMN_ALIASES.
EXPECTED_BOM_COLUMNS = [
    "BOM Pos",
    "Reference Part Number",
    "Part Number",
    "Part Name",
    "Complexity Category",
    "Final Decision Board",
    "Material",
    "Material Key Number",
    "Surface",
    "Metallized",
    "Special Surface",
    "Weight",
    "Machine/Tonnage",
    "Process (1K/2K/3K)",
]

# Part-number patterns and scores are configurable here, not hardcoded in utils.py.
# Higher score wins when several candidates are found in Part Name.
PART_NUMBER_PATTERNS = [
    # Continental-style PN with revision/suffix/handedness examples:
    # 290.671-01, 290.671-01/02, 290.671-01 LH, 290.671-01 RH, 290.671-01A
    {"name": "continental_dotted_revision_handed", "regex": r"\b\d{3}[.]\d{3}-\d{2}(?:/\d{2})?(?:[A-Z])?(?:\s+(?:LH|RH))?\b", "score": 100},
    {"name": "continental_dotted_optional_revision", "regex": r"\b\d{3}[.]\d{3}(?:[-/]\d{2,})?(?:/\d{2})?(?:[A-Z])?(?:\s+(?:LH|RH))?\b", "score": 90},
    {"name": "a2c_code", "regex": r"\bA2C\d{5,}[A-Z0-9-]*\b", "score": 85},
    {"name": "long_numeric_revision", "regex": r"\b\d{6,}(?:[-/]\d+)?(?:[A-Z])?(?:\s+(?:LH|RH))?\b", "score": 70},
]


OUTPUT_COLUMNS = [
    "Project", "SOP", "FOTP", "MOB Presentation Date", "Production Plant", "EOP", "Volume",
    "Part Number", "Complexity Category", "Final Decision Board", "Part Name", "Part Size X",
    "Part Size Y", "Part Size Z", "Material", "Material Key Number", "Surface", "Metallized",
    "Special Surface", "Weight", "Machine/Tonnage", "Process (1K/2K/3K)",
    "Source File", "Sheet Name", "Year", "Region", "Validation Status", "Comments",
]

REGION_TOKENS = ["EU", "EMEA", "APAC", "NSA", "NA", "SAM", "LATAM", "CHINA", "INDIA"]
