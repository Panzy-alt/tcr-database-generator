"""Utility helpers for robust matching, validation and part number cleanup."""

from __future__ import annotations

import re
from pathlib import Path
from typing import Any

from .config import PART_NUMBER_PATTERNS, REGION_TOKENS


def display_text(value: Any) -> str:
    """Convert a cell value to a clean display string without altering meaning."""
    if value is None:
        return ""
    return str(value).strip()


def normalize_text(value: Any) -> str:
    """Human-readable normalization: lower case and collapse whitespace."""
    text = display_text(value).replace("\n", " ").replace("\r", " ").lower()
    return re.sub(r"\s+", " ", text).strip()


def normalize_key(value: Any) -> str:
    """Strict matching key that ignores case, spaces and punctuation."""
    text = normalize_text(value)
    return re.sub(r"[^a-z0-9]+", "", text)


def is_empty(value: Any) -> bool:
    return value is None or str(value).strip() == ""


def first_non_empty(*values: Any) -> Any:
    for value in values:
        if not is_empty(value):
            return value
    return ""


def extract_year_from_path(path: str | Path) -> str:
    match = re.search(r"\b(20\d{2})\b", str(path))
    return match.group(1) if match else ""


def extract_region_from_path(path: str | Path) -> str:
    upper_path = str(path).upper()
    for token in REGION_TOKENS:
        if re.search(rf"\b{re.escape(token.upper())}\b", upper_path):
            return token.upper()
    return ""


def extract_part_number_from_text(text: Any) -> tuple[str, str]:
    """Extract the best matching part number and remove only that exact match.

    If several candidates exist, the match with the highest configured score wins.
    Ties are resolved by longer match length, then earlier text position. Other
    text is never deleted.
    """
    original = display_text(text)
    if not original:
        return "", ""

    candidates: list[tuple[int, int, int, int, str]] = []
    for item in PART_NUMBER_PATTERNS:
        regex = item["regex"]
        score = int(item.get("score", 0))
        for match in re.finditer(regex, original, flags=re.IGNORECASE):
            value = match.group(0).strip()
            candidates.append((score, len(value), -match.start(), match.end(), value))

    if not candidates:
        return "", original

    # best tuple = highest score, longest candidate, earliest occurrence
    score, length, neg_start, end, value = max(candidates)
    start = -neg_start
    cleaned = original[:start] + original[end:]
    cleaned = re.sub(r"\s{2,}", " ", cleaned).strip(" -_/.,;:")
    return value, cleaned


def clean_part_number(value: Any) -> str:
    """Preserve the visible part number exactly as text."""
    return display_text(value)


def build_alias_lookup(alias_map: dict[str, list[str]]) -> dict[str, str]:
    """Build normalized alias -> canonical name lookup."""
    lookup: dict[str, str] = {}
    for canonical, aliases in alias_map.items():
        lookup[normalize_key(canonical)] = canonical
        for alias in aliases:
            lookup[normalize_key(alias)] = canonical
    return lookup
