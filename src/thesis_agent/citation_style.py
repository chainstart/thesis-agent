from __future__ import annotations

import re


def normalize_inline_citations(text: str) -> str:
    """Repair common malformed bracket citations without changing citation meaning."""
    normalized = text
    previous = None
    while previous != normalized:
        previous = normalized
        normalized = re.sub(r"\[(\d+)\s*[,，]\s*\[(\d+)\]\]", r"[\1,\2]", normalized)
        normalized = re.sub(r"\[(\d+)\s*[-－]\s*\[(\d+)\]\]", r"[\1-\2]", normalized)
        normalized = re.sub(r"\[\[(\d+)\]\]", r"[\1]", normalized)
    normalized = re.sub(r"\[(\d+)\s*[,，]\s*(\d+)\]", r"[\1,\2]", normalized)
    normalized = re.sub(r"\[(\d+)\s*[-－]\s*(\d+)\]", r"[\1-\2]", normalized)
    return normalized


def has_malformed_inline_citation(text: str) -> bool:
    return bool(
        re.search(r"\[\d+\s*[,，]\s*\[\d+\]\]", text)
        or re.search(r"\[\d+\s*[-－]\s*\[\d+\]\]", text)
        or re.search(r"\[\[\d+\]\]", text)
    )
