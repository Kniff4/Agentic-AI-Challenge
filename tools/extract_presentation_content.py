from ibm_watsonx_orchestrate.agent_builder.tools import tool
import json
import re
from typing import List, Dict


def _clean_text(text: str) -> str:
    text = text.replace("\r", "\n")
    text = re.sub(r"\n{2,}", "\n", text)
    text = re.sub(r"[ \t]+", " ", text)
    return text.strip()


def _split_sentences(text: str) -> List[str]:
    # Very lightweight sentence splitter for MVP
    parts = re.split(r'(?<=[.!?])\s+|\n+', text)
    return [p.strip(" -•\t") for p in parts if p.strip(" -•\t")]


def _extract_bullets_and_points(text: str) -> List[str]:
    lines = [line.strip() for line in text.splitlines() if line.strip()]
    bullet_like = []

    for line in lines:
        if line.startswith(("-", "•", "*")):
            bullet_like.append(line.lstrip("-•* ").strip())

    if bullet_like:
        return bullet_like

    # fallback: use sentence-like chunks
    sentences = _split_sentences(text)
    return sentences[:8]


def _guess_topic(text: str) -> str:
    lines = [line.strip() for line in text.splitlines() if line.strip()]
    if lines:
        first = lines[0]
        if len(first) <= 80:
            return first

    # fallback: first sentence truncated
    sentences = _split_sentences(text)
    if sentences:
        return sentences[0][:80]

    return "Presentation Draft"


def _guess_audience(text: str) -> str:
    t = text.lower()
    if any(word in t for word in ["leadership", "executive", "board", "management"]):
        return "leadership"
    if any(word in t for word in ["client", "customer", "external", "prospect"]):
        return "client"
    if any(word in t for word in ["team", "internal", "staff", "employee"]):
        return "internal"
    return "general"


def _guess_purpose(text: str) -> str:
    t = text.lower()
    if any(word in t for word in ["proposal", "recommend", "pilot", "rollout"]):
        return "proposal"
    if any(word in t for word in ["strategy", "strategic", "roadmap"]):
        return "strategy"
    if any(word in t for word in ["summary", "overview", "update"]):
        return "informational"
    return "informational"


def _deduplicate_preserve_order(items: List[str]) -> List[str]:
    seen = set()
    result = []
    for item in items:
        norm = item.strip().lower()
        if norm and norm not in seen:
            seen.add(norm)
            result.append(item.strip())
    return result


def _extract_key_points(text: str, max_points: int = 6) -> List[str]:
    raw_points = _extract_bullets_and_points(text)
    cleaned = []

    for point in raw_points:
        point = re.sub(r"\s+", " ", point).strip()
        if not point:
            continue
        if len(point) > 140:
            point = point[:137].rstrip() + "..."
        cleaned.append(point)

    cleaned = _deduplicate_preserve_order(cleaned)
    return cleaned[:max_points]


def _recommend_sections(audience: str, purpose: str, key_points: List[str]) -> List[str]:
    if audience == "leadership":
        return [
            "Executive Summary",
            "Business Impact",
            "Recommended Approach",
            "Next Steps"
        ]
    if purpose == "proposal":
        return [
            "Overview",
            "Opportunity",
            "Recommendation",
            "Next Steps"
        ]
    return [
        "Summary",
        "Main Points",
        "Impact",
        "Next Steps"
    ]


def _recommend_slide_count(audience: str, key_points: List[str]) -> int:
    if audience == "leadership":
        return 5
    if len(key_points) >= 6:
        return 6
    return 4


@tool
def extract_presentation_content(source_text: str) -> str:
    """
    Extract a lightweight structured presentation brief from raw source text.

    Args:
        source_text (str): Raw user text or text extracted from a document.

    Returns:
        str: JSON string with structured presentation planning information.
    """
    text = _clean_text(source_text)
    topic = _guess_topic(text)
    audience = _guess_audience(text)
    purpose = _guess_purpose(text)
    key_points = _extract_key_points(text, max_points=6)
    recommended_sections = _recommend_sections(audience, purpose, key_points)
    recommended_slide_count = _recommend_slide_count(audience, key_points)

    result: Dict[str, object] = {
        "topic": topic,
        "audience": audience,
        "purpose": purpose,
        "recommended_slide_count": recommended_slide_count,
        "key_points": key_points,
        "recommended_sections": recommended_sections,
        "source_summary": text[:2500]
    }

    return json.dumps(result)