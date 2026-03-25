from __future__ import annotations

import os
import re
import time
from dataclasses import dataclass
from typing import Any, Dict, Optional

from dotenv import load_dotenv

try:
    import anthropic
except Exception as exc:  # pragma: no cover
    raise ImportError(
        "anthropic SDK is required. Install with: pip install anthropic"
    ) from exc


MODEL_NAME = "claude-sonnet-4-6"
MAX_TOKENS = 1000

SYSTEM_PROMPT = (
    "You are a senior apparel factory audit expert with 20+ years \n"
    "of experience in garment manufacturing compliance. You are \n"
    "filling out a factory audit on behalf of a mid-to-large sized, \n"
    "reasonably compliant apparel manufacturing factory. Answer all \n"
    "questions professionally using standard audit language. Always \n"
    "assume the factory has basic-to-good compliance in place unless \n"
    "the question implies otherwise."
)


@dataclass(frozen=True)
class AuditAnswer:
    score: str
    present_status: str
    improvement_plan: str

    def to_dict(self) -> Dict[str, str]:
        return {
            "score": self.score,
            "present_status": self.present_status,
            "improvement_plan": self.improvement_plan,
        }


_CLIENT: Optional["anthropic.Anthropic"] = None


def get_client() -> "anthropic.Anthropic":
    """Create (and cache) an Anthropic client.

    Loads `.env` on first call, and expects `ANTHROPIC_API_KEY`.
    """

    global _CLIENT
    if _CLIENT is not None:
        return _CLIENT

    load_dotenv(override=False)
    api_key = os.getenv("ANTHROPIC_API_KEY", "").strip()
    if not api_key:
        raise RuntimeError(
            "Missing ANTHROPIC_API_KEY. Add it to your environment or .env file."
        )

    _CLIENT = anthropic.Anthropic(api_key=api_key)
    return _CLIENT


def build_user_prompt(
    question_text: str,
    criteria_10: Optional[str] = None,
    criteria_5: Optional[str] = None,
    criteria_3: Optional[str] = None,
    criteria_0: Optional[str] = None,
) -> str:
    """Build the per-question user prompt exactly in the required structure."""

    def norm(v: Optional[str]) -> str:
        v = (v or "").strip()
        return v if v else "Not provided"

    prompt = (
        f"Audit Question: {question_text}\n\n"
        "Scoring Criteria:\n"
        f"Score 10: {norm(criteria_10)}\n"
        f"Score 5: {norm(criteria_5)}\n"
        f"Score 3: {norm(criteria_3)}\n"
        f"Score 0: {norm(criteria_0)}\n\n"
        "Based on a standard compliant apparel factory, provide:\n"
        "1. SCORE: (choose from 10, 5, 3, or 0)\n"
        "2. PRESENT STATUS: (2-3 lines, professional audit language,\n"
        "   factual and specific, sounds like filled by a real auditor)\n"
        "3. IMPROVEMENT PLAN: (1-2 lines if score is not 10, else write\n"
        "   'Maintain current standards and continue monitoring.')\n\n"
        "Return ONLY in this exact format:\n"
        "SCORE: [number]\n"
        "PRESENT STATUS: [text]\n"
        "IMPROVEMENT PLAN: [text]"
    )

    return prompt


def answer_audit_question(
    question_text: str,
    criteria: Optional[Dict[str, Any]] = None,
    *,
    delay_seconds: float = 0.5,
    max_retries: int = 2,
) -> Dict[str, str]:
    """Call Claude and return a structured audit answer.

    - Retries up to `max_retries` times after the initial attempt.
    - Adds `delay_seconds` between attempts to reduce rate-limit issues.
    """

    criteria = criteria or {}
    user_prompt = build_user_prompt(
        question_text=question_text,
        criteria_10=_string_or_none(criteria.get("10")),
        criteria_5=_string_or_none(criteria.get("5")),
        criteria_3=_string_or_none(criteria.get("3")),
        criteria_0=_string_or_none(criteria.get("0")),
    )

    last_exc: Optional[Exception] = None

    for attempt in range(max_retries + 1):
        if attempt > 0 and delay_seconds > 0:
            time.sleep(delay_seconds)

        try:
            raw_text = _call_claude(user_prompt)
            parsed = parse_claude_response(raw_text)
            return parsed.to_dict()
        except Exception as exc:
            last_exc = exc

    # If all retries fail, return a safe fallback rather than crashing the pipeline.
    return AuditAnswer(
        score="",
        present_status="",
        improvement_plan="",
    ).to_dict()


def answer_question(question_text: str, criteria: Optional[Dict[str, Any]] = None) -> Dict[str, str]:
    """Compatibility wrapper for the UI."""

    return answer_audit_question(question_text=question_text, criteria=criteria)


def _call_claude(user_prompt: str) -> str:
    client = get_client()

    resp = client.messages.create(
        model=MODEL_NAME,
        max_tokens=MAX_TOKENS,
        system=SYSTEM_PROMPT,
        messages=[{"role": "user", "content": user_prompt}],
    )

    # anthropic returns content blocks; collect all text blocks
    parts = []
    for block in getattr(resp, "content", []) or []:
        text = getattr(block, "text", None)
        if text:
            parts.append(text)

    if not parts and hasattr(resp, "content") and isinstance(resp.content, str):
        parts.append(resp.content)

    return "\n".join(parts).strip()


_SCORE_RE = re.compile(r"(?im)^\s*SCORE\s*:\s*(10|5|3|0)\s*$")
_PRESENT_RE = re.compile(r"(?is)\bPRESENT\s+STATUS\s*:\s*(.*?)\s*(?:\n\s*IMPROVEMENT\s+PLAN\s*:|\Z)")
_IMPROVE_RE = re.compile(r"(?is)\bIMPROVEMENT\s+PLAN\s*:\s*(.*?)\s*\Z")


def parse_claude_response(text: str) -> AuditAnswer:
    """Parse Claude output into structured fields.

    Expected exact format:
        SCORE: [number]
        PRESENT STATUS: [text]
        IMPROVEMENT PLAN: [text]

    Parsing is tolerant to extra whitespace/newlines.
    """

    cleaned = (text or "").strip()

    score_match = _SCORE_RE.search(cleaned)
    score = score_match.group(1).strip() if score_match else ""

    present_match = _PRESENT_RE.search(cleaned)
    present_status = _clean_multiline(present_match.group(1)) if present_match else ""

    improve_match = _IMPROVE_RE.search(cleaned)
    improvement_plan = _clean_multiline(improve_match.group(1)) if improve_match else ""

    # Fallback: line-based parsing if regex misses (e.g., model adds extra prefixes)
    if not (score and present_status and improvement_plan):
        score2, present2, improve2 = _parse_linewise(cleaned)
        score = score or score2
        present_status = present_status or present2
        improvement_plan = improvement_plan or improve2

    return AuditAnswer(
        score=score,
        present_status=present_status,
        improvement_plan=improvement_plan,
    )


def _parse_linewise(cleaned: str) -> tuple[str, str, str]:
    score = ""
    present_lines: list[str] = []
    improve_lines: list[str] = []

    mode: Optional[str] = None
    for raw_line in (cleaned or "").splitlines():
        line = raw_line.strip()
        if not line:
            continue

        up = line.upper()
        if up.startswith("SCORE:"):
            mode = "score"
            val = line.split(":", 1)[1].strip()
            m = re.match(r"^(10|5|3|0)\b", val)
            if m:
                score = m.group(1)
            continue

        if up.startswith("PRESENT STATUS:"):
            mode = "present"
            val = line.split(":", 1)[1].strip()
            if val:
                present_lines.append(val)
            continue

        if up.startswith("IMPROVEMENT PLAN:"):
            mode = "improve"
            val = line.split(":", 1)[1].strip()
            if val:
                improve_lines.append(val)
            continue

        if mode == "present":
            present_lines.append(line)
        elif mode == "improve":
            improve_lines.append(line)

    return score, _clean_multiline("\n".join(present_lines)), _clean_multiline("\n".join(improve_lines))


def _clean_multiline(value: str) -> str:
    value = (value or "").replace("\u00a0", " ")
    value = re.sub(r"\r\n?", "\n", value)
    # collapse excessive blank lines
    value = re.sub(r"\n{3,}", "\n\n", value)
    value = "\n".join(line.rstrip() for line in value.split("\n"))
    return value.strip()


def _string_or_none(value: Any) -> Optional[str]:
    if value is None:
        return None
    s = str(value).strip()
    return s if s else None
