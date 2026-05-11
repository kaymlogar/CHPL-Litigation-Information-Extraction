import csv
import json
import os
import re
import sys
import time
from collections import defaultdict
from concurrent.futures import ThreadPoolExecutor
from datetime import date, datetime
from pathlib import Path
from typing import Any

import fitz  # PyMuPDF
import pymupdf4llm
import titlecase
from openpyxl import Workbook

try:
    from openai import OpenAI as _OpenAI
    _OPENAI_AVAILABLE = True
except ImportError:
    _OPENAI_AVAILABLE = False

try:
    from pypdf import PdfReader as _PdfReader
    _PYPDF_AVAILABLE = True
except ImportError:
    _PYPDF_AVAILABLE = False

OPENROUTER_MODEL = "google/gemini-2.5-flash"  # change to any OpenRouter model ID

LLM_MAX_CHARS_ANALYSIS = 15_000   # excerpt length for Potential Impact / Why This Matters
LLM_MAX_CHARS_ISSUES   = 25_000   # claims-section cap for the LLM issues fallback

DOCKET_FOLDER_NAME = "TrialCourtDockets"
COMPLAINT_FOLDER_NAME = "TrialCourtComplaints"
DISTRICT_COURT_CODE_FILE = "28 USC Ch5 District Courts.pdf"
JUDGE_CSV_FILE = "Federal Judicial Center Export.csv"
GOALS_MAPPING_CSV_FILE = "GoalsMapping.csv"
GOALS_EXAMPLES_FILE = "GoalsExamples.csv"
ISSUES_CSV_FILE = "Issues.csv"
ISSUES_MAPPING_CSV_FILE = "IssuesMapping.csv"
LEGAL_ISSUES_EXAMPLES_FILE = "LegalIssuesExamples.csv"
ANALYSIS_EXAMPLES_FILE = "AnalysisExamples.csv"
US_STATES = [
        "ALABAMA", "ALASKA", "ARIZONA", "ARKANSAS", "CALIFORNIA",
        "COLORADO", "CONNECTICUT", "DELAWARE", "DISTRICT OF COLUMBIA", "FLORIDA", "GEORGIA",
        "HAWAII", "IDAHO", "ILLINOIS", "INDIANA", "IOWA",
        "KANSAS", "KENTUCKY", "LOUISIANA", "MAINE", "MARYLAND",
        "MASSACHUSETTS", "MICHIGAN", "MINNESOTA", "MISSISSIPPI",
        "MISSOURI", "MONTANA", "NEBRASKA", "NEVADA", "NEW HAMPSHIRE",
        "NEW JERSEY", "NEW MEXICO", "NEW YORK", "NORTH CAROLINA",
        "NORTH DAKOTA", "OHIO", "OKLAHOMA", "OREGON", "PENNSYLVANIA", "PUERTO RICO",
        "RHODE ISLAND", "SOUTH CAROLINA", "SOUTH DAKOTA", "TENNESSEE",
        "TEXAS", "UTAH", "VERMONT", "VIRGINIA", "WASHINGTON",
        "WEST VIRGINIA", "WISCONSIN", "WYOMING" 
    ]

DISTRICTS = {
    "ALABAMA": {"NORTHERN", "MIDDLE", "SOUTHERN"},
    "ARKANSAS": {"EASTERN", "WESTERN"},
    "CALIFORNIA": {"NORTHERN", "EASTERN", "CENTRAL", "SOUTHERN"},
    "FLORIDA": {"NORTHERN", "MIDDLE", "SOUTHERN"},
    "GEORGIA": {"NORTHERN", "MIDDLE", "SOUTHERN"},
    "ILLINOIS": {"NORTHERN", "CENTRAL", "SOUTHERN"},
    "INDIANA": {"NORTHERN", "SOUTHERN"},
    "IOWA": {"NORTHERN", "SOUTHERN"},
    "KENTUCKY": {"EASTERN", "WESTERN"},
    "LOUISIANA": {"EASTERN", "MIDDLE", "WESTERN"},
    "MICHIGAN": {"EASTERN", "WESTERN"},
    "MISSISSIPPI": {"NORTHERN", "SOUTHERN"},
    "MISSOURI": {"EASTERN", "WESTERN"},
    "NEW YORK": {"NORTHERN", "SOUTHERN", "EASTERN", "WESTERN"},
    "NORTH CAROLINA": {"EASTERN", "MIDDLE", "WESTERN"},
    "OHIO": {"NORTHERN", "SOUTHERN"},
    "OKLAHOMA": {"NORTHERN", "EASTERN", "WESTERN"},
    "PENNSYLVANIA": {"EASTERN", "MIDDLE", "WESTERN"},
    "TENNESSEE": {"EASTERN", "MIDDLE", "WESTERN"},
    "TEXAS": {"NORTHERN", "SOUTHERN", "EASTERN", "WESTERN"},
    "VIRGINIA": {"EASTERN", "WESTERN"},
    "WASHINGTON": {"EASTERN", "WESTERN"},
    "WEST VIRGINIA": {"NORTHERN", "SOUTHERN"},
    "WISCONSIN": {"EASTERN", "WESTERN"},
}

POTENTIAL_AGENCY_PARTIES = {
    "Department of Health and Human Services", "Centers for Medicare & Medicaid Services",
    "Food and Drug Administration", "National Institutes of Health",
    "National Science Foundation", "Department of Labor",
    "Department of Justice", "Department of Treasury", "Department of the Treasury",
    "Consumer Financial Protection Bureau", "Federal Trade Commission",
}

GOV_KEYWORDS = POTENTIAL_AGENCY_PARTIES | {"Attorney General"}

AGENCY_ABBREVIATIONS = {
    "HHS": "Department of Health and Human Services",
    "H.H.S.": "Department of Health and Human Services",
    "CMS": "Centers for Medicare & Medicaid Services",
    "C.M.S.": "Centers for Medicare & Medicaid Services",
    "FDA": "Food and Drug Administration",
    "F.D.A.": "Food and Drug Administration",
    "NIH": "National Institutes of Health",
    "N.I.H.": "National Institutes of Health",
    "NSF": "National Science Foundation",
    "N.S.F.": "National Science Foundation",
    "CFPB": "Consumer Financial Protection Bureau",
    "C.F.P.B.": "Consumer Financial Protection Bureau",
    "FTC": "Federal Trade Commission",
    "F.T.C.": "Federal Trade Commission",
}

PRE_HONORIFICS = {
    "DR.", "DR",
    "Atty General", "Attorney General", "Atty Gen", "AG",
    "AL", "AK", "AZ", "AR", "CA", "CO", "CT", "DE", "FL", "GA",
    "HI", "ID", "IL", "IN", "IA", "KS", "KY", "LA", "ME", "MD",
    "MA", "MI", "MN", "MS", "MO", "MT", "NE", "NV", "NH", "NJ",
    "NM", "NY", "NC", "ND", "OH", "OK", "OR", "PA", "RI", "SC",
    "SD", "TN", "TX", "UT", "VT", "VA", "WA", "WV", "WI", "WY", "DC", "PR",
}

POST_HONORIFICS = {
    "JD", "J.D.", "DR", "MD", "M.D.", "PHD", "D.O.", "DO",
    "MBA", "M.B.A.", "PA-C", "FNP-BC", "MHPE",
    "Atty General", "Attorney General", "Atty Gen", "AG",
    "Atty General of", "Attorney General of", "Atty Gen of", "AG of",
    "AL", "AK", "AZ", "AR", "CA", "CO", "CT", "DE", "FL", "GA",
    "HI", "ID", "IL", "IN", "IA", "KS", "KY", "LA", "ME", "MD",
    "MA", "MI", "MN", "MS", "MO", "MT", "NE", "NV", "NH", "NJ",
    "NM", "NY", "NC", "ND", "OH", "OK", "OR", "PA", "RI", "SC",
    "SD", "TN", "TX", "UT", "VT", "VA", "WA", "WV", "WI", "WY", "DC", "PR",
}

RELIEF_SIGNAL = "wherefore"

_STATE_NAME_ALT_RE = "|".join(
    re.escape(state) for state in sorted(US_STATES, key=len, reverse=True)
)
_AG_OF_STATE_BOUNDARY_BODY = (
    rf"\b(?:attorney\s+general|atty\s+general|atty\s+gen|ag)\s+of\s+"
    rf"(?:the\s+)?(?:{_STATE_NAME_ALT_RE})\b"
)
_AG_OF_STATE_BOUNDARY_RE = re.compile(
    _AG_OF_STATE_BOUNDARY_BODY,
    flags=re.IGNORECASE,
)

# Pre-compiled patterns used in multiple functions — defined once at module level.
_AGENCY_PATTERNS: list[re.Pattern] = [
    re.compile(p) for p in [
        r"\bdepartment\s+of\b",
        r"\bsecretary\s+of\b",
        r"\bunited\s+states\b",
        r"\bu\.s\.\b",
        r"\boffice\s+of\b",
        r"\badministration\b",
        r"\bagency\b",
        r"\bbureau\s+of\b",
        r"\bcommission\b",
        r"\battorney\s+general\b",
        r"\bgovernor\b",
        r"\bstate\s+of\b",
        r"\bcommonwealth\s+of\b",
    ]
]

_NOTICE_RE = re.compile(
    r"(?i)\b(pro\s+hac\s+vice|attorney\s+to\s+be\s+noticed|lead\s+attorney)\b"
)

_TERMINATED_MARKER_RE = re.compile(
    r"(?ix)\b(?:terminated|t\s*e\s*r\s*m\s*i\s*n\s*a\s*t\s*e\s*d)\b\s*:?"
)


_GOAL_OUTPUT_ORDER = [
    "Block anticompetitive practices",
    "Block enforcement of an agency action",
    "Block enforcement of an executive order",
    "Block infringement of a patent",
    "Block operation of a law",
    "Block operation of a program",
    "Block enforcement of a law",
    "Block defendant action",
    "Compel agency action",
    "Compel defendant action",
    "Appoint Independent Fiduciary",
    "Appoint court monitor",
    "Certify class",
    "Intervene in the litigation",
    "Award damages",
    "Award restitution",
    "Declaration that agency action is unlawful",
    "Declaration that executive order is unlawful",
    "Declaration that a program is unlawful",
    "Declaration that a law is unlawful",
    "Declaration that defendant breached fiduciary duty",
    "Declaration that defendant action is unlawful",
    "Vacate agency action",
    "Vacate arbitration awards",
]


def _normalize_honorific_phrase(value: str) -> str:
    """Normalize honorific phrases for case-insensitive boundary matching."""
    s = re.sub(r"[.]", "", (value or "").upper())
    return re.sub(r"\s+", " ", s).strip()


_PRE_HONORIFICS_NORMALIZED = {
    _normalize_honorific_phrase(tok) for tok in PRE_HONORIFICS if tok
}
_POST_HONORIFICS_NORMALIZED = {
    _normalize_honorific_phrase(tok) for tok in POST_HONORIFICS if tok
}
_PRE_HON_WORD_COUNTS = sorted(
    {len(tok.split()) for tok in _PRE_HONORIFICS_NORMALIZED if tok},
    reverse=True,
)
_POST_HON_WORD_COUNTS = sorted(
    {len(tok.split()) for tok in _POST_HONORIFICS_NORMALIZED if tok},
    reverse=True,
)


def _get_openrouter_client() -> "_OpenAI | None":
    """Return an OpenRouter-backed OpenAI client, or None if unavailable."""
    if not _OPENAI_AVAILABLE:
        return None
    api_key = os.environ.get("OPENROUTER_API_KEY")
    if not api_key:
        return None
    return _OpenAI(
        base_url="https://openrouter.ai/api/v1",
        api_key=api_key,
    )


def _call_llm(llm_client: Any, messages: list[dict], max_tokens: int, temperature: float = 0) -> str | None:
    """Call the LLM with exponential backoff on rate-limit errors.

    Returns the response text on success, or None if all attempts fail.
    Retries up to 5 times, waiting 5/10/20/40 seconds between attempts on 429.
    """
    for attempt in range(5):
        try:
            response = llm_client.chat.completions.create(
                model=OPENROUTER_MODEL,
                messages=messages,
                max_tokens=max_tokens,
                temperature=temperature,
            )
            return response.choices[0].message.content.strip()
        except Exception as e:
            if "429" in str(e) and attempt < 4:
                wait = 2 ** attempt * 5
                print(f"OpenRouter rate limit hit, retrying in {wait}s... ({e})", file=sys.stderr)
                time.sleep(wait)
            else:
                print(f"OpenRouter API error: {e}", file=sys.stderr)
                return None
    return None


def _load_csv_rows(path: Path) -> list[tuple]:
    """Load all data rows from a CSV file, skipping the header row.

    Returns a list of row tuples, or an empty list if the file cannot be read.
    """
    try:
        with path.open(newline="", encoding="utf-8-sig") as f:
            reader = csv.reader(f)
            rows = list(reader)
        return [tuple(row) for row in rows[1:]]  # skip header
    except Exception as e:
        print(f"Warning: could not load {path}: {e}", file=sys.stderr)
        return []


def _load_goals_examples(script_dir: Path) -> list[tuple[str, str]]:
    """Load request-to-goal pairs from GoalsExamples.xlsx.

    The file has three columns: Case, Request, Goal.
    Returns a list of (request_text, goal_label) tuples.
    Rows with no Goal label are skipped.
    """
    results = []
    for row in _load_csv_rows(script_dir / GOALS_EXAMPLES_FILE):
        if len(row) < 3:
            continue
        request, goal = row[1], row[2]
        if not request or not goal:
            continue
        goal_clean = str(goal).strip().rstrip(";").strip()
        if goal_clean:
            results.append((str(request).strip(), goal_clean))
    return results


def _build_llm_goals_context(
    examples: list[tuple[str, str]],
    mapping_dict: dict[str, str],
) -> str:
    """Build the few-shot and phrase-indicator block injected into the classify prompt."""
    lines: list[str] = []

    # --- Section 1: phrase indicators from GoalsMapping.csv grouped by goal ---
    if mapping_dict:
        by_goal: dict[str, list[str]] = defaultdict(list)
        for phrase, goal in mapping_dict.items():
            by_goal[goal.strip()].append(phrase.strip())

        lines.append(
            "PHRASE INDICATORS — when the prayer for relief contains language "
            "like the phrases below, apply the corresponding goal label:"
        )
        for goal in _GOAL_OUTPUT_ORDER:
            phrases = by_goal.get(goal, [])
            if phrases:
                formatted = "; ".join(f'"{p}"' for p in phrases)
                lines.append(f'  "{goal}": {formatted}')
        lines.append("")

    # --- Section 2: request-level examples from GoalsExamples.xlsx ---
    # Each example maps a specific prayer-for-relief sentence to its goal label.
    if examples:
        lines.append(
            "REQUEST-TO-GOAL EXAMPLES — specific prayer language mapped to the "
            "correct goal (classify the new text below, not these examples):"
        )
        for request_text, goal_label in examples:
            display = (
                request_text if len(request_text) <= 200
                else request_text[:197] + "..."
            )
            lines.append(f'  Request: "{display}"')
            lines.append(f"  Goal: {goal_label}")
            lines.append("")

    return "\n".join(lines)


def _is_agency_defendant(defendants_str: str) -> bool:
    """Return True if any defendant appears to be a government agency or official.

    Checks known agency names and broad structural patterns so it generalises
    beyond the hardcoded POTENTIAL_AGENCY_PARTIES list.
    """
    if not defendants_str:
        return False
    text = defendants_str.lower()

    for name in GOV_KEYWORDS:
        if name.lower() in text:
            return True

    return any(pat.search(text) for pat in _AGENCY_PATTERNS)


def _load_issues(script_dir: Path) -> list[str]:
    """Load the canonical legal issues list from Issues.csv, normalized."""
    path = script_dir / ISSUES_CSV_FILE
    issues: list[str] = []
    try:
        with path.open("r", encoding="utf-8-sig", newline="") as f:
            reader = csv.reader(f)
            next(reader, None)  # skip header
            for row in reader:
                if row:
                    label = row[0].replace(" ", " ").strip()
                    # Skip blank and documentation/example rows
                    if label and not label.lower().startswith(("e.g.", "example", "note:")):
                        issues.append(label)
    except FileNotFoundError:
        print(f"Warning: issues file not found: {path}", file=sys.stderr)
    return issues


def _load_legal_issues_examples(script_dir: Path) -> list[tuple[str, str]]:
    """Load labeled case→legal-issues pairs from LegalIssuesExamples.xlsx."""
    return [
        (str(row[0]).strip(), str(row[1]).strip())
        for row in _load_csv_rows(script_dir / LEGAL_ISSUES_EXAMPLES_FILE)
        if len(row) >= 2 and row[0] and row[1]
    ]


def _build_llm_issues_context(examples: list[tuple[str, str]]) -> str:
    """Build the few-shot context block injected into the issues classification prompt."""
    if not examples:
        return ""
    lines = [
        "LABELED EXAMPLES — correct legal issue classifications for reference "
        "(use only to understand patterns; return only labels from the list above):"
    ]
    for case_name, case_issues in examples:
        lines.append(f"  Case: {case_name}")
        lines.append(f"  Issues: {case_issues}")
        lines.append("")
    return "\n".join(lines)


def _load_analysis_examples(script_dir: Path) -> list[tuple[str, str, str]]:
    """Load AnalysisExamples.xlsx and return (case_name, potential_impact, why_this_matters) tuples."""
    examples: list[tuple[str, str, str]] = []
    for row in _load_csv_rows(script_dir / ANALYSIS_EXAMPLES_FILE):
        if not row or len(row) < 3:
            continue
        case   = str(row[0]).strip() if row[0] else ""
        impact = str(row[1]).strip() if row[1] else ""
        why    = str(row[2]).strip() if row[2] else ""
        if case and (impact or why):
            examples.append((case, impact, why))
    return examples


def _build_llm_impact_context(examples: list[tuple[str, str, str]]) -> str:
    """Build the few-shot context block for the potential impact prompt."""
    if not examples:
        return ""
    lines = [
        "EXAMPLES — correct Potential Impact sentences for reference "
        "(use only to understand the expected style and scope):"
    ]
    for case_name, impact, _ in examples:
        if not impact:
            continue
        lines.append(f"  Case: {case_name}")
        lines.append(f"  Potential Impact: {impact}")
        lines.append("")
    return "\n".join(lines)


def _build_llm_why_context(examples: list[tuple[str, str, str]]) -> str:
    """Build the few-shot context block for the why-this-matters prompt."""
    if not examples:
        return ""
    lines = [
        "EXAMPLES — correct Why This Matters responses for reference "
        "(use only to understand the expected style and scope):"
    ]
    for case_name, _, why in examples:
        if not why:
            continue
        lines.append(f"  Case: {case_name}")
        lines.append(f"  Why This Matters: {why}")
        lines.append("")
    return "\n".join(lines)


def _compile_issues_patterns(
    issues_mapping: dict[str, str],
    valid_issues: set[str],
) -> list[tuple[re.Pattern | None, str, str]]:
    """Pre-compile match patterns for every entry in issues_mapping.

    Returns a list of (compiled_pattern_or_None, key_norm, issue) tuples.
    Entries with compiled_pattern=None use fast substring matching instead of regex.
    Call once at startup and pass the result to _classify_issues_regex.
    """
    patterns: list[tuple[re.Pattern | None, str, str]] = []
    for key_text, issue in issues_mapping.items():
        if issue not in valid_issues:
            continue
        key_norm = re.sub(r"[ \t\r\n\xa0]+", " ", key_text).lower().strip()
        if not key_norm:
            continue
        if len(key_norm) <= 6 and re.fullmatch(r"[a-z]+", key_norm):
            compiled = re.compile(r"\b" + re.escape(key_norm) + r"\b")
        else:
            compiled = None
        patterns.append((compiled, key_norm, issue))
    return patterns


def _classify_issues_regex(
    claims_text: str,
    issues_patterns: list[tuple[re.Pattern | None, str, str]],
) -> set[str]:
    """Return the set of canonical issue labels found via case-insensitive phrase matching.

    Each key phrase in ``issues_patterns`` (built by _compile_issues_patterns) is matched
    against the claims text. Short acronyms use pre-compiled word-boundary regex;
    longer phrases use plain substring matching.
    PDF whitespace and pipe characters (table artifacts) are normalised before comparison.
    """
    if not claims_text or not issues_patterns:
        return set()

    normalized = re.sub(r"[ \t\r\n\xa0|]+", " ", claims_text).lower()
    normalized = normalized.replace(" & ", " and ")

    found: set[str] = set()
    for compiled, key_norm, issue in issues_patterns:
        if compiled is not None:
            if compiled.search(normalized):
                found.add(issue)
        else:
            if key_norm in normalized:
                found.add(issue)

    return found


def strip_party_honorifics(text: str) -> str:
    """
    Remove honorifics only at boundaries:
    - PRE_HONORIFICS at the beginning
    - POST_HONORIFICS at the end
    """
    if not text:
        return text

    s_clean = re.sub(r"\s+", " ", text).strip()
    if not s_clean:
        return ""

    # Targeted boundary strip for "Attorney General of <state>" phrase.
    s_clean = re.sub(
        rf"^\s*{_AG_OF_STATE_BOUNDARY_BODY}\s*",
        "",
        s_clean,
        flags=re.IGNORECASE,
    ).strip()
    s_clean = re.sub(
        rf"\s*{_AG_OF_STATE_BOUNDARY_BODY}\s*$",
        "",
        s_clean,
        flags=re.IGNORECASE,
    ).strip()

    words = s_clean.split(" ")
    if not words:
        return ""

    # Preserve sovereign-party names like "State of Colorado"/"Commonwealth of Massachusetts".
    if re.match(r"(?i)^(state|commonwealth)\s+of\s+", " ".join(words)):
        return re.sub(r"\s+", " ", " ".join(words)).strip()

    while words:
        removed = False
        for n in _PRE_HON_WORD_COUNTS:
            if len(words) < n:
                continue
            candidate = _normalize_honorific_phrase(" ".join(words[:n]))
            if candidate in _PRE_HONORIFICS_NORMALIZED:
                words = words[n:]
                removed = True
                break
        if not removed:
            break

    # Minimal guard: do not strip trailing honorifics from an existing '... et al.'
    # tail, which can otherwise turn it into '... et' and later produce 'Et et al.'.
    if re.search(r"(?i)\bet\s+al(?:\.+)?\s*$", " ".join(words)):
        return re.sub(r"\s+", " ", " ".join(words)).strip()

    while words:
        removed = False
        for n in _POST_HON_WORD_COUNTS:
            if len(words) < n:
                continue
            candidate = _normalize_honorific_phrase(" ".join(words[-n:]))
            if candidate in _POST_HONORIFICS_NORMALIZED:
                words = words[:-n]
                removed = True
                break
        if not removed:
            break

    return re.sub(r"\s+", " ", " ".join(words)).strip()


def strip_case_name_honorifics(case_name: str) -> str:
    """Apply boundary-only honorific stripping to each side of a case caption."""
    parts = _CASE_NAME_VS_SPLIT.split(case_name, maxsplit=1)
    if len(parts) < 2:
        return strip_party_honorifics(case_name)
    left = strip_party_honorifics(parts[0].strip())
    right = strip_party_honorifics(parts[1].strip())
    return f"{left} v. {right}"


SECTION_NUMBERS = ["81", "81A"] + [str(i) for i in range(82, 133)]

JUDGE_TITLE_KEYWORDS = [
    "HONORABLE", "JUDGE", "CHIEF", "THE", "DISTRICT",
    "US", "U.S.", "MAGISTRATE", "SENIOR", "CIRCUIT", "BANKRUPTCY",
]

GOVERNMENT_PARTY_KEYWORD = "OFFICIAL CAPACITY"

DEFENDANTS_PARSE_FAILURE_MESSAGE = (
    "No defendants identified - check docket for more information."
)

BLOOMBERG_COPYRIGHT_STRING = "© 2026 Bloomberg Industry Group, Inc. All Rights Reserved. Terms of Service"

SPELL_CHECK = {"LLLP": "LLP"}

# Kept all-caps in case titles. identify_case_name lowercases first, then titlecase(),
# which would otherwise turn e.g. "ACA" into "Aca"; the titlecase `callback` fixes that.
# Add acronyms as they appear in docket titles. (Short tokens can theoretically
# match a rare personal name—curate if that comes up.)
CASE_NAME_ACRONYMS = frozenset({
    "ACA",
    "PPACA",
    "ADEA",
    "ADA",
    "AIA",
    "CFPB",
    "COBRA",
    "ERISA",
    "FCC",
    "FDA",
    "FLSA",
    "FMLA",
    "FTC",
    "HHS",
    "HIPAA",
    "IRC",
    "IRS",
    "MSPB",
    "NLRA",
    "NLRB",
    "OSHA",
    "SEC",
})


def case_name_acronym_callback(word: str, all_caps: bool = False) -> str | None:
    if not word:
        return None
    core = word.rstrip(".,;:")
    trail = word[len(core) :]
    key = core.upper()
    if key in CASE_NAME_ACRONYMS:
        return key + trail
    return None


def _collapse_ws(s: str) -> str:
    return re.sub(r"\s+", " ", s.strip())


def expand_agency_abbreviations_in_case_name(case_name: str) -> str:
    """Replace AGENCY_ABBREVIATIONS keys with their full agency names (longest keys first)."""
    if not case_name:
        return case_name
    result = case_name
    for abbr in sorted(AGENCY_ABBREVIATIONS.keys(), key=len, reverse=True):
        pat = r"(?<!\w)" + re.escape(abbr) + r"(?!\w)"
        result = re.sub(pat, AGENCY_ABBREVIATIONS[abbr], result, flags=re.IGNORECASE)
    return result


_CASE_NAME_VS_SPLIT = re.compile(r"\s+v\.\s+", re.IGNORECASE)


def _agency_name_prefixes_longest_first() -> tuple[str, ...]:
    return tuple(
        sorted(
            (
                re.sub(r"^the\s+", "", _collapse_ws(a).lower())
                for a in POTENTIAL_AGENCY_PARTIES
            ),
            key=len,
            reverse=True,
        )
    )


def is_united_states_sole_party(side: str) -> bool:
    """True when this caption side names only the sovereign U.S., not an executive agency."""
    s = _collapse_ws(side).strip(".,;: ")
    sl = re.sub(r"^the\s+", "", s.lower())
    return bool(
        re.fullmatch(r"united\s+states(?:\s+of\s+america)?", sl)
        or re.fullmatch(r"u\.?\s*s\.?\s*a\.?", sl)
        or re.fullmatch(r"u\.?\s*s\.?", sl)
        or sl == "us"
    )


def _remainder_starts_with_listed_agency(remainder: str) -> bool:
    """True if remainder begins with a POTENTIAL_AGENCY_PARTIES name (allows leading 'the')."""
    r = _collapse_ws(remainder).lower()
    r = re.sub(r"^the\s+", "", r)
    return any(r.startswith(pfx) for pfx in _agency_name_prefixes_longest_first())


def _side_contains_listed_agency_substring(side: str) -> bool:
    n = _collapse_ws(side).lower()
    return any(_collapse_ws(a).lower() in n for a in POTENTIAL_AGENCY_PARTIES)


def _strip_leading_us_from_agency_party_side(side: str) -> str:
    """
    For a party that names a listed agency, drop a leading United States / U.S.A. / U.S.
    wrapper only when what follows immediately begins with that agency's name.

    Does not alter a side that is solely United States / U.S. Omits stripping when the
    country prefix does not directly precede an agency caption (e.g. relator suits).
    """
    if not side or not side.strip():
        return side
    s = side.strip()
    if is_united_states_sole_party(s):
        return s
    if not _side_contains_listed_agency_substring(s):
        return s

    prefixes = (
        r"(?:the\s+)?united\s+states\s+of\s+america\s+",
        r"(?:the\s+)?united\s+states\s+",
        r"(?:the\s+)?u\.?\s*s\.?\s*a\.?\s+",
        r"(?:the\s+)?u\.?\s*s\.?\s+",
    )
    for pref in prefixes:
        m = re.compile(r"^" + pref, re.IGNORECASE).match(s)
        if not m:
            continue
        remainder = s[m.end() :].strip()
        if not remainder:
            continue
        if _remainder_starts_with_listed_agency(remainder):
            return remainder
    return s


def normalize_case_name_us_agencies(case_name: str) -> str:
    """Each side of 'v.': strip leading United States/U.S. only for listed agency parties."""
    parts = _CASE_NAME_VS_SPLIT.split(case_name, maxsplit=1)
    if len(parts) < 2:
        return case_name
    left, right = parts[0].strip(), parts[1].strip()
    left = _strip_leading_us_from_agency_party_side(left)
    right = _strip_leading_us_from_agency_party_side(right)
    return f"{left} v. {right}"


# Match one-or-more trailing "et al" / "et al." chunks (optionally comma-separated).
_ET_AL_TRAILING = re.compile(
    r"(?:(?:,\s*)?\bet\s+al(?:\.+)?\s*)+$",
    re.IGNORECASE,
)


def _strip_trailing_article_the_from_party_side(side: str) -> str:
    """Drop a trailing standalone 'the' (e.g. 'Estate of X the et al.' -> 'Estate of X et al.')."""
    s = (side or "").strip()
    if not s:
        return s
    # '… the et al…' before et al is stripped elsewhere
    s = re.sub(
        r"\s+\bthe\s+(?=\bet\s+al(?:\.+)?\b)",
        " ",
        s,
        flags=re.IGNORECASE,
    )
    s = re.sub(r"\s+\bthe\s*$", "", s, flags=re.IGNORECASE).strip()
    return re.sub(r"\s+", " ", s).strip()


def _move_trailing_the_to_front(name: str) -> str:
    """Move a trailing article 'the' to the front of a party name, capitalized.

    e.g. 'Estate of Gene B. Lokken the' -> 'The Estate of Gene B. Lokken'
    Does nothing if no trailing 'the' is present, or if the name already
    starts with 'The'.
    """
    m = re.search(r"\s+\bthe\s*$", name, flags=re.IGNORECASE)
    if not m:
        return name
    core = name[: m.start()].strip()
    if re.match(r"(?i)^the\b", core):
        return core
    return "The " + core


def _strip_trailing_the_each_case_name_side(case_name: str) -> str:
    """Apply _strip_trailing_article_the_from_party_side to each side of ' v. '."""
    parts = _CASE_NAME_VS_SPLIT.split(case_name, maxsplit=1)
    if len(parts) < 2:
        return _strip_trailing_article_the_from_party_side(case_name)
    left, right = parts[0].strip(), parts[1].strip()
    return (
        f"{_strip_trailing_article_the_from_party_side(left)} v. "
        f"{_strip_trailing_article_the_from_party_side(right)}"
    )


def count_named_parties(plaintiffs_or_defendants_field: str) -> int:
    """Count parties in a '; '-joined field from identify_plaintiffs / identify_defendants."""
    if (
        not plaintiffs_or_defendants_field
        or plaintiffs_or_defendants_field == "None"
        or plaintiffs_or_defendants_field == DEFENDANTS_PARSE_FAILURE_MESSAGE
    ):
        return 0
    return len([p for p in plaintiffs_or_defendants_field.split(";") if p.strip()])


def _ensure_side_ends_with_et_al(side: str) -> str:
    """Normalize trailing comma-style et al., then end with a single ' et al.'."""
    base = side.strip()
    base = _ET_AL_TRAILING.sub("", base).rstrip().rstrip(",").strip()
    # If stripping missed a lone stray period (e.g. odd OCR), drop only repeated dots.
    base = re.sub(r"\.{2,}\s*$", "", base).rstrip().strip()
    base = _strip_trailing_article_the_from_party_side(base)
    return f"{base} et al."


def apply_et_al_to_case_name(
    case_name: str, num_plaintiffs: int, num_defendants: int
) -> str:
    """When multiple plaintiffs or defendants exist, ensure each caption side ends with 'et al.'."""
    if num_plaintiffs < 2 and num_defendants < 2:
        return _strip_trailing_the_each_case_name_side(case_name)

    parts = _CASE_NAME_VS_SPLIT.split(case_name, maxsplit=1)
    if len(parts) < 2:
        return case_name

    left, right = parts[0].strip(), parts[1].strip()
    if num_plaintiffs >= 2:
        left = _ensure_side_ends_with_et_al(left)
    else:
        left = _ET_AL_TRAILING.sub("", left).rstrip().rstrip(",").strip()
    if num_defendants >= 2:
        right = _ensure_side_ends_with_et_al(right)
    else:
        right = _ET_AL_TRAILING.sub("", right).rstrip().rstrip(",").strip()

    return f"{left} v. {right}"


def read_pdf_text(pdf_path: str | Path) -> str:
    """Read a PDF file and return its text content as a string."""
    
    path = Path(pdf_path)
    
    # Confirm that the pdf file exists and raise an error if not
    if not path.exists():
        raise FileNotFoundError(f"PDF not found: {path}")
    if path.suffix.lower() != ".pdf":
        raise ValueError(f"Not a PDF file: {path}")

    # Primary path: markdown-oriented extraction.
    markdown_text = pymupdf4llm.to_markdown(path)
    if _looks_like_substantive_text(markdown_text):
        return markdown_text

    # Fallback: pypdf extraction can recover content from some PDFs where
    # markdown extraction yields mostly page metadata.
    raw_text = read_pdf_text_with_pypdf(path)
    if _looks_like_substantive_text(raw_text):
        return raw_text

    return markdown_text


def read_pdf_text_with_pypdf(pdf_path: str | Path) -> str:
    """Read text from a PDF using pypdf as a fallback extractor."""
    if not _PYPDF_AVAILABLE:
        return ""

    path = Path(pdf_path)
    try:
        reader = _PdfReader(path)
    except Exception:
        return ""

    parts: list[str] = []
    for page in reader.pages:
        text = page.extract_text()
        if text:
            parts.append(text)
    return "\n".join(parts)


def _looks_like_substantive_text(text: str) -> bool:
    """Heuristic filter to detect whether extracted PDF text has meaningful prose."""
    if not text:
        return False
    words = re.findall(r"[A-Za-z]{3,}", text)
    if len(words) < 80:
        return False
    unique_words = {w.lower() for w in words}
    if len(unique_words) < 40:
        return False
    metadata_hits = len(
        re.findall(r"(?i)\b(case|document|filed|page|cv|usdc)\b", text)
    )
    # If nearly everything is docket metadata boilerplate, treat as low quality.
    if metadata_hits > len(words) * 0.35:
        return False
    return True


def read_pdf_files_in_folder(folder_path: str | Path) -> dict[str, str]:
    """Read all PDF files in a folder and return a ``{filename: text}`` dict."""

    path = Path(folder_path)
    if not path.is_dir():
        raise NotADirectoryError(f"Not a folder: {path}")

    # pymupdf4llm is not thread-safe; PDFs must be read sequentially.
    result = {}
    for pdf_path in path.glob("*.pdf"):
        result[pdf_path.name] = read_pdf_text(pdf_path)

    # Return the dictionary created in the above loop
    return result


def read_mapping_csv(csv_path: str | Path) -> dict[str, str]:
    """
    Read any CSV file and return a dictionary where:
    - first column values are keys
    - second column values are values
    """
    mapping_path = Path(csv_path)
    if not mapping_path.exists():
        raise FileNotFoundError(f"CSV file not found: {mapping_path}")

    mapping: dict[str, str] = {}
    with mapping_path.open("r", encoding="utf-8-sig", newline="") as f:
        reader = csv.reader(f)
        first_row = next(reader, None)
        if first_row is None:
            return mapping

        if len(first_row) >= 2:
            first_col = first_row[0].replace("\u00a0", " ").strip().lower()
            second_col = first_row[1].replace("\u00a0", " ").strip().lower()
            looks_like_header = (
                first_col in {"key", "keys", "key text", "lookup", "name"}
                or "correspond" in second_col
                or second_col in {"value", "values", "mapped value", "goal", "goals"}
            )
            if not looks_like_header:
                key_text = first_row[0].replace("\u00a0", " ").strip()
                mapped_value = first_row[1].replace("\u00a0", " ").strip()
                if key_text:
                    mapping[key_text] = mapped_value

        for row in reader:
            if len(row) < 2:
                continue
            key_text = row[0].replace("\u00a0", " ").strip()
            mapped_value = row[1].replace("\u00a0", " ").strip()
            if key_text:
                mapping[key_text] = mapped_value

    return mapping


def complaint_pdf_filename_from_docket_pdf_filename(docket_pdf_name: str) -> str | None:
    """
    Map paired filenames such as 'a v. b, Docket.pdf' -> 'a v. b, Complaint.pdf'.
    Returns None if the basename does not end with ', Docket' before '.pdf'.
    """
    if not docket_pdf_name or Path(docket_pdf_name).suffix.lower() != ".pdf":
        return None
    stem = Path(docket_pdf_name).stem
    m = re.match(r"^(?P<prefix>.+),\s*Docket\s*$", stem, re.IGNORECASE)
    if not m:
        return None
    return f"{m.group('prefix')}, Complaint.pdf"


def complaint_text_for_docket_pdf(
    docket_pdf_name: str,
    complaint_text_by_name: dict[str, str],
) -> tuple[str | None, str | None]:
    """
    Find complaint PDF basename (may differ only by case from derived name) and its text.
    Returns (complaint_filename_or_derived_expected, text_or_None).
    """
    expected = complaint_pdf_filename_from_docket_pdf_filename(docket_pdf_name)
    if expected is None:
        # Try best-effort normalization-based matching for non-standard names.
        docket_key = _normalize_case_filename_for_matching(docket_pdf_name)
        for fname, txt in complaint_text_by_name.items():
            if _normalize_case_filename_for_matching(fname) == docket_key:
                return fname, txt
        return None, None
    if expected in complaint_text_by_name:
        return expected, complaint_text_by_name[expected]
    exp_lo = expected.lower()
    for fname, txt in complaint_text_by_name.items():
        if fname.lower() == exp_lo:
            return fname, txt
    expected_key = _normalize_case_filename_for_matching(expected)
    for fname, txt in complaint_text_by_name.items():
        if _normalize_case_filename_for_matching(fname) == expected_key:
            return fname, txt
    return expected, None


def _normalize_case_filename_for_matching(file_name: str) -> str:
    """
    Normalize complaint/docket PDF basenames so variant naming still matches.
    Examples:
    - 'X, Docket.pdf' <-> 'X, Complaint.pdf'
    - 'X, Docket, Complaint.pdf'
    - truncated punctuation variants such as 'v. U..pdf' vs 'v. U., Complaint.pdf'
    """
    stem = Path(file_name).stem.lower().replace("\u00a0", " ")
    stem = re.sub(r",?\s*docket\b", " ", stem)
    stem = re.sub(r",?\s*complaint\b", " ", stem)
    stem = re.sub(r"[^a-z0-9]+", " ", stem)
    return re.sub(r"\s+", " ", stem).strip()


def identify_case_name(docket_text: str) -> str:
    """Extract the case name from docket text and return it in cleaned title case."""

    # Isolate the case name from the cover page
    case_name = docket_text.split(", Docket No.")[0].split("==**")[1].strip()

    case_name = expand_agency_abbreviations_in_case_name(case_name)
    case_name = normalize_case_name_us_agencies(case_name)
    case_name = strip_case_name_honorifics(case_name)

    # Preserve camelCase brand names (e.g. "eHealth", "eBay") before lowercasing.
    # titlecase(s.lower()) would produce "Ehealth"; we map the titlecase form back
    # to the original so we can restore it after the titlecase pass.
    _camel_map: dict[str, str] = {}
    for _m in re.finditer(r'\b([a-z][A-Z][a-zA-Z]*|[A-Z][a-z]+[A-Z][a-zA-Z]*)\b', case_name):
        _word = _m.group(0)                              # e.g. "eHealth"
        _tc_form = _word[0].upper() + _word[1:].lower()  # e.g. "Ehealth"
        _camel_map[_tc_form] = _word

    # Clean up case name
    case_name = titlecase.titlecase(
        case_name.lower(), callback=case_name_acronym_callback
    )

    # Restore camelCase brand names that titlecase flattened
    for _tc_form, _original in _camel_map.items():
        case_name = re.sub(r'\b' + re.escape(_tc_form) + r'\b', _original, case_name)

    case_name = re.sub(r'\b([a-zA-Z]+(?:\.[a-zA-Z]+)+)\b', lambda m: m.group(0).upper(), case_name)
    case_name = case_name.replace("Et Al", "et al.").replace("Et Al.", "et al.")
    case_name = case_name.replace("(Lead)", "").replace(",", "").strip()
    case_name = _strip_trailing_the_each_case_name_side(case_name)

    return case_name


def identify_docket_number(docket_text: str) -> str:
    """Extract and return the docket number from docket text."""

    # Isolate the docket number
    docket_number = docket_text.split(", Docket No. ")[1].split()[0]

    # Normalize: insert missing hyphen between case-type letters and case number digits
    docket_number = re.sub(r'([a-z])(\d)', r'\1-\2', docket_number)

    return docket_number


def identify_date_filed(docket_text: str) -> str:
    """Extract and return the filing date from docket text as a DD/MM/YYYY string."""

    # Bloomberg Law dockets sometimes emit the label twice: once as a standalone
    # heading (empty value) and once with the actual date. Try all occurrences and
    # use the first that yields a parseable date.
    parts = docket_text.split("**Date Filed:**")[1:]
    for part in parts:
        candidate = part.split("**")[0].strip()
        if not candidate:
            continue
        try:
            return datetime.strptime(candidate, "%b %d, %Y").strftime("%d/%m/%Y")
        except ValueError:
            continue

    return "Date not found"


def normalize_party_name(name: str) -> str:
    """Clean and deduplicate repeated party names from extracted docket text."""
    if not name:
        return name

    cleaned = name.replace("*", " ").replace("|", " ")
    cleaned = re.sub(r"\s+", " ", cleaned).strip()
    cleaned = cleaned.replace(",", "")
    # Preserve all-caps acronym prefixes before a colon (e.g. "AGLP:") because
    # titlecase downcases uppercase words that contain vowels, turning "AGLP:" into "Aglp:".
    _acronym_prefix = re.match(r"^([A-Z]{2,}):", cleaned)
    cleaned = titlecase.titlecase(cleaned)
    if _acronym_prefix:
        cleaned = re.sub(r"^[A-Za-z]+:", _acronym_prefix.group(1) + ":", cleaned)

    # Restore Roman numerals that titlecase downcases because they contain
    # the vowel I (e.g. "II"→"Ii", "IV"→"Iv", "VI"→"Vi", "XI"→"Xi" …).
    cleaned = re.sub(
        r"\b(Ii{1,2}|Iv|Vi{0,3}|Ix|Xi{0,2}|Xiv|Xv|Xvi{0,2}|Xix|Xx)\b",
        lambda m: m.group(0).upper(),
        cleaned,
    )

    # Restore "ex rel." which titlecase capitalizes to "Ex Rel."
    cleaned = re.sub(r"\bEx\s+Rel\.", "ex rel.", cleaned)

    # Fix Irish "O'" surname prefix: titlecase treats all-caps after apostrophe as
    # a potential acronym and leaves it uppercased (e.g. "O'NEILL" → "O'NEILL").
    cleaned = re.sub(r"\bO'([A-Z][A-Z]+)\b", lambda m: "O'" + m.group(1).capitalize(), cleaned)

    cleaned = strip_party_honorifics(cleaned)

    # Remove exact repeated party names like "Foo Bar Foo Bar"
    while True:
        duplicate_match = re.fullmatch(r"(.+?)\s+\1", cleaned)
        if not duplicate_match:
            break
        cleaned = duplicate_match.group(1).strip()

    cleaned = _move_trailing_the_to_front(cleaned)

    return cleaned


def canonicalize_united_states_party(name: str) -> str:
    """If the whole party name is only the sovereign U.S., use 'United States'."""
    if name and is_united_states_sole_party(name):
        return "United States"
    return name


def identify_court(docket_text: str, district_court_code_text: str) -> dict[str, str]:
    """Given the text in a docket pdf, returns a dict with 'court' and 'division' keys."""

    # Create dict to store the court state, district, and division
    court_details = {}

    # Isolate the text containing the court and
    # identify the applicable state, district, and division
    court_text = docket_text.split("CIVIL DOCKET FOR CASE #:")[0].split("District Court")[1]

    # Normalize to upper case for matching
    capitalized_text = court_text.upper()

    # Establish default values
    court_state = "Unknown"
    court_district = "NONE"

    # Try to find any state name as a substring
    for state in US_STATES:
        if state in capitalized_text:
            court_state = state
            possible_districts = DISTRICTS.get(court_state, set())
            # Look for a matching district for that state
            for district in possible_districts:
                if district in capitalized_text:
                    court_district = district
                    break
            break

    # Compose human-readable court string (state + district)
    if court_district == "NONE":
        court = court_state.title()
    else:
        court = f"{court_state.title()} {court_district.title()}"

    court_details["court"] = court
    court_details["division"] = identify_division(district_court_code_text, docket_text, court_state, court_district)

    return court_details


def isolate_state_text(code_text: str, state: str) -> str:
    """Given the text of 28 USC Ch5 and a state name, returns the portion pertaining to that state."""
    
    # Create a variable containing the section heading for the current state
    state_index = US_STATES.index(state.upper())
    section_number = SECTION_NUMBERS[state_index]
    section_heading = f"§{section_number}. {state}"

    # Capitalize the district court code text to standardize
    district_court_code_text = code_text.upper()

    # Isolate the text containing the relevant information for the appropriate state
    state_segment = district_court_code_text.split(section_heading)[1]
    
    if state == "WYOMING":
        end_segment = "HISTORICAL"
    else:
        next_state_index = state_index + 1
        next_state = US_STATES[next_state_index]
        next_state_section_number = SECTION_NUMBERS[next_state_index]
        end_segment = f"§{next_state_section_number}. {next_state}"
    
    state_segment = state_segment.split(end_segment)[0]

    return state_segment


def generate_districts_list(text: str) -> list[str]:
    """Given the portion of 28 USC Ch5 for a state, returns the list of district names."""

    # Isolate and clean districts and put into list
    districts_text = text.upper().split("TO BE KNOWN AS THE")[1].split("DISTRICTS OF")[0]
    districts_text_clean = districts_text.replace("\n", " ")

    if "," in districts_text_clean:
        districts_list = [district.strip() for district in districts_text_clean.split(",") if district.strip()]
        districts_list[-1] = districts_list[-1].replace("AND ", "").replace(".", "").strip()
    else:
        districts_list = [district.strip() for district in districts_text_clean.split("AND") if district.strip()]
        districts_list[-1] = districts_list[-1].replace(".", "").strip()

    return districts_list
    

def isolate_district_text(state_segment: str, district: str) -> str:
    """Given a state's code segment and a district name, returns the portion specific to that district."""
    
    if "ONE JUDICIAL DISTRICT" in state_segment:
        district_segment = state_segment
    else:
        # Generate list of districts within state
        districts_list = generate_districts_list(state_segment)
        if not districts_list or district not in districts_list:
            # Cannot map district in code text; return whole state segment for best-effort division check
            district_segment = state_segment
            return district_segment
        else:
            # Determine how to find the end of the text segment pertaining to the given district
            district_index = districts_list.index(district)
            if district_index + 1 == len(districts_list):
                end_district_text = "NOTES"
            else:
                end_district_text = f"{districts_list[district_index + 1].upper()} DISTRICT"

        # Figure out appropriate index for district information based on 
        # how many times district name appears
        district_name = f"{district} DISTRICT"
        parts = state_segment.split(district_name)
        if len(parts) >= 3:
            district_segment = parts[2]
        elif len(parts) == 2:
            district_segment = parts[1]
        else:
            district_segment = ""

        # Use the split function to isolate the text segment pertaining to the given district
        district_segment = district_segment.split(end_district_text)[0]

    return district_segment


def identify_division(district_court_code_text: str, docket_text: str, state: str, district: str) -> str:
    """Given the state, district, and docket text, returns the division name or 'None'."""

    if state and state.title() == "Minnesota":
        return "Division exists but requires further research"

    # Identify the portion of the docket text that contains information regarding the division
    docket_division_segment = docket_text.split("CIVIL DOCKET FOR CASE #")[0]

    # Look for an explicit "word Division" label in the court header.
    # Using regex handles nested parens like "(Southern Division (1))" where
    # rfind("(") would incorrectly grab the inner "(1)" instead of the outer phrase.
    m = re.search(r"\b([A-Za-z]+)\s+Division\b", docket_division_segment, re.IGNORECASE)
    if m:
        return m.group(1).title()

    # Fall back to last-parenthetical extraction for the 28 USC code lookup.
    beginning_index = docket_division_segment.rfind("(")
    ending_index = docket_division_segment.rfind(")")
    docket_division = docket_division_segment[beginning_index + 1 : ending_index].upper()

    # district_court_code_text is pre-loaded in main() and passed in — no file read needed here.
    # Reaches here only when no explicit "word Division" label was found in the court header.

    # Isolate the portion of the code pertaining to the given state
    state_segment = isolate_state_text(district_court_code_text, state)

    # Isolate the text containing the relevant information for the appropriate district
    district_segment = isolate_district_text(state_segment, district)
    
    # If there are no divisions, return "None"
    if "DIVISION" not in district_segment:
        return "None"

    # Identify the division based on docket_division and patterns within the statute
    if docket_division not in district_segment:
        return "None"
    elif district_segment.split(docket_division)[1].split()[0] == "DIVISION":
        division_segment = docket_division
    else:
        division_segment = district_segment.split(docket_division)[0].split(") THE")[1].split("DIVISION")[0].strip()

    return division_segment.title()


def identify_judge(docket_text: str) -> str:
    """Given the text in a docket pdf, returns the formatted name of the assigned judge."""

    # Account for possibility of not identifying judge
    if "**Assigned to:**" not in docket_text:
        return "Judge not found"

    # Isolate and clean the text with the Judge's name and title
    judge_full = docket_text.split("**Assigned to:**")[1]
    judge_full = judge_full.split("**")[0].strip()

    # Set up variable for isolating judge's name
    judge_name = judge_full.split()

    # Set up variable to track what word begins judge's actual name
    name_starting_index = 0

    # Adjust name starting index based on whether title keywords precede judge's name
    for word in judge_name:
        if word.upper() in JUDGE_TITLE_KEYWORDS:
            name_starting_index += 1
        else:
            break
    
    # Adjust judge name to reflect updated starting index
    judge_name = judge_name[name_starting_index:]
    judge_name = " ".join(judge_name)
    judge_name_split = split_name(judge_name)

    formatted_judge_name = judge_name_split[2] + ", " + judge_name_split[0] + " " + judge_name_split[1]
    formatted_judge_name = titlecase.titlecase(formatted_judge_name).strip()

    return formatted_judge_name


def _normalize_name_for_lookup(s: str) -> str:
    normalized_string = s.strip().upper()
    normalized_string = re.sub(r"[^\w\s]", " ", normalized_string)
    normalized_string = re.sub(r"\s+", " ", normalized_string).strip()
    return normalized_string


def split_name(full_name: str) -> tuple[str, str, str]:
    parts = full_name.split()
    first_name = parts[0] if len(parts) > 0 else ""
    middle_name = " ".join(parts[1:-1]) if len(parts) > 2 else ""
    last_name = parts[-1] if len(parts) > 1 else ""
    return first_name, middle_name, last_name


def identify_president(judge_name: str) -> str:
    """
    Given a string containing the name of a judge, returns a string containing the name
    of the president who appointed that judge, based on the data found at and previously
    downloaded from https://www.fjc.gov/history/judges 
    """
    
    # This function reads from local file created by update_judge_data.py.
    # Expected local file: ./judges.csv (project root / current script directory)
    local_csv = Path(__file__).resolve().parent.parent / JUDGE_CSV_FILE
    if not local_csv.exists():
        return "President not found (run update_judge_data.py first)"

    # `identify_judge()` returns "Last, First Middle". We must preserve that structure:
    judge_last_raw = ""
    judge_first_raw = ""
    judge_middle_raw = ""
    if "," in (judge_name or ""):
        last_part, rest = judge_name.split(",", 1)
        judge_last_raw = last_part.strip()
        rest_parts = rest.strip().split()
        if rest_parts:
            judge_first_raw = rest_parts[0]
            judge_middle_raw = " ".join(rest_parts[1:]) if len(rest_parts) > 1 else ""
    else:
        first, middle, last = split_name((judge_name or "").strip())
        judge_first_raw, judge_middle_raw, judge_last_raw = first, middle, last

    judge_last = _normalize_name_for_lookup(judge_last_raw)
    judge_first = _normalize_name_for_lookup(judge_first_raw)
    judge_middle = _normalize_name_for_lookup(judge_middle_raw)
    if not judge_last or not judge_first:
        return "President not found. Judge name not parsed."

    with local_csv.open("r", encoding="utf-8", newline="") as f:
        reader = csv.DictReader(f)

        # Set up variable to store imprecise matches
        potential_matches = {}
        candidate_rows = []

        # First pass: collect only rows with matching normalized last name.
        # This avoids relying on CSV sort order (which can break with diacritics).
        for row in reader:
            row_judge_last = _normalize_name_for_lookup(row.get("Last Name", ""))
            if row_judge_last == judge_last:
                candidate_rows.append(row)

        # Second pass: apply current first/middle matching logic on candidates.
        for row in candidate_rows:
            # Store name variables from csv file
            row_judge_last = _normalize_name_for_lookup(row.get("Last Name", ""))
            row_judge_first = _normalize_name_for_lookup(row.get("First Name", ""))
            if not row_judge_first:
                continue
            row_judge_first_init = row_judge_first[0]
            row_judge_middle = _normalize_name_for_lookup(row.get("Middle Name", ""))
            if len(row_judge_middle) > 0:
                row_judge_middle_init = row_judge_middle[0]
            else:
                row_judge_middle_init = ""

            # Look for exact match and imprecise matches
            appointing_president = row.get("Appointing President (1)", "").strip()
            match_key = f"{row_judge_first} {row_judge_middle} {row_judge_last}"
            if row_judge_first == judge_first and row_judge_middle == judge_middle and row_judge_last == judge_last:
                potential_matches[match_key] = appointing_president
            elif (row_judge_first == judge_first and row_judge_last == judge_last and not judge_middle):
                potential_matches[match_key] = appointing_president
            elif row_judge_first_init == (judge_first[0] if judge_first else "") and row_judge_middle == judge_middle and row_judge_last == judge_last:
                potential_matches[match_key] = appointing_president
            elif row_judge_first == judge_first and row_judge_middle_init == (judge_middle[0] if judge_middle else "") and row_judge_last == judge_last:
                potential_matches[match_key] = appointing_president
    
    # Return results based on number of matches identified
    if not potential_matches:
        return "President not found. Try running update_judge_data.py and reviewing Federal Judicial Center Export."
    elif len(potential_matches) == 1:
        return list(potential_matches.values())[0]
    else:
        parts = ["Multiple possibilities identified:"]
        for name, president in potential_matches.items():
            parts.append(f" If the judge's full name is {name.title()}, then the judge was appointed by {president}.")
        return "".join(parts)


def parse_page_in_columns(
    page: fitz.Page,
    include_text: bool = False,
    *,
    gutter_ratio: float = 0.5,
) -> tuple[str, str, bool]:
    """
    Split visible text into left and right columns using span bounding boxes.

    gutter_ratio: horizontal split as a fraction of page width (0.5 = middle).
    Spans whose center x is left of (width * gutter_ratio) go to the left column.
    """
    rect = page.rect
    mid_x = rect.width * gutter_ratio

    left: list[tuple[float, float, str]] = []
    right: list[tuple[float, float, str]] = []

    data = page.get_text("dict")

    for block in data.get("blocks", []):
        if block.get("type") != 0:
            continue
        for line in block.get("lines", []):
            t = "".join(span.get("text", "") for span in line.get("spans")).strip()
            if not t:
                continue
            if "PARTIES AND ATTORNEYS" in t.upper():
                include_text = True
            if not include_text:
                continue
            x0, y0, x1, y1 = line["bbox"]
            if x0 < mid_x:
                left.append((y0, x0, t))
            else:
                right.append((y0, x0, t))
            if "DOCKET ENTRIES" in t.upper():
                include_text = False

    def join_column(items: list[tuple[float, float, str]]) -> str:
        items.sort(key=lambda r: (r[0], r[1]))
        parts: list[str] = []
        last_y: float | None = None
        for y0, _x0, t in items:
            if last_y is not None:
                line_gap = y0 - last_y
                if line_gap > 18:
                    parts.append("\n\n")
                elif line_gap > 5:
                    parts.append("\n")
                else:
                    parts.append(" ")
            parts.append(t)
            last_y = y0
        return "".join(parts).strip()

    return join_column(left), join_column(right), include_text


def split_columns_into_blocks(column_text: str) -> list[str]:
    """Split a column's text into separate representative blocks."""
    return [block.strip() for block in column_text.split("\n\n") if block.strip()]


def read_pages_into_columns(pdf_path: str | Path) -> dict[str, str]:
    """Read a docket PDF's party section into left and right column text."""

    doc = fitz.open(pdf_path)
    left_parts: list[str] = []
    right_parts: list[str] = []
    include_text = False

    for page in doc:
        left_column, right_column, include_text = parse_page_in_columns(page, include_text)
        if "PRINTED BY" in left_column.upper() or "PRINTED BY" in right_column.upper():
            continue
        left_parts.append(left_column)
        right_parts.append(right_column)
        if "DOCKET ENTRIES" in left_column.upper() or "DOCKET ENTRIES" in right_column.upper():
            break

    return {"Left Column": "".join(left_parts), "Right Column": "".join(right_parts)}


def join_split_blocks(left_blocks: list[str], right_blocks: list[str]) -> str:
    """Given two lists of strings representing blocks of text, combines them 
    into a string in the order they'd be read on the page, 
    while identifying blocks that were split mid-block and joins them together"""

    side = "left"
    left_idx = 0
    right_idx = 0
    parts: list[str] = []
    keywords_to_continue = ["TERMINATED", "ATTORNEY"]

    while left_idx < len(left_blocks) or right_idx < len(right_blocks):
        if side == "left":
            if left_idx >= len(left_blocks):
                side = "right"
                continue
            next_block = left_blocks[left_idx]
            left_idx += 1
            next_block_lower = next_block.lower()
            switch_sides = any(keyword in next_block for keyword in keywords_to_continue) or next_block_lower in {"plaintiff", "plaintiffs", "defendant", "defendants"}
            if switch_sides:
                parts.append(next_block + "\n\n")
                if right_idx < len(right_blocks):
                    side = "right"
            else:
                parts.append(next_block + " ")
        else:
            if right_idx >= len(right_blocks):
                side = "left"
                continue
            next_block = right_blocks[right_idx]
            right_idx += 1
            next_block_lower = next_block.lower()
            switch_sides = any(keyword in next_block for keyword in keywords_to_continue) or next_block.count("\n") == 0 or next_block_lower in {"plaintiff", "plaintiffs", "defendant", "defendants"}
            if switch_sides:
                parts.append(next_block + "\n\n")
                if left_idx < len(left_blocks):
                    side = "left"
            else:
                parts.append(next_block + " ")

    return "".join(parts).strip()


def clean_blocks(blocks: str, title: str) -> str:
    result = blocks.replace(title, "")
    result = result.replace(BLOOMBERG_COPYRIGHT_STRING, "")
    result = re.sub(r"(?im)\s*//\s*PAGE\s*\d+\s*$", "", result)
    result = re.sub(r"(?i)//\s*PAGE\s*\d+", "", result)
    return result


def _contains_state_of_us_state_phrase(text: str) -> bool:
    """True only for explicit phrases like 'State of Texas' (avoids matching 'Estate of ...')."""
    lowered = re.sub(r"\s+", " ", (text or "").lower())
    return any(
        re.search(rf"\bstate\s+of\s+{re.escape(state.lower())}\b", lowered)
        for state in US_STATES
    )


# Fallback map: email domain → canonical firm name.
# Used when a party's attorneys are listed with email addresses only (no firm
# name text in the extracted block), which happens in some Bloomberg dockets.
_EMAIL_DOMAIN_FIRM_MAP: dict[str, str] = {
    "mofo.com": "Morrison & Foerster LLP",
}


def get_lead_counsel_from_block(block: str) -> str:
    """Given a string containing a block of text pertaining to either the plaintiff(s) or defendant(s),
    returns a string containing the name of the lead counsel for the plaintiff(s) or defendant(s) in that block"""

    divided_block = block.split("\n\n")

    if re.search(r"official\s+capacity", block, re.IGNORECASE) or _contains_state_of_us_state_phrase(block) or "commonwealth of" in block.lower():
        return "Government"

    for section in divided_block:
        if "TERMINATED" in section.upper() or "MOVANT" in section.upper() or "AMICUS" in section.upper():
            continue
        elif any(gov_keyword in section for gov_keyword in GOV_KEYWORDS) or re.search(
            r"\b(?:U\.?S\.?|United\s+States)\s+Attorney'?s?\s+Office",
            section,
            re.IGNORECASE,
        ) or re.search(
            r"\b(?:District|County|City|State|Solicitor)\s+Attorney'?s?\s+Office"
            r"|\bOffice\s+of\s+(?:the\s+)?(?:District|County|City|State|Solicitor)\s+Attorney\b",
            section,
            re.IGNORECASE,
        ):
            return "Government"
        elif "\n" not in section:
            continue
        else:
            lines = section.split("\n")
            section_counsel = ""
            if "  " in lines[0]:
                candidate = lines[0].split("  ")[1].strip()
                if (not any(c.isdigit() for c in candidate)
                        and "@" not in candidate
                        and not re.search(r"(?i)\bsee\s+above\b", candidate)
                        and not _NOTICE_RE.search(candidate)):
                    section_counsel = candidate + " "
            skip_next = False
            for line in lines[1:]:
                if skip_next:
                    skip_next = False
                    continue
                if "Representation" in line:
                    # If the attorney name is hyphen-split across lines
                    # (e.g. "...Gonzalez-" / "Pagan"), skip the continuation
                    # line so it doesn't bleed into the firm name.
                    if line.rstrip().endswith("-"):
                        skip_next = True
                    continue
                elif any(char.isdigit() for char in line):
                    break
                elif "@" in line:
                    # Email address — not a firm name; skip but remember domain
                    continue
                elif re.search(r"(?i)\bsee\s+above\b", line):
                    continue
                elif _NOTICE_RE.search(line):
                    continue
                section_counsel += line + " "

            if section_counsel.strip():
                return titlecase.titlecase(section_counsel).strip()
            # No firm name found in this section — continue to the next one.

    # Fallback: infer firm from the first non-terminated attorney's email domain.
    for section in divided_block:
        if "TERMINATED" in section.upper() or "MOVANT" in section.upper() or "AMICUS" in section.upper():
            continue
        email_m = re.search(r"@([\w.]+)", section)
        if email_m:
            domain = email_m.group(1).lower()
            if domain in _EMAIL_DOMAIN_FIRM_MAP:
                return _EMAIL_DOMAIN_FIRM_MAP[domain]

    return "Counsel not identified"


def apply_spell_check(name: str) -> str:
    """Replace any known misspellings in a counsel name using SPELL_CHECK."""
    if not name:
        return name

    corrected = name
    for wrong, right in SPELL_CHECK.items():
        corrected = re.sub(rf"\b{re.escape(wrong)}\b", right, corrected, flags=re.IGNORECASE)

    return corrected


def _plaintiff_text_from_first_plaintiff_label(text: str) -> str:
    """If the docket lists movants/others before the party block, start at the first 'Plaintiff' label."""
    m = re.search(r"(?i)\bplaintiff\b", text)
    return text[m.start() :] if m else text


def identify_lead_counsel(pdf_path: str | Path, title: str) -> tuple[str, str]:
    """Given a docket PDF path, returns (plaintiff_lead_counsel, defendant_lead_counsel)."""
    
    columns = read_pages_into_columns(pdf_path)
    left_column = columns["Left Column"]
    left_column_blocks = split_columns_into_blocks(left_column)[1:]
    right_column = columns["Right Column"] 
    right_column_blocks = split_columns_into_blocks(right_column)
    
    combined_blocks = join_split_blocks(left_column_blocks, right_column_blocks)
    combined_blocks = clean_blocks(combined_blocks, title)

    parts = re.split(r"Defendant", combined_blocks, maxsplit=1, flags=re.IGNORECASE)
    plaintiff_blocks = _plaintiff_text_from_first_plaintiff_label(parts[0])
    defendant_blocks = parts[1] if len(parts) > 1 else ""

    p_lead_counsel = get_lead_counsel_from_block(plaintiff_blocks)
    d_lead_counsel = get_lead_counsel_from_block(defendant_blocks)

    p_lead_counsel = apply_spell_check(p_lead_counsel)
    d_lead_counsel = apply_spell_check(d_lead_counsel)

    return p_lead_counsel, d_lead_counsel


def identify_title(docket_text: str) -> str:
    """Given the text in a docket pdf as input, returns the title of the case."""
    return docket_text.split("**Current")[0].strip().split("\n")[-1].strip()


def _party_preamble_before_representation(party_chunk: str) -> str:
    """Text for one party block up to (but not including) the first counsel Representation line."""
    m = re.search(r"(?i)\brepresentation\b", party_chunk)
    if m:
        return party_chunk[: m.start()]
    return "\n".join(party_chunk.split("\n")[:15])


def _party_marked_terminated_in_docket(party_chunk: str) -> bool:
    """True when the docket marks the party itself as terminated (not counsel status)."""
    if not party_chunk:
        return False

    preamble = _party_preamble_before_representation(party_chunk) or ""
    lines = [ln.strip() for ln in preamble.splitlines() if ln.strip()]
    # Keep termination checks tightly scoped to the party header area.
    header_lines = lines[:12]
    header_text = "\n".join(header_lines)

    # Explicit party-status phrasing should always count.
    if re.search(r"(?i)\b(?:status|party\s+status)\b.{0,40}\bterminated\b", header_text):
        return True

    # Pattern like "John Doe (TERMINATED)" in party header.
    if re.search(r"(?i)\(\s*terminated\s*\)", header_text):
        return True

    # Line-level check so attorney/counsel termination lines do not mask party termination.
    for line in header_lines:
        if not _TERMINATED_MARKER_RE.search(line):
            continue
        if re.search(r"(?i)\battorney\b|\bcounsel\b", line) and not re.search(
            r"(?i)\bparty\b|\bplaintiff\b|\bdefendant\b|\bintervenor\b",
            line,
        ):
            continue
        return True

    return False


def _name_after_representation(block: str) -> str:
    """Extract party name from a block where the name follows a 'Representation' label.

    More robust than the previous split("**")[2] approach, which relied on a
    fixed number of bold markers before the name.
    """
    parts = re.split(r"\bRepresentation\b", block, maxsplit=1, flags=re.IGNORECASE)
    if len(parts) < 2:
        return ""
    for segment in parts[1].split("**"):
        segment = segment.strip()
        if segment and not re.match(r"^#+$", segment):
            return segment
    return ""


def _is_party_terminated_quick(block: str) -> bool:
    """Return True when the party header signals TERMINATED before counsel info.

    Scans only the text up to the first ATTORNEY keyword, digit, or parenthesis
    so that attorney-level TERMINATED markers deeper in the block do not
    accidentally skip the whole party.  Respects the 'formerly known as' guard:
    a TERMINATED that appears only inside an alias clause is not a true skip.
    """
    attorney_idx = block.upper().find("ATTORNEY")
    digit_match = re.search(r"\d", block)
    digit_idx = digit_match.start() if digit_match else -1
    paren_match = re.search(r"\(", block)
    paren_idx = paren_match.start() if paren_match else -1
    candidates = [i for i in (attorney_idx, digit_idx, paren_idx) if i >= 0]
    boundary = min(candidates) if candidates else len(block)
    term_area = block[:boundary]
    terminated_pos = term_area.upper().find("TERMINATED")
    formerly_pos = term_area.lower().find("formerly known as")
    return terminated_pos >= 0 and (formerly_pos < 0 or terminated_pos < formerly_pos)


def _parse_party_name_from_block(block: str, *, strip_capacity: bool = False) -> str:
    """Extract the raw party name from a single post-role-label block.

    Handles three docket formats:
    - Standard PACER / inline: name on the first line after the role marker.
    - Table: name in the first cell, possibly continued on subsequent rows.
    - Bloomberg heading: name in a '## **NAME**' element inside the block.

    ``strip_capacity`` removes trailing capacity phrases such as
    "in his official capacity" (used for defendants).
    """
    first_line = block.split("\n")[0].strip().strip("**")

    if first_line:
        tentative = first_line
        if "Representation" in tentative:
            tentative = tentative.split("Representation")[0].strip().strip("**")
        if "**" in tentative and not tentative.startswith("|"):
            pre_close = tentative.split("**")[0].strip()
            if pre_close:
                tentative = pre_close
        if strip_capacity:
            capacity_m = re.search(
                r"\bin (?:his|her|their)(?:\s+official)?\s+capacity\b",
                tentative,
                re.IGNORECASE,
            )
            if capacity_m:
                tentative = tentative[: capacity_m.start()].strip().strip("**")
        if "|" in tentative:
            for cont_line in block.split("\n")[1:]:
                stripped = cont_line.strip()
                if re.match(r"^\|[-|]+\|$", stripped):
                    continue
                cont_m = re.match(r"^\|\|+\*\*([^|*]+)\*\*\|", stripped)
                if cont_m:
                    extra = cont_m.group(1).strip()
                    if re.match(r"^doing\s+business", extra, re.IGNORECASE):
                        break
                    tentative += " " + extra
                else:
                    break
    else:
        bloomberg_m = re.search(
            r"##\s+\*\*([^*\n]+)\*\*\s*\n+##\s+ATTORNEY\s+TO\s+BE\s+NOTICED",
            block,
            re.IGNORECASE,
        )
        if bloomberg_m:
            tentative = bloomberg_m.group(1).strip()
        else:
            tentative = next(
                (p.strip() for p in block.split("**")[1:]
                 if p.strip() and not re.match(r"^#+$", p.strip())),
                "",
            )

    if "Representation" in tentative:
        return _name_after_representation(block)
    return tentative


def _extract_intervenors(text: str, role: str) -> list[str]:
    """Extract intervenor party names from a docket section.

    Handles all three Bloomberg Law formats (table, heading, inline) as well
    as standard PACER bold-marker format.  ``role`` should be "Plaintiff" or
    "Defendant" to match the correct intervenor label.
    """
    blocks = re.split(
        rf"\*\*Intervenor\s+{re.escape(role)}[s]?\b",
        text,
        flags=re.IGNORECASE,
    )
    names: list[str] = []
    for block in blocks[1:]:
        first_line = block.split("\n")[0].strip()

        table_m = re.match(r"\*?\*?\|+\*\*([^|*\n]+)\*\*\|", first_line)
        if table_m:
            name = table_m.group(1).strip()
        elif not first_line.strip("*").strip():
            heading_m = re.search(r"##\s+\*\*([^*\n]+)\*\*", block)
            name = heading_m.group(1).strip() if heading_m else ""
        else:
            raw = first_line.lstrip("*").strip()
            if "Representation" in raw:
                raw = raw.split("Representation")[0].strip()
            if "**" in raw:
                raw = raw.split("**")[0].strip()
            name = raw.strip(",").strip()

        if name and not _party_marked_terminated_in_docket(block):
            names.append(canonicalize_united_states_party(normalize_party_name(name)))
    return names


def _normalize_party_labels(text: str) -> str:
    """Normalize heading-formatted party labels to standard bold format.

    Some docket providers (e.g. Bloomberg Law) emit party role labels as
    Markdown headings (``## **plaintiff**``) rather than the standard bold
    form (``**Plaintiff``).  Normalizing them here lets all downstream
    split/count logic use a single consistent pattern.

    Also handles page-break-separated names: Bloomberg Law sometimes places
    the party name on the line immediately before ``**Plaintiff**`` due to a
    PDF page break splitting the entry.  Convert to inline form so the name
    follows the role label.
    """
    text = re.sub(r"#{1,6}\s+\*\*(plaintiff)\*\*", "**Plaintiff", text, flags=re.IGNORECASE)
    text = re.sub(r"#{1,6}\s+\*\*(defendant)\*\*", "**Defendant", text, flags=re.IGNORECASE)
    # e.g. "**SARAH HENN**\n\n**Plaintiff**" → "**Plaintiff **SARAH HENN**"
    text = re.sub(
        r"\*\*([^*\n|]{1,80})\*\*(\s*\n+)\*\*Plaintiff\*\*",
        r"**Plaintiff **\1**",
        text,
        flags=re.IGNORECASE,
    )
    return text


def identify_plaintiffs(docket_text):
    """
    Given the text in a docket pdf, returns a dictionary where the keys are
    "Plaintiff" and "Intervenor Plaintiff(s)" and the values are strings containing
    the names of the plaintiffs and intervenor plaintiffs in the case, separated by semicolons
    """
    docket_text = _normalize_party_labels(docket_text)

    # Isolate the portion of the docket text that will contain all of the plaintiffs
    plaintiff_text = docket_text.split("**Defendant")[0]

    # Intervenor plaintiffs are easiest to identify by role labels in this section.
    intervenor_ps_list = _extract_intervenors(plaintiff_text, "Plaintiff")

    # Remove intervenor lines from the primary plaintiff parsing block so they do not
    # get counted as regular plaintiffs.
    plaintiff_text_main = re.sub(
        r"Intervenor\s+Plaintiff[s]?[^\n]*",
        "",
        plaintiff_text,
        flags=re.IGNORECASE,
    )

    num_plaintiffs = plaintiff_text_main.count("**Plaintiff")
    all_p_blocks = plaintiff_text_main.split("**Plaintiff")
    plaintiffs_list = []
    skip_indices: set[int] = set()

    for i in range(num_plaintiffs):
        if i in skip_indices:
            continue

        block = all_p_blocks[i + 1]

        if _is_party_terminated_quick(block):
            continue

        name = _parse_party_name_from_block(block)

        # Qui tam / ex rel: merge relator name from the next block.
        if name:
            pre_rep = re.split(r"\bRepresentation\b", block, maxsplit=1, flags=re.IGNORECASE)[0]
            if re.search(r"\bex\s+rel\b", pre_rep, re.IGNORECASE):
                for j in range(i + 2, len(all_p_blocks)):
                    relator_block = all_p_blocks[j]
                    if re.search(r"\brelator\b", relator_block[:300], re.IGNORECASE):
                        relator_first_line = relator_block.split("\n")[0].strip().strip("**").split("**")[0].strip()
                        if relator_first_line:
                            name = f"{name} ex rel. {relator_first_line.strip()}"
                        skip_indices.add(j - 1)
                        break

        if name:
            plaintiffs_list.append(
                canonicalize_united_states_party(normalize_party_name(name))
            )

    return {
        "Plaintiffs": "; ".join(plaintiffs_list) or "None",
        "Intervenor Plaintiffs": "; ".join(intervenor_ps_list) or "None",
    }


def extract_defendants_from_raw_text(pdf_path: Path) -> list[str]:
    """Extract defendant names from raw PDF text using fitz as a fallback.
    Returns a list of defendant names found in the raw text."""
    try:
        doc = fitz.open(pdf_path)
        text = ''.join(page.get_text() for page in doc)

        # Isolate the parties section and skip any docket entries or later content.
        if "Parties and Attorneys" not in text:
            return []
        parties = text.split("Parties and Attorneys", 1)[1]
        if "Docket Entries" in parties:
            parties = parties.split("Docket Entries", 1)[0]

        # Capture all lines immediately following a Defendant marker.
        defendants = []
        for match in re.finditer(r"(?:^|\n)\s*(?!Intervenor\s)Defendant\s*\n([^\n]+)", parties, flags=re.IGNORECASE):
            name_line = match.group(1).strip()
            if name_line and "TERMINATED" not in name_line.upper():
                defendants.append(name_line)

        # If no names were found with the primary pattern, fall back to a looser search.
        if not defendants:
            for match in re.finditer(r"(?:^|\n)\s*Defendant\s*([^\n]+)", parties, flags=re.IGNORECASE):
                name_line = match.group(1).strip()
                if (
                    name_line
                    and name_line.lower() != "representation"
                    and "TERMINATED" not in name_line.upper()
                ):
                    defendants.append(name_line)

        return defendants
    except Exception:
        return []


_ENTITY_SPLIT_PAT = re.compile(
    r'\b(Inc|LLC|LLP|LLLP|Corp|Ltd|P\.A|P\.C|N\.A|L\.P)\.?\s+(?=[A-Z])',
    re.IGNORECASE,
)

def _split_merged_entities(name: str) -> list[str]:
    """Split a string that may contain multiple merged legal entity names.

    Splits after a legal suffix (Inc., LLC, Corp., etc.) when followed by
    what looks like the start of a new entity (a capital letter). Returns
    a single-element list when no merge boundary is detected.
    """
    s = re.sub(r"\s+", " ", name).strip()
    result = _ENTITY_SPLIT_PAT.sub(r'\1\n', s)
    parts = [p.strip() for p in result.split('\n') if p.strip()]
    return parts if len(parts) > 1 else [s]



def identify_defendants(docket_text, pdf_path=None):
    """ Given the text in a docket pdf, returns a dictionary in which the keys are
    "Defendant" and "Intervenor Defendant" and the values are strings containing
    the names of the defendants and intervenor defendants in the case, separated by semicolons
    """
    docket_text = _normalize_party_labels(docket_text)

    # Isolate the portion of the docket text that will contain all of the defendants
    defendant_text = docket_text.split("Parties and Attorneys")[1].split("Docket Entries")[0]

    intervenor_ds_list = _extract_intervenors(defendant_text, "Defendant")

    num_defendants = defendant_text.count("**Defendant")
    all_d_blocks = defendant_text.split("**Defendant")
    defendant_list = []

    for i in range(num_defendants):
        block = all_d_blocks[i + 1]

        if _is_party_terminated_quick(block):
            continue

        name = _parse_party_name_from_block(block, strip_capacity=True)

        if name:
            for expanded_name in _split_merged_entities(name):
                normalized = canonicalize_united_states_party(normalize_party_name(expanded_name))
                if re.fullmatch(r"(?i)u\.?s\.?\s+anesthesia\s+partners\s+of\s+texas(?:\s+p\.?a\.?)?", normalized):
                    normalized = "U.S. Anesthesia Partners of Texas, P.A."
                defendant_list.append(normalized)

    # De-duplicate while preserving order.
    deduped_defendants: list[str] = []
    seen_defs: set[str] = set()
    for d in defendant_list:
        key = re.sub(r"[^a-z0-9]+", " ", d.lower()).strip()
        if key in seen_defs:
            continue
        seen_defs.add(key)
        deduped_defendants.append(d)

    return {
        "Defendants": "; ".join(deduped_defendants)
        or DEFENDANTS_PARSE_FAILURE_MESSAGE,
        "Intervenor Defendants": "; ".join(intervenor_ds_list) or "None",
    }


def identify_goals(
    complaint_text: str,
    llm_client: Any,
    llm_context: str = "",
    is_agency_defendant: bool = False,
) -> str:
    """Classify relief goals from prayer text using an LLM via OpenRouter.

    ``is_agency_defendant`` injects a disambiguation note that steers the model
    toward or away from agency-specific labels (e.g. "Block enforcement of an
    agency action" vs "Block defendant action"), which are easily confused when
    the prayer language is similar.

    Returns a semicolon-separated, alphabetically sorted string of goal labels
    drawn from _GOAL_OUTPUT_ORDER, or 'Not identified' if the call fails or
    the response contains no recognizable labels.
    """

    if llm_client is None:
        return "Not identified"

    requests = _extract_relief_requests(complaint_text)
    relief_requests = "\n\n".join(requests) if requests else complaint_text

    valid_labels = set(_GOAL_OUTPUT_ORDER)
    label_list = "\n".join(f"- {g}" for g in sorted(valid_labels))

    context_block = f"{llm_context}\n" if llm_context else ""

    if is_agency_defendant:
        agency_note = (
            "DEFENDANT TYPE: The defendant in this case IS a government agency "
            "or government official. Prefer agency-specific labels where "
            "applicable:\n"
            '  - Use "Block enforcement of an agency action" (not "Block '
            'defendant action") for injunctions against an agency rule/decision.\n'
            '  - Use "Declaration that agency action is unlawful" (not '
            '"Declaration that defendant action is unlawful") for declaratory '
            "relief about agency conduct.\n"
            '  - Use "Vacate agency action" for requests to set aside an agency '
            "rule or order.\n"
            '  - Use "Compel agency action" for mandamus-type relief.\n\n'
        )
    else:
        agency_note = (
            "DEFENDANT TYPE: The defendant in this case is a private party, "
            "NOT a government agency. Avoid agency-specific labels "
            '("Block enforcement of an agency action", "Declaration that agency '
            'action is unlawful", "Vacate agency action", "Compel agency action") '
            "unless the text explicitly references agency conduct.\n\n"
        )

    damages_note = (
        'AWARD DAMAGES — only select "Award damages" when the prayer explicitly '
        "requests a monetary award to compensate for harm, such as compensatory, "
        "punitive, treble, nominal, or consequential damages. Do NOT select it for: "
        "requests limited to attorneys' fees or litigation costs; catch-all phrases "
        'like "such other and further relief as this Court deems just and proper"; '
        "or cases seeking primarily injunctive, declaratory, or vacatur relief "
        "where no specific damages claim is asserted. Disgorgement of profits or "
        'wrongful payments belongs under "Award restitution", not "Award damages".\n\n'
    )

    prompt = (
        "You are a legal document classifier. Read the prayer for relief below "
        "and identify which goals from the following fixed list are being sought.\n\n"
        f"{agency_note}"
        f"{damages_note}"
        f"{context_block}"
        f"Available goal labels:\n{label_list}\n\n"
        "Instructions:\n"
        "1. Return ONLY labels from the list above, spelled exactly as shown.\n"
        "2. Sort the selected labels alphabetically.\n"
        "3. Separate labels with a semicolon and a space ('; ').\n"
        "4. If none of the listed goals apply, return exactly: Not identified\n"
        "5. Do not include any explanation, preamble, or additional text.\n\n"
        f"Prayer for relief:\n{relief_requests}"
    )

    raw = _call_llm(llm_client, [{"role": "user", "content": prompt}], max_tokens=300)
    if raw is None or raw == "Not identified":
        return "Not identified"

    # Validate: only keep labels that exactly match the known set.
    parsed = [label.strip() for label in raw.split(";")]
    matched = sorted(label for label in parsed if label in valid_labels)
    return "; ".join(matched) if matched else "Not identified"


def identify_issues(
    complaint_text: str,
    llm_client: Any,
    llm_context: str = "",
    issues: list[str] | None = None,
    issues_patterns: list[tuple[re.Pattern | None, str, str]] | None = None,
) -> str:
    """Identify which legal issues from the canonical list appear in a complaint.

    Strategy:
    1. Isolate the legal claims section (counts, causes of action, etc.).
    2. Run regex / phrase matching against pre-compiled issues_patterns.  Legal
       claims are almost always cited by name, so this path is fast, free, and
       highly reliable.
    3. Fall back to the LLM only when regex finds nothing — typically cases
       where issues are argued implicitly rather than named explicitly.

    Returns a semicolon-separated, alphabetically sorted string of issue labels
    drawn from the supplied ``issues`` list, or 'Not identified' if none match
    or the call cannot be completed.
    """
    if not complaint_text or not issues:
        return "Not identified"

    claims_section = _extract_legal_claims_section(complaint_text)
    if not claims_section:
        return "Not identified"

    # Prefer count titles for precise matching; fall back to full claims section.
    # Extract titles from the claims section (not the full text) so that
    # back-references to counts in the prayer-for-relief are excluded.
    count_titles = _extract_count_titles(claims_section)
    matching_text = count_titles if count_titles else claims_section

    valid_labels = set(issues)

    # --- Primary path: regex + co-occurrence matching ---
    # APA issues are only reliable when explicitly named in a count title.
    # When claims_section is used as matching_text (no structured count titles),
    # exclude APA key texts to avoid "arbitrary and capricious" and similar
    # phrases firing in non-APA contexts (insurance, ERISA, etc.).
    _APA_PREFIX = "Administrative Procedure Act"
    has_count_titles = matching_text != claims_section
    regex_found: set[str] = set()
    if issues_patterns:
        if has_count_titles:
            regex_found |= _classify_issues_regex(matching_text, issues_patterns)
        else:
            body_safe = [(p, k, lbl) for p, k, lbl in issues_patterns
                         if not lbl.startswith(_APA_PREFIX)]
            regex_found |= _classify_issues_regex(claims_section, body_safe)
    # Co-occurrence detects issues present in claim bodies but not in titles
    # (e.g. "dormant Commerce Clause", "Tenth Amendment", amendment+clause pairs).
    # Also search up to 10 000 chars before the claims section to capture
    # constitutional context stated in background/intro paragraphs before COUNT I.
    _LOOKBACK = 10_000
    section_pos = complaint_text.find(claims_section[:80]) if claims_section else -1
    if section_pos > 0:
        pre_section = complaint_text[max(0, section_pos - _LOOKBACK) : section_pos]
        cooccurrence_text = pre_section + claims_section
    else:
        cooccurrence_text = claims_section
    cooccurrence_found = _classify_issues_cooccurrence(cooccurrence_text, valid_labels)
    combined = regex_found | cooccurrence_found
    if combined:
        return "; ".join(sorted(combined))

    # --- Fallback path: LLM (used only when regex finds nothing) ---
    if llm_client is None:
        return "Not identified"

    llm_claims = matching_text[:LLM_MAX_CHARS_ISSUES]
    issue_list = "\n".join(f"- {i}" for i in issues)
    context_block = f"{llm_context}\n" if llm_context else ""

    prompt = (
        "You are a legal document classifier. Read the legal claims section "
        "below and identify which issues from the following fixed list are raised.\n\n"
        f"{context_block}"
        f"Available issue labels:\n{issue_list}\n\n"
        "Instructions:\n"
        "1. Return ONLY labels from the list above, spelled exactly as shown.\n"
        "2. Sort the selected labels alphabetically.\n"
        "3. Separate labels with a semicolon and a space ('; ').\n"
        "4. If none of the listed issues are raised, return exactly: Not identified\n"
        "5. Do not include any explanation, preamble, or additional text.\n\n"
        f"Legal claims section:\n{llm_claims}"
    )

    raw = _call_llm(llm_client, [{"role": "user", "content": prompt}], max_tokens=500)
    if raw is None or raw == "Not identified":
        return "Not identified"

    parsed = [label.strip() for label in raw.split(";")]
    matched = sorted(label for label in parsed if label in valid_labels)
    return "; ".join(matched) if matched else "Not identified"


def identify_potential_impact(
    complaint_text: str,
    case_name: str,
    llm_client: Any,
    llm_context: str = "",
) -> str:
    """Generate a one-sentence description of the case's big-picture implications for healthcare.

    Returns a single sentence, or 'Not identified' if the LLM call fails.
    """
    if llm_client is None or not complaint_text:
        return "Not identified"

    excerpt = complaint_text[:LLM_MAX_CHARS_ANALYSIS]
    context_block = f"{llm_context}\n" if llm_context else ""

    prompt = (
        "You are a healthcare policy analyst reviewing a legal complaint. "
        "Based on the complaint excerpt below, write exactly ONE sentence describing "
        "the big-picture implications of this case for the broader healthcare landscape. "
        "Focus on systemic or policy-level consequences, not case-specific procedural outcomes.\n\n"
        f"Case: {case_name}\n\n"
        f"{context_block}"
        "Instructions:\n"
        "1. Return exactly one sentence.\n"
        "2. Do not include any preamble, label, or explanation.\n\n"
        f"Complaint excerpt:\n{excerpt}"
    )

    raw = _call_llm(llm_client, [{"role": "user", "content": prompt}], max_tokens=150)
    return raw if raw is not None else "Not identified"


def identify_why_this_matters(
    complaint_text: str,
    case_name: str,
    identified_issues: str,
    llm_client: Any,
    llm_context: str = "",
) -> str:
    """Generate a 2-3 sentence explanation of why the case matters for healthcare.

    Covers: the legal issues raised, the statute/provision at issue, and why the
    case is relevant to the healthcare landscape.  When the issues include
    'Administrative Procedure Act - Contrary to Constitutional Right', the
    underlying constitutional right is named in the narrative but is not listed
    as a standalone legal issue.

    Returns the generated text, or 'Not identified' if the LLM call fails.
    """
    if llm_client is None or not complaint_text:
        return "Not identified"

    excerpt = complaint_text[:LLM_MAX_CHARS_ANALYSIS]
    context_block = f"{llm_context}\n" if llm_context else ""

    issues_block = ""
    if identified_issues and identified_issues != "Not identified":
        issues_block = f"Identified legal issues: {identified_issues}\n\n"

    apa_const_note = ""
    if "Administrative Procedure Act - Contrary to Constitutional Right" in (identified_issues or ""):
        apa_const_note = (
            "IMPORTANT: One of the identified issues is "
            "'Administrative Procedure Act - Contrary to Constitutional Right'. "
            "In your response, identify and name the specific constitutional right or "
            "provision being violated (e.g., First Amendment free exercise, Fifth Amendment "
            "due process, etc.) and incorporate it into your explanation. Do NOT list it as "
            "a separate standalone legal issue.\n\n"
        )

    prompt = (
        "You are a healthcare policy analyst reviewing a legal complaint. "
        "Based on the complaint excerpt and identified issues below, write 2-3 sentences that: "
        "(1) describe the legal issues raised in the case and the specific statute, regulation, "
        "or constitutional provision at issue, and (2) explain why this case is relevant to the "
        "healthcare landscape. Be specific about what is being challenged.\n\n"
        f"Case: {case_name}\n\n"
        f"{issues_block}"
        f"{apa_const_note}"
        f"{context_block}"
        "Instructions:\n"
        "1. Return exactly 2-3 sentences.\n"
        "2. Do not include any preamble, label, or explanation.\n\n"
        f"Complaint excerpt:\n{excerpt}"
    )

    raw = _call_llm(llm_client, [{"role": "user", "content": prompt}], max_tokens=300)
    return raw if raw is not None else "Not identified"


# ---------------------------------------------------------------------------
# Module-level constants for claims-section parsing
# Shared by _extract_legal_claims_section, _extract_count_titles, and
# _classify_issues_cooccurrence so each function sees the same patterns.
# ---------------------------------------------------------------------------
_NUM = (
    r"(?:[ivxlcdm]{1,8}|\d{1,3}"
    r"|one|two|three|four|five|six|seven|eight|nine|ten"
    r"|first|second|third|fourth|fifth|sixth|seventh|eighth|ninth|tenth)"
)
_ORDINAL = r"(?:first|second|third|fourth|fifth|sixth|seventh|eighth|ninth|tenth)"

_COUNT_PAT = re.compile(rf"(?i)\bcount\s+(?:no\.?\s*)?{_NUM}\b")
_CLAIM_PAT = re.compile(rf"(?i)\b{_ORDINAL}\s+claim\s+for\s+relief\b")

_TIER1_PATTERNS: list[str] = [
    rf"\bcount\s+(?:no\.?\s*)?{_NUM}\b",
    rf"\bclaim\s+(?:no\.?\s*)?{_NUM}\b",
    r"\bfirst\s+cause\s+of\s+action\b",
    rf"\b{_ORDINAL}\s+claim\s+for\s+relief\b",
]
_TIER2_PATTERNS: list[str] = [
    r"\bcauses?\s+of\s+action\b",
    r"\bclaims?\s+for\s+relief\b",
]
_END_SIGNALS: list[str] = [
    "prayer for relief",
    "wherefore",
    "demand for relief",
    "demand for jury trial",
    "relief requested",
    "relief sought",
    "respectfully submitted",
    "certificate of service",
]


def _is_toc_line(text_lower: str, match_pos: int) -> bool:
    """Return True if the match sits on a table-of-contents line (dots + page number).

    Checks 200 chars past the match start to handle labels that wrap across
    lines before the dot-leaders appear (e.g. 'COUNT TWO Unlawful Acquisition
    Section Seven ...\\nAct ......... 70').
    """
    line_start = text_lower.rfind("\n", 0, match_pos) + 1
    look_ahead_end = min(len(text_lower), match_pos + 200)
    return bool(re.search(r"\.{4,}", text_lower[line_start:look_ahead_end]))


def _extract_legal_claims_section(full_text: str) -> str:
    """Extract the legal claims section from a complaint.

    Searches for the first occurrence of standard claim-header patterns
    (Count I/1/One, Claim No. 1, First Cause of Action, etc.) and returns the
    text from that point up to the prayer-for-relief/wherefore section.

    Falls back to the full complaint text when no explicit claims section is
    found, so the LLM can still identify issues from complaints that integrate
    claims narratively rather than enumerating them under headings.
    """
    if not full_text:
        return ""

    text = full_text.replace(" ", " ")
    lower = text.lower()

    def _find_start(patterns: list[str]) -> int | None:
        best: int | None = None
        for pat in patterns:
            for m in re.finditer(pat, lower):
                if _is_toc_line(lower, m.start()):
                    continue
                if best is None or m.start() < best:
                    best = m.start()
                break
        return best

    # Tier-1: explicitly numbered/ordinal headers.
    # Tier-2: generic fallbacks used only when Tier-1 finds nothing.
    start_idx = _find_start(_TIER1_PATTERNS)
    if start_idx is None:
        start_idx = _find_start(_TIER2_PATTERNS)

    if start_idx is None:
        return text  # no explicit section; let the LLM scan the full text

    section = text[start_idx:]
    lower_section = section.lower()

    end_idx = len(section)
    for sig in _END_SIGNALS:
        for em in re.finditer(re.escape(sig), lower_section):
            i = em.start()
            if i <= 0 or i >= end_idx:
                continue
            # Only treat the signal as a section-end marker when it appears at the
            # start of a line (possibly indented), not embedded in a sentence.
            # e.g. "the relief requested below would remedy" must NOT terminate the
            # section, but a standalone "RELIEF REQUESTED" header should.
            line_start = lower_section.rfind("\n", 0, i) + 1
            if lower_section[line_start:i].strip() == "":
                end_idx = i
            break  # only inspect the first occurrence per signal

    return section[:end_idx].strip()


def _extract_count_titles(full_text: str) -> str:
    """Extract only the heading text of each count/claim from a complaint.

    For each non-ToC count header (Count I, COUNT TWO, etc.) collects the
    header line plus any descriptive title lines that follow, stopping before
    the first paragraph-body line (identified by a leading paragraph number
    like ``39.`` or ``104.``).

    Returns all titles joined by newlines, suitable for focused issue matching.
    Returns an empty string when no structured counts are detected.
    """
    if not full_text:
        return ""

    text = full_text.replace(" ", " ")

    # Collect header match positions from both patterns; sort by position so
    # titles appear in document order regardless of which pattern matched.
    header_matches = sorted(
        list(_COUNT_PAT.finditer(text)) + list(_CLAIM_PAT.finditer(text)),
        key=lambda m: m.start(),
    )

    titles = []
    for m in header_matches:
        # Skip table-of-contents lines: check the line AND the next 200 chars
        # for 4+ consecutive dots (dots may wrap to the line after the count label).
        line_start = text.rfind("\n", 0, m.start()) + 1
        look_ahead_end = min(len(text), m.start() + 200)
        if re.search(r"\.{4,}", text[line_start:look_ahead_end]):
            continue

        # Grab text from the count/claim header up to the first paragraph-body
        # line. Paragraph body lines begin with a number and period (e.g. "39."
        # or "225.").
        snippet = text[m.start() : m.start() + 600]
        para_m = re.search(r"\n\s*\d{1,4}\.", snippet)
        title_text = snippet[: para_m.start()] if para_m else snippet[:300]

        title_text = re.sub(r"\s+", " ", title_text).strip()
        if title_text:
            titles.append(title_text)

    return "\n".join(titles)


# Co-occurrence rules for constitutional and statutory claims.
# Each entry: (anchor phrase, set of clause/context keywords, canonical issue label).
# Applied per individual claim body: if the anchor phrase AND any keyword both
# appear anywhere within the same claim body, the issue is added.
# This lets body text (not just count titles) identify issues when the title is
# too generic (e.g. "Fourteenth Amendment Claim") or when the canonical phrase
# only appears in the body (e.g. "dormant Commerce Clause").
_COOCCURRENCE_RULES: list[tuple[str, frozenset[str], str]] = [
    # --- Fourteenth Amendment ---
    ("fourteenth amendment", frozenset({"due process", "void for vagueness", "vagueness"}), "Fourteenth Amendment - Due Process Clause"),
    ("fourteenth amendment", frozenset({"equal protection"}), "Fourteenth Amendment - Equal Protection Clause"),
    # --- First Amendment ---
    ("first amendment", frozenset({"free speech", "freedom of speech", "speech", "expression", "expressive"}), "First Amendment - Free Speech Clause"),
    # Free Exercise: NOT in co-occurrence — Establishment Clause claim bodies routinely
    # cite "Free Exercise Clause" as a contrast, which would cause false positives.
    # Free Exercise is identified via key-text matching on count titles instead.
    ("first amendment", frozenset({"establishment clause", "establishment of religion", "lemon test"}), "First Amendment - Establishment Clause"),
    # Overbreadth: require "overbreadth doctrine" or "facial overbreadth" — the bare
    # word "overbroad" appears too often as a general adjective in free-speech claims.
    ("first amendment", frozenset({"overbreadth doctrine", "facial overbreadth"}), "First Amendment - Overbreadth Doctrine"),
    # --- Fifth Amendment ---
    ("fifth amendment", frozenset({"due process"}), "Fifth Amendment - Due Process Clause"),
    ("fifth amendment", frozenset({"taking", "just compensation", "takings"}), "Fifth Amendment - Takings Clause"),
    ("fifth amendment", frozenset({"equal protection"}), "Fifth Amendment - Equal Protection Clause"),
    # --- Dormant Commerce Clause ---
    # Counts often title themselves "Commerce Clause" without "dormant"; body text
    # contains "dormant Commerce Clause", "extraterritorial", or "Pike balancing".
    ("commerce clause", frozenset({"dormant", "extraterritorial", "pike", "burdening interstate"}), "Dormant Commerce Clause"),
    # --- Tenth Amendment ---
    # Specific enough that any mention in a claim body is likely substantive.
    ("tenth amendment", frozenset({"reserved", "state", "states", "powers"}), "Tenth Amendment"),
]

_CLAIM_SPLIT_PAT = re.compile(
    r"(?i)(?:"
    r"\bcount\s+(?:no\.?\s*)?(?:[ivxlcdm]{1,8}|\d{1,3}|one|two|three|four|five|six|seven|eight|nine|ten)\b"
    r"|"
    r"\b(?:first|second|third|fourth|fifth|sixth|seventh|eighth|ninth|tenth)\s+claim\s+for\s+relief\b"
    r")"
)


def _classify_issues_cooccurrence(claims_section: str, valid_labels: set[str]) -> set[str]:
    """Detect constitutional amendment issues via per-claim co-occurrence.

    Splits the claims section into individual claim bodies and, for each body,
    checks whether an amendment phrase and a clause keyword appear together
    anywhere in that body — regardless of proximity. This correctly identifies
    claims like 'Fourteenth Amendment... depriving... without due process of law'
    where the amendment name and clause keyword are separated by many words.
    """
    if not claims_section:
        return set()

    # Split into individual claim bodies.  The preamble (text before the first
    # count header) is included as its own body so that legal-background sections
    # that name the constitutional framework — e.g. "the Tenth Amendment provides
    # that powers not delegated to the United States are reserved to the States"
    # — are also matched even when the count titles do not spell out every issue.
    positions = [m.start() for m in _CLAIM_SPLIT_PAT.finditer(claims_section)]
    if not positions:
        bodies = [claims_section]
    else:
        positions.append(len(claims_section))
        bodies = [claims_section[: positions[0]]]  # preamble before first count
        bodies += [claims_section[positions[i] : positions[i + 1]] for i in range(len(positions) - 1)]

    found: set[str] = set()
    for body in bodies:
        body_lower = body.lower()
        for amendment_key, clause_keywords, issue_label in _COOCCURRENCE_RULES:
            if issue_label not in valid_labels:
                continue
            if amendment_key in body_lower:
                if any(kw in body_lower for kw in clause_keywords):
                    found.add(issue_label)

    return found


def _extract_relief_requests(full_text: str) -> list[str]:
    """Extract and split prayer-for-relief requests from complaint text."""
    if not full_text:
        return []
    text = full_text.replace("\u00a0", " ")
    lower = text.lower()

    start_idx = -1
    for signal in (
        RELIEF_SIGNAL,
        "prayer for relief",
        "prayer for judgment",
        "relief requested",
        "requests that this court",
        "plaintiff seeks the following relief",
        "demand for relief"
    ):
        start_idx = lower.find(signal)
        if start_idx != -1:
            break
    if start_idx == -1:
        return []

    section = text[start_idx:]
    lower_section = section.lower()
    end_candidates = [i for i in (
        lower_section.find("dated"),
        lower_section.find("respectfully submitted"),
        lower_section.find("demand for jury trial"),
        lower_section.find("certificate of service"),
    ) if i != -1]
    if end_candidates:
        section = section[:min(end_candidates)]

    items = [
        s.strip()
        for s in re.split(r"(?m)(?=^\s*(?:\d+|[A-Za-z])[\.\)])", section)
        if s.strip()
    ]
    return items


def extract_case_attributes(
    docket_text: str,
    complaint_text: str,
    pdf_path: Path,
    district_court_code_text: str = "",
    llm_client: Any = None,
    llm_context: str = "",
    llm_issues_context: str = "",
    llm_impact_context: str = "",
    llm_why_context: str = "",
    issues: list[str] | None = None,
    issues_patterns: list[tuple[re.Pattern | None, str, str]] | None = None,
) -> dict[str, str]:
    """Extract all structured case attributes from docket and complaint text.

    Returns a dict with keys: ``case_name``, ``docket_number``, ``date_filed``,
    ``court``, ``division``, ``judge``, ``president``, ``plaintiffs``,
    ``intervenor_plaintiffs``, ``lead_counsel_plaintiffs``, ``defendants``,
    ``intervenor_defendants``, ``lead_counsel_defendants``, ``goals``,
    ``issues``, ``potential_impact``, ``why_this_matters``.
    """

    court_info = identify_court(docket_text, district_court_code_text)
    plaintiffs_info = identify_plaintiffs(docket_text)
    defendants_info = identify_defendants(docket_text, pdf_path)
    title = identify_title(docket_text)
    lead_p, lead_d = identify_lead_counsel(pdf_path, title)
    judge = identify_judge(docket_text)

    case_name = identify_case_name(docket_text)
    case_name = apply_et_al_to_case_name(
        case_name,
        count_named_parties(plaintiffs_info["Plaintiffs"]),
        count_named_parties(defendants_info["Defendants"]),
    )

    # If the first plaintiff is a qui tam "ex rel." party, inject that portion
    # into the plaintiff side of the case name (before any " et al.").
    first_plaintiff = plaintiffs_info["Plaintiffs"].split("; ")[0]
    ex_rel_m = re.search(r"\s+ex\s+rel\.\s+(.+)", first_plaintiff, re.IGNORECASE)
    if ex_rel_m:
        relator_full_name = ex_rel_m.group(1).strip()
        relator_last_name = relator_full_name.split()[-1]
        ex_rel_suffix = f" ex rel. {relator_last_name}"
        base_name = first_plaintiff[: ex_rel_m.start()]
        cn_parts = _CASE_NAME_VS_SPLIT.split(case_name, maxsplit=1)
        if len(cn_parts) >= 2:
            left_side = cn_parts[0].strip()
            left_updated = re.sub(
                re.escape(base_name),
                base_name + ex_rel_suffix,
                left_side,
                count=1,
                flags=re.IGNORECASE,
            )
            if left_updated != left_side:
                case_name = f"{left_updated} v. {cn_parts[1].strip()}"

    all_defendants = "; ".join(filter(None, [
        defendants_info["Defendants"],
        defendants_info["Intervenor Defendants"],
    ]))
    is_agency = _is_agency_defendant(all_defendants)

    # Compute issues first (usually the fast regex path) so it can be passed
    # to identify_why_this_matters, then run the 3 remaining LLM calls concurrently.
    text = complaint_text or ""
    complaint_readable = _looks_like_substantive_text(text)
    identified_issues = identify_issues(
        text,
        llm_client=llm_client,
        llm_context=llm_issues_context,
        issues=issues,
        issues_patterns=issues_patterns,
    )

    # If a complaint PDF was provided but yielded no readable text (e.g. scanned),
    # fall back to the docket's Cause and Nature of suit fields and flag the
    # result for manual review.  When no complaint was provided at all (text=""),
    # leave identified_issues as-is rather than showing a misleading message.
    if text and not complaint_readable:
        cause_m = re.search(r"\*\*Cause:\*\*\s*([^*\n]+)", docket_text, re.IGNORECASE)
        nos_m = re.search(r"\*\*Nature of suit:\*\*\s*([^*\n]+)", docket_text, re.IGNORECASE)
        docket_cause_text = " ".join(filter(None, [
            cause_m.group(1).strip() if cause_m else "",
            nos_m.group(1).strip() if nos_m else "",
        ]))
        if docket_cause_text:
            docket_issues = identify_issues(
                docket_cause_text,
                llm_client=None,
                issues=issues,
                issues_patterns=issues_patterns,
            )
            if docket_issues and docket_issues != "Not identified":
                identified_issues = (
                    "The Complaint pdf provided is not readable. Based on the docket, "
                    "issues include the following, but review the complaint to confirm: "
                    + docket_issues
                )
            else:
                identified_issues = (
                    "The Complaint pdf provided is not readable. "
                    "Issues could not be identified from the docket."
                )

    with ThreadPoolExecutor(max_workers=3) as pool:
        goals_future = pool.submit(
            identify_goals, text, llm_client, llm_context, is_agency
        )
        impact_future = pool.submit(
            identify_potential_impact, text, case_name, llm_client, llm_impact_context
        )
        why_future = pool.submit(
            identify_why_this_matters, text, case_name, identified_issues, llm_client, llm_why_context
        )
        goals_result  = goals_future.result()
        impact_result = impact_future.result()
        why_result    = why_future.result()

    if text and not complaint_readable and goals_result == "Not identified":
        goals_result = "Not identified - complaint is not readable"

    case_attributes = {
        "case_name": case_name,
        "docket_number": identify_docket_number(docket_text),
        "date_filed": identify_date_filed(docket_text),
        "plaintiffs": plaintiffs_info["Plaintiffs"],
        "intervenor_plaintiffs": plaintiffs_info["Intervenor Plaintiffs"],
        "lead_counsel_plaintiffs": lead_p,
        "defendants": defendants_info["Defendants"],
        "intervenor_defendants": defendants_info["Intervenor Defendants"],
        "lead_counsel_defendants": lead_d,
        "court": court_info["court"],
        "division": court_info["division"],
        "judge": judge,
        "president": identify_president(judge),
        "goals": goals_result,
        "issues": identified_issues,
        "potential_impact": impact_result,
        "why_this_matters": why_result,
    }
    return case_attributes


def _unique_path(path: Path) -> Path:
    """Return path unchanged if it doesn't exist; otherwise append (1), (2), … until unique."""
    if not path.exists():
        return path
    counter = 1
    while True:
        candidate = path.with_stem(f"{path.stem} ({counter})")
        if not candidate.exists():
            return candidate
        counter += 1


def write_to_excel(script_dir: Path, today_str: str, headers: list[str], rows: list[list[object]]) -> Path:
    """Write rows to an Excel workbook and return the saved .xlsx path."""
    xlsx_path = _unique_path(script_dir / f"Tracker Data Summary {today_str}.xlsx")

    wb = Workbook()
    ws = wb.active
    ws.title = "District Court Cases"

    ws.append(headers)
    for row in rows:
        ws.append(row)

    wb.save(xlsx_path)

    return xlsx_path


def write_to_csv(script_dir: Path, today_str: str, headers: list[str], rows: list[list[object]]) -> Path:
    """Write rows to a CSV file and return the saved path."""
    csv_path = _unique_path(script_dir / f"Tracker Data Summary {today_str}.csv")

    with csv_path.open("w", newline="", encoding="utf-8") as f:
        writer = csv.writer(f)
        writer.writerow(headers)
        writer.writerows(rows)

    return csv_path


def write_to_json(script_dir: Path, today_str: str, headers: list[str], rows: list[list[object]]) -> Path:
    """Write rows to a JSON file and return the saved path.

    Each case is a JSON object whose keys are the column headers, making
    the output easy to consume programmatically without parsing a CSV or
    opening an Excel workbook.
    """
    json_path = _unique_path(script_dir / f"Tracker Data Summary {today_str}.json")

    records = [dict(zip(headers, row)) for row in rows]
    with json_path.open("w", encoding="utf-8") as f:
        json.dump(records, f, indent=2, ensure_ascii=False)

    return json_path


def main(data_dir: Path | None = None) -> None:
    """ Obtain all the necessary attributes for each case where the docket is found in the
    TrialCourtDockets folder and where the complaint is found in the TrialCourtComplaints folder
    and write these attributes to an Excel file
    """
    script_dir = (data_dir if data_dir is not None else Path(__file__).resolve().parent.parent).resolve()
    docket_target = script_dir / DOCKET_FOLDER_NAME
    complaint_target = script_dir / COMPLAINT_FOLDER_NAME

    # pymupdf4llm is not thread-safe; all PDF reads run sequentially.
    try:
        docket_text_by_name = read_pdf_files_in_folder(docket_target)
    except NotADirectoryError:
        print(f"Not a folder: {docket_target}", file=sys.stderr)
        sys.exit(1)

    try:
        complaint_text_by_name = read_pdf_files_in_folder(complaint_target)
    except NotADirectoryError:
        print(f"Not a folder: {complaint_target}", file=sys.stderr)
        sys.exit(1)

    district_court_code_text = read_pdf_text(script_dir / DISTRICT_COURT_CODE_FILE)

    goals_mapping_path = script_dir / GOALS_MAPPING_CSV_FILE
    goals_mapping_dict = read_mapping_csv(goals_mapping_path)

    llm_client = _get_openrouter_client()
    if llm_client is None:
        print("OPENROUTER_API_KEY not set or openai package not installed — "
              "using regex classification.", file=sys.stderr)

    goals_examples = _load_goals_examples(script_dir)
    llm_context = _build_llm_goals_context(goals_examples, goals_mapping_dict)

    issues = _load_issues(script_dir)
    issues_mapping = read_mapping_csv(script_dir / ISSUES_MAPPING_CSV_FILE)
    issues_patterns = _compile_issues_patterns(issues_mapping, set(issues))
    issues_examples = _load_legal_issues_examples(script_dir)
    llm_issues_context = _build_llm_issues_context(issues_examples)

    analysis_examples = _load_analysis_examples(script_dir)
    llm_impact_context = _build_llm_impact_context(analysis_examples)
    llm_why_context = _build_llm_why_context(analysis_examples)

    # Obtain case attributes for each complaint and print them,
    # while also collecting rows for the Excel/CSV summary
    headers = [
        "File Name",
        "Case Name",
        "Docket Number",
        "Date Filed",
        "Court",
        "Division",
        "Judge",
        "President Who Appointed Judge",
        "Plaintiffs",
        "Intervenor Plaintiffs",
        "Lead Counsel for Plaintiff(s)",
        "Defendants",
        "Intervenor Defendants",
        "Lead Counsel for Defendant(s)",
        "Goals",
        "Issues",
        "Potential Impact",
        "Why This Matters",
    ]
    rows: list[list[object]] = []
    for name, text in docket_text_by_name.items():
        resolved_complaint_name, complaint_text = complaint_text_for_docket_pdf(
            name, complaint_text_by_name
        )
        # Missing complaint: show "None" only (no stderr noise). Pairing uses the docket
        # PDF basename (e.g. '…, Docket.pdf'), not the case title extracted from the PDF.
        complaint_file_name = (
            resolved_complaint_name if complaint_text is not None else "None"
        )

        case_attributes = extract_case_attributes(
            text,
            complaint_text or "",
            docket_target / name,
            district_court_code_text=district_court_code_text,
            llm_client=llm_client,
            llm_context=llm_context,
            llm_issues_context=llm_issues_context,
            llm_impact_context=llm_impact_context,
            llm_why_context=llm_why_context,
            issues=issues,
            issues_patterns=issues_patterns,
        )

        rows.append(
            [
                name,
                case_attributes["case_name"],
                case_attributes["docket_number"],
                case_attributes["date_filed"],
                case_attributes["court"],
                case_attributes["division"],
                case_attributes["judge"],
                case_attributes["president"],
                case_attributes["plaintiffs"],
                case_attributes["intervenor_plaintiffs"],
                case_attributes["lead_counsel_plaintiffs"],
                case_attributes["defendants"],
                case_attributes["intervenor_defendants"],
                case_attributes["lead_counsel_defendants"],
                case_attributes["goals"],
                case_attributes["issues"],
                case_attributes["potential_impact"],
                case_attributes["why_this_matters"],
            ]
        )

    # Write summary files at the end of main()
    today_str = date.today().strftime("%m.%d.%y")
    write_to_excel(script_dir, today_str, headers, rows)
    write_to_csv(script_dir, today_str, headers, rows)
    write_to_json(script_dir, today_str, headers, rows)
