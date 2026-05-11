"""
oneill_tracker — public API for the O'Neill Healthcare Litigation Tracker.

Quick start (standard data directory):
---------------------------------------
    from pathlib import Path
    import oneill_tracker

    config = oneill_tracker.CaseAnalyzerConfig.from_directory(
        Path("/path/to/ONeillTrackerData"),
        api_key="your-openrouter-key",   # or set OPENROUTER_API_KEY env var
    )
    result = oneill_tracker.analyze_case(
        docket_pdf="path/to/Case, Docket.pdf",
        complaint_pdf="path/to/Case, Complaint.pdf",
        config=config,
    )
    print(result["case_name"])
    print(result["issues"])

Custom domain (bring your own mappings):
-----------------------------------------
    config = oneill_tracker.CaseAnalyzerConfig(
        goals_mapping={"seeks injunction": "Block defendant action"},
        issues=["Breach of Contract", "Fraud"],
        issues_mapping={"breach of contract": "Breach of Contract"},
        llm_client=my_openai_compatible_client,
    )
    result = oneill_tracker.analyze_case(
        docket_pdf="path/to/docket.pdf",
        complaint_pdf="path/to/complaint.pdf",
        config=config,
    )
"""

from __future__ import annotations

import os
from dataclasses import dataclass, field
from pathlib import Path
from typing import Any

from . import pipeline as _pipeline


# ---------------------------------------------------------------------------
# Low-level public re-exports
# Callers can use individual pipeline steps without going through
# the high-level analyze_case() wrapper.
# ---------------------------------------------------------------------------
extract_case_attributes   = _pipeline.extract_case_attributes
identify_goals            = _pipeline.identify_goals
identify_issues           = _pipeline.identify_issues
identify_potential_impact = _pipeline.identify_potential_impact
identify_why_this_matters = _pipeline.identify_why_this_matters
read_pdf_text             = _pipeline.read_pdf_text
read_pdf_files_in_folder  = _pipeline.read_pdf_files_in_folder
read_mapping_csv          = _pipeline.read_mapping_csv


# ---------------------------------------------------------------------------
# CaseAnalyzerConfig
# ---------------------------------------------------------------------------
@dataclass
class CaseAnalyzerConfig:
    """All configuration needed to run case analysis.

    Build with ``CaseAnalyzerConfig.from_directory()`` to load the
    standard data files, or construct directly to supply your own
    mappings for a different legal domain.

    Parameters
    ----------
    goals_mapping:
        ``{phrase: goal_label}`` dict used to build the goals prompt context.
    issues:
        Ordered list of canonical issue labels (the allowed output values).
    issues_mapping:
        ``{phrase: issue_label}`` dict used by the regex classifier.
        Patterns are compiled automatically on first use if ``issues_patterns``
        is not supplied.
    issues_patterns:
        Pre-compiled patterns built from ``issues_mapping``.  Populated
        automatically by ``from_directory()`` and ``__post_init__``.
    district_court_code_text:
        Full text of 28 U.S.C. Ch. 5 used for court division look-ups.
        Leave empty to skip division identification.
    llm_goals_context:
        Few-shot context block injected into the goals classification prompt.
    llm_issues_context:
        Few-shot context block injected into the issues classification prompt.
    llm_impact_context:
        Few-shot context block injected into the potential impact prompt.
    llm_why_context:
        Few-shot context block injected into the why-this-matters prompt.
    llm_client:
        An OpenAI-compatible client.  If ``None``, LLM-dependent fields
        return "Not identified".
    """

    goals_mapping:            dict[str, str] = field(default_factory=dict)
    issues:                   list[str]       = field(default_factory=list)
    issues_mapping:           dict[str, str]  = field(default_factory=dict)
    issues_patterns:          list            = field(default_factory=list)
    district_court_code_text: str             = ""
    llm_goals_context:        str             = ""
    llm_issues_context:       str             = ""
    llm_impact_context:       str             = ""
    llm_why_context:          str             = ""
    llm_client:               Any             = None

    def __post_init__(self) -> None:
        # Auto-compile regex patterns when a mapping was supplied without patterns.
        if self.issues_mapping and not self.issues_patterns and self.issues:
            self.issues_patterns = _pipeline._compile_issues_patterns(
                self.issues_mapping, set(self.issues)
            )

    @classmethod
    def from_directory(
        cls,
        data_dir: str | Path,
        api_key: str = "",
        llm_client: Any = None,
    ) -> "CaseAnalyzerConfig":
        """Load all configuration from a data directory.

        Reads the standard files that ship with the tracker:
        ``GoalsMapping.csv``, ``GoalsExamples.xlsx``, ``Issues.csv``,
        ``IssuesMapping.csv``, ``LegalIssuesExamples.xlsx``,
        ``AnalysisExamples.xlsx``, and ``28 USC Ch5 District Courts.pdf``.

        Parameters
        ----------
        data_dir:
            Path to the directory containing the data files.
        api_key:
            OpenRouter API key.  Falls back to the ``OPENROUTER_API_KEY``
            environment variable if omitted.
        llm_client:
            Pass an already-constructed client to skip API key handling.
        """
        data_dir = Path(data_dir)

        # Goals
        goals_mapping = _pipeline.read_mapping_csv(
            data_dir / _pipeline.GOALS_MAPPING_CSV_FILE
        )
        goals_examples = _pipeline._load_goals_examples(data_dir)
        llm_goals_context = _pipeline._build_llm_goals_context(
            goals_examples, goals_mapping
        )

        # Issues
        issues = _pipeline._load_issues(data_dir)
        issues_mapping = _pipeline.read_mapping_csv(
            data_dir / _pipeline.ISSUES_MAPPING_CSV_FILE
        )
        issues_patterns = _pipeline._compile_issues_patterns(
            issues_mapping, set(issues)
        )
        issues_examples = _pipeline._load_legal_issues_examples(data_dir)
        llm_issues_context = _pipeline._build_llm_issues_context(issues_examples)

        # Analysis narrative
        analysis_examples = _pipeline._load_analysis_examples(data_dir)
        llm_impact_context = _pipeline._build_llm_impact_context(analysis_examples)
        llm_why_context    = _pipeline._build_llm_why_context(analysis_examples)

        # District court reference document
        district_path = data_dir / _pipeline.DISTRICT_COURT_CODE_FILE
        district_court_code_text = (
            _pipeline.read_pdf_text(district_path)
            if district_path.exists()
            else ""
        )

        # LLM client
        if llm_client is None:
            key = api_key or os.environ.get("OPENROUTER_API_KEY", "")
            if key:
                os.environ["OPENROUTER_API_KEY"] = key
            llm_client = _pipeline._get_openrouter_client()

        return cls(
            goals_mapping=goals_mapping,
            issues=issues,
            issues_mapping=issues_mapping,
            issues_patterns=issues_patterns,
            district_court_code_text=district_court_code_text,
            llm_goals_context=llm_goals_context,
            llm_issues_context=llm_issues_context,
            llm_impact_context=llm_impact_context,
            llm_why_context=llm_why_context,
            llm_client=llm_client,
        )


# ---------------------------------------------------------------------------
# analyze_case — high-level single-case entry point
# ---------------------------------------------------------------------------
def analyze_case(
    docket_pdf: str | Path,
    complaint_pdf: str | Path | None,
    config: CaseAnalyzerConfig,
) -> dict:
    """Extract structured attributes from a single case.

    Parameters
    ----------
    docket_pdf:
        Path to the PACER docket PDF.
    complaint_pdf:
        Path to the complaint PDF, or ``None`` if unavailable.
    config:
        A ``CaseAnalyzerConfig`` instance.

    Returns
    -------
    dict
        Keys: ``case_name``, ``docket_number``, ``date_filed``, ``court``,
        ``division``, ``judge``, ``president``, ``plaintiffs``,
        ``intervenor_plaintiffs``, ``lead_counsel_plaintiffs``,
        ``defendants``, ``intervenor_defendants``,
        ``lead_counsel_defendants``, ``goals``, ``issues``,
        ``potential_impact``, ``why_this_matters``.
    """
    docket_path    = Path(docket_pdf)
    docket_text    = _pipeline.read_pdf_text(docket_path)
    complaint_text = (
        _pipeline.read_pdf_text(Path(complaint_pdf)) if complaint_pdf else ""
    )

    return _pipeline.extract_case_attributes(
        docket_text=docket_text,
        complaint_text=complaint_text,
        pdf_path=docket_path,
        district_court_code_text=config.district_court_code_text,
        llm_client=config.llm_client,
        llm_context=config.llm_goals_context,
        llm_issues_context=config.llm_issues_context,
        llm_impact_context=config.llm_impact_context,
        llm_why_context=config.llm_why_context,
        issues=config.issues,
        issues_patterns=config.issues_patterns,
    )


# ---------------------------------------------------------------------------
# Public API surface
# ---------------------------------------------------------------------------
__all__ = [
    # High-level
    "CaseAnalyzerConfig",
    "analyze_case",
    # Mid-level (individual pipeline steps)
    "extract_case_attributes",
    "identify_goals",
    "identify_issues",
    "identify_potential_impact",
    "identify_why_this_matters",
    # Utilities
    "read_pdf_text",
    "read_pdf_files_in_folder",
    "read_mapping_csv",
]
