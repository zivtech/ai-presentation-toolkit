"""
Template Diagnostics

Pre-flight checks that catch template problems before migration starts.
Checks for missing placeholders, broken references, slide count mismatches, etc.
"""

import os
import re
import sys
import zipfile
from dataclasses import dataclass, field
from enum import Enum
from pathlib import Path
from typing import Any, Dict, List, Optional, Union

from lxml import etree

from .pptx_utils import NSMAP, EMU_PER_INCH, find_placeholder, find_shape_by_name


class Severity(str, Enum):
    """Severity of a diagnostic issue."""
    error = "error"
    warning = "warning"
    info = "info"


@dataclass
class DiagnosticIssue:
    """A single diagnostic finding."""
    code: str
    severity: Severity
    message: str
    slide_index: Optional[int] = None
    category: str = ""
    detail: str = ""


@dataclass
class DiagnosticReport:
    """Aggregated diagnostic results."""
    issues: List[DiagnosticIssue] = field(default_factory=list)
    slide_count: int = 0
    layout_count: int = 0
    master_count: int = 0

    @property
    def errors(self) -> List[DiagnosticIssue]:
        return [i for i in self.issues if i.severity == Severity.error]

    @property
    def warnings(self) -> List[DiagnosticIssue]:
        return [i for i in self.issues if i.severity == Severity.warning]

    @property
    def has_blocking_issues(self) -> bool:
        return len(self.errors) > 0

    def print_report(self, file=None) -> None:
        """Print a human-readable diagnostic report."""
        out = file or sys.stdout

        print("=" * 60, file=out)
        print("TEMPLATE DIAGNOSTIC REPORT", file=out)
        print("=" * 60, file=out)
        print(f"Slides: {self.slide_count}  |  Layouts: {self.layout_count}  |  Masters: {self.master_count}", file=out)
        print(f"Issues: {len(self.errors)} errors, {len(self.warnings)} warnings, "
              f"{len(self.issues) - len(self.errors) - len(self.warnings)} info", file=out)
        print("-" * 60, file=out)

        for issue in self.issues:
            prefix = {
                Severity.error: "ERROR  ",
                Severity.warning: "WARN   ",
                Severity.info: "INFO   ",
            }[issue.severity]

            slide_str = f" [slide {issue.slide_index}]" if issue.slide_index is not None else ""
            print(f"  {prefix} {issue.code}{slide_str}: {issue.message}", file=out)
            if issue.detail:
                print(f"         {issue.detail}", file=out)

        print("-" * 60, file=out)
        if self.has_blocking_issues:
            print("RESULT: BLOCKING errors found. Fix before migration.", file=out)
        elif self.warnings:
            print("RESULT: Warnings found. Migration will proceed but some content may be affected.", file=out)
        else:
            print("RESULT: Template looks good!", file=out)

    def to_dict(self) -> Dict[str, Any]:
        """Convert to JSON-serializable dict."""
        return {
            "slide_count": self.slide_count,
            "layout_count": self.layout_count,
            "master_count": self.master_count,
            "error_count": len(self.errors),
            "warning_count": len(self.warnings),
            "has_blocking_issues": self.has_blocking_issues,
            "issues": [
                {
                    "code": i.code,
                    "severity": i.severity.value,
                    "message": i.message,
                    "slide_index": i.slide_index,
                    "category": i.category,
                    "detail": i.detail,
                }
                for i in self.issues
            ],
        }


# ============================================================
# DIAGNOSTIC CHECKS
# ============================================================

def _check_file_exists(template_path: Path, report: DiagnosticReport) -> bool:
    """TMPL-001: Check that the template file exists."""
    if not template_path.exists():
        report.issues.append(DiagnosticIssue(
            code="TMPL-001",
            severity=Severity.error,
            message="Template file not found",
            category="file",
            detail=str(template_path),
        ))
        return False
    return True


def _check_valid_zip(template_path: Path, report: DiagnosticReport) -> bool:
    """TMPL-002: Check that the file is a valid ZIP/PPTX."""
    if not zipfile.is_zipfile(template_path):
        report.issues.append(DiagnosticIssue(
            code="TMPL-002",
            severity=Severity.error,
            message="Not a valid ZIP/PPTX file",
            category="file",
            detail=str(template_path),
        ))
        return False
    return True


def _check_required_parts(zf: zipfile.ZipFile, report: DiagnosticReport) -> bool:
    """TMPL-003: Check for required OOXML parts."""
    required = [
        "[Content_Types].xml",
        "ppt/presentation.xml",
    ]
    names = zf.namelist()
    ok = True
    for part in required:
        if part not in names:
            report.issues.append(DiagnosticIssue(
                code="TMPL-003",
                severity=Severity.error,
                message=f"Missing required OOXML part: {part}",
                category="structure",
            ))
            ok = False
    return ok


def _check_config_slide_indices(
    zf: zipfile.ZipFile, config: Any, report: DiagnosticReport
) -> None:
    """TMPL-004: Check that config references valid slide indices."""
    if config is None:
        return

    names = zf.namelist()
    slide_files = [n for n in names if re.match(r"ppt/slides/slide\d+\.xml$", n)]
    max_slide = 0
    for sf in slide_files:
        m = re.search(r"slide(\d+)", sf)
        if m:
            max_slide = max(max_slide, int(m.group(1)))

    if hasattr(config, "get_all_slide_indices"):
        catalog = config.get_all_slide_indices()
        for category, indices in catalog.items():
            for idx in indices:
                if idx < 0 or idx > max_slide:
                    report.issues.append(DiagnosticIssue(
                        code="TMPL-004",
                        severity=Severity.warning,
                        message=f"Config references slide index {idx} for '{category}' but template has {max_slide} slides",
                        category="config",
                    ))


def _check_slide_placeholders(
    zf: zipfile.ZipFile, report: DiagnosticReport
) -> None:
    """TMPL-010/011/012: Check each slide for TITLE, BODY, PICTURE placeholders."""
    names = zf.namelist()
    slide_files = sorted(
        [n for n in names if re.match(r"ppt/slides/slide\d+\.xml$", n)],
        key=lambda x: int(re.search(r"slide(\d+)", x).group(1)),
    )

    for sf in slide_files:
        slide_num = int(re.search(r"slide(\d+)", sf).group(1))
        data = zf.read(sf)
        root = etree.fromstring(data)

        title_shape = find_placeholder(root, "title")
        body_shape = find_placeholder(root, "body")
        if body_shape is None:
            body_shape = find_placeholder(root, "body", idx="1")

        pic_shapes = root.xpath('.//p:sp[.//p:ph[@type="pic"]]', namespaces=NSMAP)

        if title_shape is None:
            report.issues.append(DiagnosticIssue(
                code="TMPL-010",
                severity=Severity.warning,
                message="Slide missing TITLE placeholder",
                slide_index=slide_num,
                category="placeholder",
            ))

        if body_shape is None:
            report.issues.append(DiagnosticIssue(
                code="TMPL-011",
                severity=Severity.warning,
                message="Slide missing BODY placeholder",
                slide_index=slide_num,
                category="placeholder",
            ))

        if not pic_shapes:
            report.issues.append(DiagnosticIssue(
                code="TMPL-012",
                severity=Severity.info,
                message="Slide has no PICTURE placeholder",
                slide_index=slide_num,
                category="placeholder",
            ))


def _check_stats_dashboard_shapes(
    zf: zipfile.ZipFile, config: Any, report: DiagnosticReport
) -> None:
    """TMPL-020: Check stats_dashboard slides for named stat shapes."""
    if config is None:
        return

    if not hasattr(config, "get_all_slide_indices"):
        return

    catalog = config.get_all_slide_indices()
    stat_indices = catalog.get("stats_dashboard", [])

    for idx in stat_indices:
        sf = f"ppt/slides/slide{idx}.xml"
        try:
            data = zf.read(sf)
        except KeyError:
            continue

        root = etree.fromstring(data)
        expected = ["Stat1_Number", "Stat1_Label", "Stat2_Number", "Stat2_Label"]

        for name in expected:
            shape = find_shape_by_name(root, name)
            if shape is None:
                report.issues.append(DiagnosticIssue(
                    code="TMPL-020",
                    severity=Severity.warning,
                    message=f"stats_dashboard slide missing named shape '{name}'",
                    slide_index=idx,
                    category="named_shape",
                ))
                break  # Report once per slide


def _check_case_study_shapes(
    zf: zipfile.ZipFile, config: Any, report: DiagnosticReport
) -> None:
    """TMPL-021: Check case_study_full slides for Quote/Attribution shapes."""
    if config is None:
        return

    if not hasattr(config, "get_all_slide_indices"):
        return

    catalog = config.get_all_slide_indices()
    cs_indices = catalog.get("case_study_full", [])

    for idx in cs_indices:
        sf = f"ppt/slides/slide{idx}.xml"
        try:
            data = zf.read(sf)
        except KeyError:
            continue

        root = etree.fromstring(data)
        for name in ["Quote", "Attribution"]:
            shape = find_shape_by_name(root, name)
            if shape is None:
                report.issues.append(DiagnosticIssue(
                    code="TMPL-021",
                    severity=Severity.warning,
                    message=f"case_study_full slide missing '{name}' shape",
                    slide_index=idx,
                    category="named_shape",
                ))


def _check_master_count(zf: zipfile.ZipFile, report: DiagnosticReport) -> int:
    """TMPL-030: Warn if template has >3 slide masters."""
    names = zf.namelist()
    masters = [n for n in names if re.match(r"ppt/slideMasters/slideMaster\d+\.xml$", n)]
    count = len(masters)

    if count > 3:
        report.issues.append(DiagnosticIssue(
            code="TMPL-030",
            severity=Severity.warning,
            message=f"Template has {count} slide masters (>3 may cause bloat)",
            category="structure",
        ))

    return count


def _check_missing_media(zf: zipfile.ZipFile, report: DiagnosticReport) -> None:
    """TMPL-040: Check for relationship references to missing media files."""
    names_set = set(zf.namelist())
    rels_files = [n for n in names_set if n.startswith("ppt/slides/_rels/") and n.endswith(".rels")]

    for rf in rels_files:
        slide_match = re.search(r"slide(\d+)", rf)
        slide_num = int(slide_match.group(1)) if slide_match else None

        data = zf.read(rf)
        rels_root = etree.fromstring(data)

        for rel in rels_root:
            target = rel.get("Target", "")
            rel_type = rel.get("Type", "")

            if "image" in rel_type.lower() and target.startswith("../media/"):
                media_path = "ppt/" + target.lstrip("../").replace("../", "")
                # Normalize path
                media_path = "ppt/media/" + os.path.basename(target)
                if media_path not in names_set:
                    report.issues.append(DiagnosticIssue(
                        code="TMPL-040",
                        severity=Severity.warning,
                        message=f"Relationship references missing media file: {os.path.basename(target)}",
                        slide_index=slide_num,
                        category="media",
                    ))


def _check_large_shapes(zf: zipfile.ZipFile, report: DiagnosticReport) -> None:
    """TMPL-050: Check for shapes that cover >80% of slide area."""
    # Standard slide dimensions: 10" x 7.5" = 9144000 x 6858000 EMUs
    slide_area = 9144000 * 6858000

    names = zf.namelist()
    slide_files = [n for n in names if re.match(r"ppt/slides/slide\d+\.xml$", n)]

    for sf in slide_files:
        slide_num = int(re.search(r"slide(\d+)", sf).group(1))
        data = zf.read(sf)
        root = etree.fromstring(data)

        for sp in root.xpath(".//p:sp", namespaces=NSMAP):
            xfrm = sp.find(".//a:xfrm", namespaces=NSMAP)
            if xfrm is None:
                continue
            ext = xfrm.find("a:ext", namespaces=NSMAP)
            if ext is None:
                continue

            cx = int(ext.get("cx", 0))
            cy = int(ext.get("cy", 0))
            shape_area = cx * cy

            if shape_area > 0.8 * slide_area:
                # Get shape name
                cNvPr = sp.find(".//p:cNvPr", namespaces=NSMAP)
                name = cNvPr.get("name", "unnamed") if cNvPr is not None else "unnamed"

                report.issues.append(DiagnosticIssue(
                    code="TMPL-050",
                    severity=Severity.info,
                    message=f"Shape '{name}' covers >80% of slide area (potential overlap)",
                    slide_index=slide_num,
                    category="layout",
                ))


# ============================================================
# PUBLIC API
# ============================================================

def diagnose_template(
    template_path: Union[str, Path],
    config: Any = None,
) -> DiagnosticReport:
    """Run all diagnostic checks on a template and return a report.

    Args:
        template_path: Path to the template PPTX file
        config: Optional BrandConfig for config-aware checks

    Returns:
        DiagnosticReport with all findings
    """
    template_path = Path(template_path)
    report = DiagnosticReport()

    # File-level checks
    if not _check_file_exists(template_path, report):
        return report

    if not _check_valid_zip(template_path, report):
        return report

    with zipfile.ZipFile(template_path, "r") as zf:
        if not _check_required_parts(zf, report):
            return report

        # Count slides and layouts
        names = zf.namelist()
        slide_files = [n for n in names if re.match(r"ppt/slides/slide\d+\.xml$", n)]
        layout_files = [n for n in names if re.match(r"ppt/slideLayouts/slideLayout\d+\.xml$", n)]

        report.slide_count = len(slide_files)
        report.layout_count = len(layout_files)

        # Config checks
        _check_config_slide_indices(zf, config, report)

        # Placeholder checks
        _check_slide_placeholders(zf, report)

        # Named shape checks
        _check_stats_dashboard_shapes(zf, config, report)
        _check_case_study_shapes(zf, config, report)

        # Structure checks
        report.master_count = _check_master_count(zf, report)

        # Media checks
        _check_missing_media(zf, report)

        # Layout checks
        _check_large_shapes(zf, report)

    return report
