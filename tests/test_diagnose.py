"""Tests for template diagnostic checks."""

import tempfile
import zipfile
from pathlib import Path

import pytest

from presentation_toolkit.diagnose import (
    DiagnosticIssue,
    DiagnosticReport,
    Severity,
    diagnose_template,
)


# ============================================================
# HELPER: Create minimal PPTX ZIP files for testing
# ============================================================

MINIMAL_CONTENT_TYPES = """<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
  <Default Extension="xml" ContentType="application/xml"/>
  <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
  <Override PartName="/ppt/presentation.xml"
    ContentType="application/vnd.openxmlformats-officedocument.presentationml.presentation.main+xml"/>
  <Override PartName="/ppt/slides/slide1.xml"
    ContentType="application/vnd.openxmlformats-officedocument.presentationml.slide+xml"/>
</Types>"""

MINIMAL_PRESENTATION = """<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<p:presentation xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"
  xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main"
  xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
  <p:sldIdLst>
    <p:sldId id="256" r:id="rId2"/>
  </p:sldIdLst>
</p:presentation>"""

SLIDE_WITH_PLACEHOLDERS = """<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<p:sld xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"
  xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main"
  xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
  <p:cSld>
    <p:spTree>
      <p:nvGrpSpPr><p:cNvPr id="1" name=""/><p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr>
      <p:grpSpPr/>
      <p:sp>
        <p:nvSpPr>
          <p:cNvPr id="2" name="Title"/>
          <p:cNvSpPr><a:spLocks noGrp="1"/></p:cNvSpPr>
          <p:nvPr><p:ph type="title"/></p:nvPr>
        </p:nvSpPr>
        <p:spPr>
          <a:xfrm><a:off x="0" y="0"/><a:ext cx="9144000" cy="1000000"/></a:xfrm>
        </p:spPr>
        <p:txBody>
          <a:bodyPr/><a:lstStyle/>
          <a:p><a:r><a:t>Title</a:t></a:r></a:p>
        </p:txBody>
      </p:sp>
      <p:sp>
        <p:nvSpPr>
          <p:cNvPr id="3" name="Body"/>
          <p:cNvSpPr><a:spLocks noGrp="1"/></p:cNvSpPr>
          <p:nvPr><p:ph type="body" idx="1"/></p:nvPr>
        </p:nvSpPr>
        <p:spPr>
          <a:xfrm><a:off x="0" y="1000000"/><a:ext cx="9144000" cy="5000000"/></a:xfrm>
        </p:spPr>
        <p:txBody>
          <a:bodyPr/><a:lstStyle/>
          <a:p><a:r><a:t>Body content</a:t></a:r></a:p>
        </p:txBody>
      </p:sp>
    </p:spTree>
  </p:cSld>
</p:sld>"""

SLIDE_WITHOUT_PLACEHOLDERS = """<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<p:sld xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"
  xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main"
  xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
  <p:cSld>
    <p:spTree>
      <p:nvGrpSpPr><p:cNvPr id="1" name=""/><p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr>
      <p:grpSpPr/>
    </p:spTree>
  </p:cSld>
</p:sld>"""


def _create_test_pptx(slide_xml=SLIDE_WITH_PLACEHOLDERS, extra_files=None):
    """Create a minimal PPTX file for testing."""
    tmp = tempfile.NamedTemporaryFile(suffix=".pptx", delete=False)
    with zipfile.ZipFile(tmp.name, "w") as zf:
        zf.writestr("[Content_Types].xml", MINIMAL_CONTENT_TYPES)
        zf.writestr("ppt/presentation.xml", MINIMAL_PRESENTATION)
        zf.writestr("ppt/slides/slide1.xml", slide_xml)
        if extra_files:
            for name, content in extra_files.items():
                zf.writestr(name, content)
    return Path(tmp.name)


# ============================================================
# REPORT TESTS
# ============================================================

def test_report_empty():
    report = DiagnosticReport()
    assert report.errors == []
    assert report.warnings == []
    assert not report.has_blocking_issues


def test_report_with_error():
    report = DiagnosticReport(issues=[
        DiagnosticIssue(code="TMPL-001", severity=Severity.error, message="Not found")
    ])
    assert len(report.errors) == 1
    assert report.has_blocking_issues


def test_report_with_warning_only():
    report = DiagnosticReport(issues=[
        DiagnosticIssue(code="TMPL-010", severity=Severity.warning, message="Missing title")
    ])
    assert len(report.warnings) == 1
    assert not report.has_blocking_issues


def test_report_to_dict():
    report = DiagnosticReport(
        issues=[
            DiagnosticIssue(code="TMPL-010", severity=Severity.warning, message="test", slide_index=1)
        ],
        slide_count=5,
        layout_count=3,
        master_count=1,
    )
    d = report.to_dict()
    assert d["slide_count"] == 5
    assert d["warning_count"] == 1
    assert d["error_count"] == 0
    assert len(d["issues"]) == 1
    assert d["issues"][0]["code"] == "TMPL-010"


# ============================================================
# FILE-LEVEL CHECKS
# ============================================================

def test_tmpl_001_file_not_found():
    report = diagnose_template("/nonexistent/template.pptx")
    assert report.has_blocking_issues
    codes = [i.code for i in report.issues]
    assert "TMPL-001" in codes


def test_tmpl_002_not_a_zip():
    tmp = tempfile.NamedTemporaryFile(suffix=".pptx", delete=False)
    tmp.write(b"this is not a zip file")
    tmp.close()
    try:
        report = diagnose_template(tmp.name)
        assert report.has_blocking_issues
        codes = [i.code for i in report.issues]
        assert "TMPL-002" in codes
    finally:
        Path(tmp.name).unlink()


def test_tmpl_003_missing_parts():
    tmp = tempfile.NamedTemporaryFile(suffix=".pptx", delete=False)
    with zipfile.ZipFile(tmp.name, "w") as zf:
        zf.writestr("[Content_Types].xml", MINIMAL_CONTENT_TYPES)
        # Missing ppt/presentation.xml
    try:
        report = diagnose_template(tmp.name)
        assert report.has_blocking_issues
        codes = [i.code for i in report.issues]
        assert "TMPL-003" in codes
    finally:
        Path(tmp.name).unlink()


# ============================================================
# PLACEHOLDER CHECKS
# ============================================================

def test_slide_with_placeholders_no_warnings():
    pptx = _create_test_pptx(SLIDE_WITH_PLACEHOLDERS)
    try:
        report = diagnose_template(str(pptx))
        # Should not have TMPL-010 or TMPL-011 for this slide
        placeholder_warnings = [i for i in report.issues if i.code in ("TMPL-010", "TMPL-011")]
        assert len(placeholder_warnings) == 0
    finally:
        pptx.unlink()


def test_slide_without_placeholders_warns():
    pptx = _create_test_pptx(SLIDE_WITHOUT_PLACEHOLDERS)
    try:
        report = diagnose_template(str(pptx))
        codes = [i.code for i in report.issues]
        assert "TMPL-010" in codes  # Missing title
        assert "TMPL-011" in codes  # Missing body
    finally:
        pptx.unlink()


# ============================================================
# VALID TEMPLATE
# ============================================================

def test_valid_template_counts():
    pptx = _create_test_pptx(SLIDE_WITH_PLACEHOLDERS)
    try:
        report = diagnose_template(str(pptx))
        assert report.slide_count == 1
        assert not report.has_blocking_issues
    finally:
        pptx.unlink()


def test_print_report_runs(capsys):
    report = DiagnosticReport(
        issues=[
            DiagnosticIssue(code="TMPL-010", severity=Severity.warning, message="Test warning", slide_index=1),
            DiagnosticIssue(code="TMPL-050", severity=Severity.info, message="Test info"),
        ],
        slide_count=5,
    )
    report.print_report()
    captured = capsys.readouterr()
    assert "TEMPLATE DIAGNOSTIC REPORT" in captured.out
    assert "TMPL-010" in captured.out
    assert "TMPL-050" in captured.out
