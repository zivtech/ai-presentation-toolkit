"""Tests for content document models, conversion, and serialization."""

import json
import tempfile
from pathlib import Path

import pytest

from presentation_toolkit.content import (
    CaseStudyZones,
    ContentDocument,
    ContentMetadata,
    ContentType,
    ImagePlacement,
    SlideContent,
    SlideImage,
    StatZone,
    StatsDashboardZones,
    content_document_to_slides,
    load_content_document,
    save_content_document,
    slides_to_content_document,
    validate_content_json,
)


# ============================================================
# MODEL VALIDATION TESTS
# ============================================================

def test_content_type_enum_values():
    """Verify all expected content type values exist."""
    expected = {
        "auto", "statistic", "stats_dashboard", "quote", "numbered_step",
        "bullet_list", "comparison", "section_header", "case_study",
        "case_study_full", "statement", "feature", "detailed_content",
        "title_opening", "closing",
    }
    actual = {ct.value for ct in ContentType}
    assert actual == expected


def test_image_placement_enum():
    assert ImagePlacement.auto == "auto"
    assert ImagePlacement.background == "background"


def test_slide_image_defaults():
    img = SlideImage(path="/tmp/img.png")
    assert img.width == 0
    assert img.height == 0
    assert img.placement == ImagePlacement.auto


def test_slide_content_minimal():
    slide = SlideContent(number=1, title="Hello", body="World")
    assert slide.content_type == ContentType.auto
    assert slide.zones is None
    assert slide.images == []
    assert slide.extraction_notes == []


def test_slide_content_with_stats_zones():
    zones = StatsDashboardZones(
        stats=[StatZone(number="85%", label="Satisfaction")]
    )
    slide = SlideContent(
        number=1, title="Stats", body="",
        content_type=ContentType.stats_dashboard,
        zones=zones,
    )
    assert isinstance(slide.zones, StatsDashboardZones)
    assert slide.zones.stats[0].number == "85%"


def test_slide_content_with_case_study_zones():
    zones = CaseStudyZones(
        company_name="Acme Corp",
        description="A success story",
        bullets="- Point 1\n- Point 2",
        quote="Great product",
        attribution="CEO",
    )
    slide = SlideContent(
        number=1, title="Acme", body="",
        content_type=ContentType.case_study_full,
        zones=zones,
    )
    assert isinstance(slide.zones, CaseStudyZones)
    assert slide.zones.company_name == "Acme Corp"


def test_zones_type_mismatch_raises():
    """Stats zones with case_study_full content_type should fail."""
    zones = StatsDashboardZones(
        stats=[StatZone(number="42%", label="Users")]
    )
    with pytest.raises(ValueError, match="StatsDashboardZones requires content_type='stats_dashboard'"):
        SlideContent(
            number=1, title="X", body="",
            content_type=ContentType.case_study_full,
            zones=zones,
        )


def test_zones_auto_content_type_allows_any():
    """With content_type='auto', any zones type is accepted."""
    zones = StatsDashboardZones(
        stats=[StatZone(number="10x", label="Growth")]
    )
    slide = SlideContent(number=1, title="X", body="", zones=zones)
    assert slide.content_type == ContentType.auto
    assert slide.zones is not None


def test_content_document_requires_at_least_one_slide():
    meta = ContentMetadata(source_file="test.pptx", source_format="pptx")
    with pytest.raises(Exception):
        ContentDocument(version="1.0", metadata=meta, slides=[])


def test_content_metadata_format_normalization():
    meta = ContentMetadata(source_file="test.pptx", source_format=".PPTX")
    assert meta.source_format == "pptx"


# ============================================================
# CONVERSION TESTS
# ============================================================

def _make_legacy_slides():
    return [
        {
            "number": 1,
            "title": "Title Slide",
            "body": "Subtitle here",
            "layout": "TITLE",
            "images": [{"path": "/tmp/img.png", "width": 800, "height": 600, "ext": "png"}],
            "_extraction_notes": ["Note 1"],
        },
        {
            "number": 2,
            "title": "Stats Page",
            "body": "85% satisfaction",
            "layout": "DEFAULT",
            "images": [],
        },
    ]


def test_slides_to_content_document():
    slides = _make_legacy_slides()
    doc = slides_to_content_document(slides, "input.pptx", "pptx")

    assert doc.version == "1.0"
    assert doc.metadata.source_file == "input.pptx"
    assert doc.metadata.source_format == "pptx"
    assert len(doc.slides) == 2
    assert doc.slides[0].title == "Title Slide"
    assert len(doc.slides[0].images) == 1
    assert doc.slides[0].extraction_notes == ["Note 1"]


def test_content_document_to_slides():
    slides = _make_legacy_slides()
    doc = slides_to_content_document(slides, "test.pptx", "pptx")
    result = content_document_to_slides(doc)

    assert len(result) == 2
    assert result[0]["title"] == "Title Slide"
    assert result[0]["number"] == 1
    assert len(result[0]["images"]) == 1
    assert result[0]["images"][0]["path"] == "/tmp/img.png"


def test_roundtrip_preserves_content_type():
    """If _content_type is set on legacy slide, it roundtrips through ContentDocument."""
    slides = [{"number": 1, "title": "Stat", "body": "85%", "_content_type": "statistic"}]
    doc = slides_to_content_document(slides, "test.pptx", "pptx")
    assert doc.slides[0].content_type == ContentType.statistic

    result = content_document_to_slides(doc)
    assert result[0]["_content_type"] == "statistic"


def test_roundtrip_preserves_stats_zones():
    """Zones added to ContentDocument roundtrip back to slide dicts."""
    zones = StatsDashboardZones(
        stats=[
            StatZone(number="85%", label="Satisfaction"),
            StatZone(number="$2.5M", label="Revenue"),
        ]
    )
    doc = ContentDocument(
        version="1.0",
        metadata=ContentMetadata(source_file="test.pptx", source_format="pptx"),
        slides=[SlideContent(
            number=1, title="Stats", body="",
            content_type=ContentType.stats_dashboard,
            zones=zones,
        )],
    )
    result = content_document_to_slides(doc)
    assert "_zones" in result[0]
    assert result[0]["_zones"]["type"] == "stats_dashboard"
    assert len(result[0]["_zones"]["stats"]) == 2
    assert result[0]["_zones"]["stats"][0]["number"] == "85%"


def test_roundtrip_preserves_case_study_zones():
    zones = CaseStudyZones(
        company_name="Acme",
        description="Story",
        bullets="- A\n- B",
        quote="Wow",
        attribution="CEO",
    )
    doc = ContentDocument(
        version="1.0",
        metadata=ContentMetadata(source_file="test.pptx", source_format="pptx"),
        slides=[SlideContent(
            number=1, title="Acme", body="",
            content_type=ContentType.case_study_full,
            zones=zones,
        )],
    )
    result = content_document_to_slides(doc)
    assert result[0]["_zones"]["type"] == "case_study_full"
    assert result[0]["_zones"]["company_name"] == "Acme"


# ============================================================
# SERIALIZATION TESTS
# ============================================================

def test_save_and_load_roundtrip():
    slides = _make_legacy_slides()
    doc = slides_to_content_document(slides, "test.pptx", "pptx")

    with tempfile.NamedTemporaryFile(suffix=".json", delete=False, mode="w") as f:
        temp_path = f.name

    try:
        save_content_document(doc, temp_path)

        loaded = load_content_document(temp_path)
        assert loaded.version == doc.version
        assert loaded.metadata.source_file == doc.metadata.source_file
        assert len(loaded.slides) == len(doc.slides)
        assert loaded.slides[0].title == doc.slides[0].title
    finally:
        Path(temp_path).unlink()


def test_validate_content_json_valid():
    slides = _make_legacy_slides()
    doc = slides_to_content_document(slides, "test.pptx", "pptx")

    with tempfile.NamedTemporaryFile(suffix=".json", delete=False, mode="w") as f:
        temp_path = f.name

    try:
        save_content_document(doc, temp_path)
        errors = validate_content_json(temp_path)
        assert errors == []
    finally:
        Path(temp_path).unlink()


def test_validate_content_json_invalid():
    with tempfile.NamedTemporaryFile(suffix=".json", delete=False, mode="w") as f:
        json.dump({"bad": "data"}, f)
        temp_path = f.name

    try:
        errors = validate_content_json(temp_path)
        assert len(errors) > 0
    finally:
        Path(temp_path).unlink()


def test_validate_content_json_not_json():
    with tempfile.NamedTemporaryFile(suffix=".json", delete=False, mode="w") as f:
        f.write("not json at all {{{")
        temp_path = f.name

    try:
        errors = validate_content_json(temp_path)
        assert len(errors) == 1
        assert "Invalid JSON" in errors[0]
    finally:
        Path(temp_path).unlink()


def test_validate_content_json_missing_file():
    errors = validate_content_json("/nonexistent/path.json")
    assert len(errors) == 1
    assert "not found" in errors[0].lower() or "File not found" in errors[0]
