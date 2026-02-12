"""Tests for the layout cookbook recipes and XML generation."""

import pytest
from lxml import etree

from presentation_toolkit.cookbook import (
    COOKBOOK,
    BoxPosition,
    LayoutRecipe,
    TextBoxSpec,
    apply_recipe_to_slide,
    build_slide_from_recipe,
    build_text_box_xml,
    get_recipe,
    list_recipes,
)
from presentation_toolkit.pptx_utils import EMU_PER_INCH, NSMAP


# ============================================================
# BoxPosition TESTS
# ============================================================

def test_box_position_from_inches():
    pos = BoxPosition.from_inches(1.0, 2.0, 3.0, 4.0)
    assert pos.x == int(1.0 * EMU_PER_INCH)
    assert pos.y == int(2.0 * EMU_PER_INCH)
    assert pos.cx == int(3.0 * EMU_PER_INCH)
    assert pos.cy == int(4.0 * EMU_PER_INCH)


def test_box_position_zero():
    pos = BoxPosition.from_inches(0, 0, 0, 0)
    assert pos.x == 0
    assert pos.cy == 0


# ============================================================
# RECIPE REGISTRY TESTS
# ============================================================

def test_cookbook_has_expected_recipes():
    expected = {
        "title_opening", "feature_default", "content_image_left",
        "content_image_right", "statement_center", "stat_default",
        "stats_dashboard", "quote_default", "two_column",
        "section_divider", "case_study_full", "closing_cta", "hero_photo",
    }
    actual = set(COOKBOOK.keys())
    assert expected == actual


def test_list_recipes_returns_all():
    names = list_recipes()
    assert len(names) == 13
    assert "feature_default" in names
    assert "stats_dashboard" in names


def test_get_recipe_found():
    recipe = get_recipe("feature_default")
    assert recipe is not None
    assert recipe.name == "feature_default"
    assert recipe.title is not None
    assert recipe.body is not None


def test_get_recipe_not_found():
    result = get_recipe("nonexistent_layout")
    assert result is None


def test_all_recipes_have_title():
    """Every recipe should have at least a title spec."""
    for name, recipe in COOKBOOK.items():
        assert recipe.title is not None or recipe.extra_text_boxes, \
            f"Recipe '{name}' has no title or extra boxes"


def test_stats_dashboard_has_12_extra_boxes():
    """Stats dashboard should have 6 number + 6 label = 12 extra text boxes."""
    recipe = get_recipe("stats_dashboard")
    assert recipe is not None
    assert len(recipe.extra_text_boxes) == 12
    # Check naming pattern
    names = [tb.name for tb in recipe.extra_text_boxes]
    assert "Stat1_Number" in names
    assert "Stat6_Label" in names


def test_case_study_full_has_extra_zones():
    recipe = get_recipe("case_study_full")
    assert recipe is not None
    assert len(recipe.extra_text_boxes) == 3
    names = [tb.name for tb in recipe.extra_text_boxes]
    assert "bullets" in names
    assert "Quote" in names
    assert "Attribution" in names


def test_hero_photo_has_image_box():
    recipe = get_recipe("hero_photo")
    assert recipe is not None
    assert len(recipe.image_boxes) == 1
    assert recipe.image_boxes[0].name == "background"
    assert recipe.background.image is True


# ============================================================
# XML GENERATION TESTS
# ============================================================

def test_build_text_box_xml_structure():
    spec = TextBoxSpec(
        name="test_box",
        position=BoxPosition.from_inches(1, 1, 5, 2),
        font_size_pt=24,
        bold=True,
        alignment="ctr",
    )
    elem = build_text_box_xml(spec, shape_id=42)

    # Check it's a p:sp element
    assert elem.tag.endswith("}sp")

    # Check shape ID and name
    cNvPr = elem.find(".//p:cNvPr", namespaces=NSMAP)
    assert cNvPr is not None
    assert cNvPr.get("id") == "42"
    assert cNvPr.get("name") == "test_box"

    # Check it's a text box
    cNvSpPr = elem.find(".//p:cNvSpPr", namespaces=NSMAP)
    assert cNvSpPr is not None
    assert cNvSpPr.get("txBox") == "1"

    # Check position
    off = elem.find(".//a:off", namespaces=NSMAP)
    assert off is not None
    assert int(off.get("x")) == int(1 * EMU_PER_INCH)

    ext = elem.find(".//a:ext", namespaces=NSMAP)
    assert ext is not None
    assert int(ext.get("cx")) == int(5 * EMU_PER_INCH)

    # Check font properties
    rPr = elem.find(".//a:rPr", namespaces=NSMAP)
    assert rPr is not None
    assert rPr.get("sz") == "2400"  # 24pt * 100
    assert rPr.get("b") == "1"

    # Check alignment
    pPr = elem.find(".//a:pPr", namespaces=NSMAP)
    assert pPr is not None
    assert pPr.get("algn") == "ctr"


def test_build_text_box_xml_italic():
    spec = TextBoxSpec(
        name="quote",
        position=BoxPosition.from_inches(0, 0, 5, 2),
        italic=True,
    )
    elem = build_text_box_xml(spec, shape_id=10)
    rPr = elem.find(".//a:rPr", namespaces=NSMAP)
    assert rPr.get("i") == "1"
    assert rPr.get("b") is None  # bold should not be set


def test_build_slide_from_recipe():
    recipe = get_recipe("feature_default")
    sld = build_slide_from_recipe(recipe, title="Test Title", body="Test body content")

    # Should be a valid p:sld element
    assert sld.tag.endswith("}sld")

    # Should have shapes in spTree
    spTree = sld.find(".//p:spTree", namespaces=NSMAP)
    assert spTree is not None

    # Should have text content
    texts = [t.text for t in sld.xpath(".//a:t", namespaces=NSMAP) if t.text]
    assert "Test Title" in texts
    assert "Test body content" in texts


def test_build_slide_from_recipe_empty_text():
    recipe = get_recipe("feature_default")
    sld = build_slide_from_recipe(recipe, title="", body="")

    # With empty text, no extra shapes should be added beyond group shape
    spTree = sld.find(".//p:spTree", namespaces=NSMAP)
    shapes = spTree.findall("p:sp", namespaces=NSMAP)
    assert len(shapes) == 0


def test_apply_recipe_to_existing_slide():
    """Apply a recipe to a minimal existing slide XML."""
    NSMAP_FULL = {
        "a": "http://schemas.openxmlformats.org/drawingml/2006/main",
        "p": "http://schemas.openxmlformats.org/presentationml/2006/main",
        "r": "http://schemas.openxmlformats.org/officeDocument/2006/relationships",
    }
    _P = NSMAP_FULL["p"]
    _A = NSMAP_FULL["a"]

    sld = etree.Element(f"{{{_P}}}sld", nsmap=NSMAP_FULL)
    cSld = etree.SubElement(sld, f"{{{_P}}}cSld")
    spTree = etree.SubElement(cSld, f"{{{_P}}}spTree")

    nvGrpSpPr = etree.SubElement(spTree, f"{{{_P}}}nvGrpSpPr")
    cNvPr = etree.SubElement(nvGrpSpPr, f"{{{_P}}}cNvPr")
    cNvPr.set("id", "1")
    cNvPr.set("name", "")
    etree.SubElement(nvGrpSpPr, f"{{{_P}}}cNvGrpSpPr")
    etree.SubElement(nvGrpSpPr, f"{{{_P}}}nvPr")
    etree.SubElement(spTree, f"{{{_P}}}grpSpPr")

    recipe = get_recipe("statement_center")
    apply_recipe_to_slide(sld, recipe, title="Big Statement", body="Supporting text")

    # Should now have shapes with the text
    texts = [t.text for t in sld.xpath(".//a:t", namespaces=NSMAP) if t.text]
    assert "Big Statement" in texts
    assert "Supporting text" in texts
