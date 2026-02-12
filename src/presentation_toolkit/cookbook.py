"""
Slide Layout Cookbook

Explicit positioning recipes for each layout type. Provides a fallback
when template placeholders are missing — content won't get silently lost.

All dimensions are in EMUs (English Metric Units).
1 inch = 914400 EMUs.
Standard slide: 10" x 7.5" (9144000 x 6858000 EMUs).
"""

from dataclasses import dataclass, field
from typing import Dict, List, Optional

from lxml import etree

from .pptx_utils import NSMAP, EMU_PER_INCH, EMU_PER_PT


# ============================================================
# DATA CLASSES
# ============================================================

@dataclass
class BoxPosition:
    """Position and size in EMUs."""
    x: int
    y: int
    cx: int
    cy: int

    @classmethod
    def from_inches(cls, x: float, y: float, w: float, h: float) -> "BoxPosition":
        """Create from inch measurements."""
        return cls(
            x=int(x * EMU_PER_INCH),
            y=int(y * EMU_PER_INCH),
            cx=int(w * EMU_PER_INCH),
            cy=int(h * EMU_PER_INCH),
        )


@dataclass
class TextBoxSpec:
    """Specification for a text box to be placed on a slide."""
    name: str
    position: BoxPosition
    font_size_pt: int = 14
    bold: bool = False
    italic: bool = False
    alignment: str = "l"  # l, ctr, r
    font_color: str = "000000"
    vertical_anchor: str = "t"  # t, ctr, b


@dataclass
class ImageBoxSpec:
    """Specification for an image placeholder position."""
    name: str
    position: BoxPosition


@dataclass
class BackgroundSpec:
    """Background specification for a recipe."""
    color: str = "FFFFFF"
    image: bool = False


@dataclass
class LayoutRecipe:
    """A complete layout recipe with positioning for all elements."""
    name: str
    description: str = ""
    title: Optional[TextBoxSpec] = None
    body: Optional[TextBoxSpec] = None
    extra_text_boxes: List[TextBoxSpec] = field(default_factory=list)
    image_boxes: List[ImageBoxSpec] = field(default_factory=list)
    background: BackgroundSpec = field(default_factory=BackgroundSpec)


# ============================================================
# BUILT-IN RECIPES
# ============================================================

def _build_recipes() -> Dict[str, LayoutRecipe]:
    """Build all 13 built-in layout recipes."""
    recipes: Dict[str, LayoutRecipe] = {}

    # 1. Title Opening
    recipes["title_opening"] = LayoutRecipe(
        name="title_opening",
        description="Opening/title slide with large centered title",
        title=TextBoxSpec(
            name="title",
            position=BoxPosition.from_inches(1.0, 2.0, 8.0, 2.0),
            font_size_pt=36,
            bold=True,
            alignment="ctr",
            vertical_anchor="ctr",
        ),
        body=TextBoxSpec(
            name="subtitle",
            position=BoxPosition.from_inches(2.0, 4.2, 6.0, 1.0),
            font_size_pt=18,
            alignment="ctr",
            font_color="666666",
        ),
    )

    # 2. Feature Default
    recipes["feature_default"] = LayoutRecipe(
        name="feature_default",
        description="Standard feature slide with title and bullet body",
        title=TextBoxSpec(
            name="title",
            position=BoxPosition.from_inches(0.7, 0.5, 8.6, 1.0),
            font_size_pt=28,
            bold=True,
        ),
        body=TextBoxSpec(
            name="body",
            position=BoxPosition.from_inches(0.7, 1.7, 8.6, 5.0),
            font_size_pt=14,
        ),
    )

    # 3. Content + Image Left
    recipes["content_image_left"] = LayoutRecipe(
        name="content_image_left",
        description="Image on left, text content on right",
        title=TextBoxSpec(
            name="title",
            position=BoxPosition.from_inches(5.2, 0.5, 4.5, 1.0),
            font_size_pt=24,
            bold=True,
        ),
        body=TextBoxSpec(
            name="body",
            position=BoxPosition.from_inches(5.2, 1.7, 4.5, 5.0),
            font_size_pt=14,
        ),
        image_boxes=[
            ImageBoxSpec(
                name="image_left",
                position=BoxPosition.from_inches(0.3, 0.3, 4.5, 6.9),
            ),
        ],
    )

    # 4. Content + Image Right
    recipes["content_image_right"] = LayoutRecipe(
        name="content_image_right",
        description="Text content on left, image on right",
        title=TextBoxSpec(
            name="title",
            position=BoxPosition.from_inches(0.5, 0.5, 4.5, 1.0),
            font_size_pt=24,
            bold=True,
        ),
        body=TextBoxSpec(
            name="body",
            position=BoxPosition.from_inches(0.5, 1.7, 4.5, 5.0),
            font_size_pt=14,
        ),
        image_boxes=[
            ImageBoxSpec(
                name="image_right",
                position=BoxPosition.from_inches(5.2, 0.3, 4.5, 6.9),
            ),
        ],
    )

    # 5. Statement Center
    recipes["statement_center"] = LayoutRecipe(
        name="statement_center",
        description="Bold centered statement with optional subtitle",
        title=TextBoxSpec(
            name="title",
            position=BoxPosition.from_inches(1.0, 2.0, 8.0, 2.5),
            font_size_pt=32,
            bold=True,
            alignment="ctr",
            vertical_anchor="ctr",
        ),
        body=TextBoxSpec(
            name="body",
            position=BoxPosition.from_inches(1.5, 4.8, 7.0, 1.5),
            font_size_pt=16,
            alignment="ctr",
            font_color="666666",
        ),
    )

    # 6. Stat Default
    recipes["stat_default"] = LayoutRecipe(
        name="stat_default",
        description="Single large statistic with label",
        title=TextBoxSpec(
            name="stat_number",
            position=BoxPosition.from_inches(1.0, 1.5, 8.0, 3.0),
            font_size_pt=72,
            bold=True,
            alignment="ctr",
            vertical_anchor="ctr",
        ),
        body=TextBoxSpec(
            name="stat_label",
            position=BoxPosition.from_inches(1.5, 4.5, 7.0, 2.0),
            font_size_pt=18,
            alignment="ctr",
            font_color="666666",
        ),
    )

    # 7. Stats Dashboard (6-zone 3x2 grid)
    stat_zones: List[TextBoxSpec] = []
    col_positions = [0.5, 3.5, 6.5]
    row_positions = [1.5, 4.2]

    for row_idx, row_y in enumerate(row_positions):
        for col_idx, col_x in enumerate(col_positions):
            zone_num = row_idx * 3 + col_idx + 1
            # Number box
            stat_zones.append(TextBoxSpec(
                name=f"Stat{zone_num}_Number",
                position=BoxPosition.from_inches(col_x, row_y, 2.5, 1.2),
                font_size_pt=36,
                bold=True,
                alignment="ctr",
                vertical_anchor="b",
            ))
            # Label box
            stat_zones.append(TextBoxSpec(
                name=f"Stat{zone_num}_Label",
                position=BoxPosition.from_inches(col_x, row_y + 1.2, 2.5, 0.8),
                font_size_pt=12,
                alignment="ctr",
                font_color="666666",
                vertical_anchor="t",
            ))

    recipes["stats_dashboard"] = LayoutRecipe(
        name="stats_dashboard",
        description="6-zone 3x2 grid for multiple statistics",
        title=TextBoxSpec(
            name="title",
            position=BoxPosition.from_inches(0.5, 0.3, 9.0, 0.8),
            font_size_pt=24,
            bold=True,
            alignment="ctr",
        ),
        extra_text_boxes=stat_zones,
    )

    # 8. Quote Default
    recipes["quote_default"] = LayoutRecipe(
        name="quote_default",
        description="Centered quote with attribution",
        title=TextBoxSpec(
            name="quote",
            position=BoxPosition.from_inches(1.5, 1.5, 7.0, 3.5),
            font_size_pt=24,
            italic=True,
            alignment="ctr",
            vertical_anchor="ctr",
        ),
        body=TextBoxSpec(
            name="attribution",
            position=BoxPosition.from_inches(2.0, 5.2, 6.0, 1.0),
            font_size_pt=14,
            alignment="ctr",
            font_color="888888",
        ),
    )

    # 9. Two Column
    recipes["two_column"] = LayoutRecipe(
        name="two_column",
        description="Two-column comparison layout",
        title=TextBoxSpec(
            name="title",
            position=BoxPosition.from_inches(0.5, 0.5, 9.0, 0.8),
            font_size_pt=24,
            bold=True,
            alignment="ctr",
        ),
        body=TextBoxSpec(
            name="col_left",
            position=BoxPosition.from_inches(0.5, 1.5, 4.2, 5.2),
            font_size_pt=14,
        ),
        extra_text_boxes=[
            TextBoxSpec(
                name="col_right",
                position=BoxPosition.from_inches(5.3, 1.5, 4.2, 5.2),
                font_size_pt=14,
            ),
        ],
    )

    # 10. Section Divider
    recipes["section_divider"] = LayoutRecipe(
        name="section_divider",
        description="Section divider with large centered text",
        title=TextBoxSpec(
            name="title",
            position=BoxPosition.from_inches(1.0, 2.5, 8.0, 2.5),
            font_size_pt=40,
            bold=True,
            alignment="ctr",
            vertical_anchor="ctr",
        ),
        background=BackgroundSpec(color="12285F"),
    )

    # 11. Case Study Full (5 zones)
    recipes["case_study_full"] = LayoutRecipe(
        name="case_study_full",
        description="Full case study with company, description, bullets, quote, attribution",
        title=TextBoxSpec(
            name="company_name",
            position=BoxPosition.from_inches(0.5, 0.3, 9.0, 0.8),
            font_size_pt=28,
            bold=True,
        ),
        body=TextBoxSpec(
            name="description",
            position=BoxPosition.from_inches(0.5, 1.3, 5.0, 2.5),
            font_size_pt=14,
        ),
        extra_text_boxes=[
            TextBoxSpec(
                name="bullets",
                position=BoxPosition.from_inches(0.5, 4.0, 5.0, 3.0),
                font_size_pt=13,
            ),
            TextBoxSpec(
                name="Quote",
                position=BoxPosition.from_inches(6.0, 1.3, 3.5, 3.5),
                font_size_pt=16,
                italic=True,
                alignment="ctr",
                vertical_anchor="ctr",
            ),
            TextBoxSpec(
                name="Attribution",
                position=BoxPosition.from_inches(6.0, 5.0, 3.5, 1.0),
                font_size_pt=12,
                alignment="ctr",
                font_color="888888",
            ),
        ],
    )

    # 12. Closing CTA
    recipes["closing_cta"] = LayoutRecipe(
        name="closing_cta",
        description="Closing slide with call to action",
        title=TextBoxSpec(
            name="title",
            position=BoxPosition.from_inches(1.0, 2.0, 8.0, 2.0),
            font_size_pt=32,
            bold=True,
            alignment="ctr",
            vertical_anchor="ctr",
        ),
        body=TextBoxSpec(
            name="cta",
            position=BoxPosition.from_inches(2.0, 4.5, 6.0, 1.5),
            font_size_pt=18,
            alignment="ctr",
            font_color="009CDE",
        ),
    )

    # 13. Hero Photo
    recipes["hero_photo"] = LayoutRecipe(
        name="hero_photo",
        description="Full-bleed photo background with text overlay",
        title=TextBoxSpec(
            name="title",
            position=BoxPosition.from_inches(0.7, 4.5, 8.6, 1.5),
            font_size_pt=32,
            bold=True,
            font_color="FFFFFF",
        ),
        body=TextBoxSpec(
            name="body",
            position=BoxPosition.from_inches(0.7, 6.0, 8.6, 1.0),
            font_size_pt=16,
            font_color="FFFFFF",
        ),
        image_boxes=[
            ImageBoxSpec(
                name="background",
                position=BoxPosition.from_inches(0, 0, 10, 7.5),
            ),
        ],
        background=BackgroundSpec(image=True),
    )

    return recipes


# Global recipe registry
COOKBOOK: Dict[str, LayoutRecipe] = _build_recipes()


def get_recipe(name: str) -> Optional[LayoutRecipe]:
    """Get a layout recipe by name. Returns None if not found."""
    return COOKBOOK.get(name)


def list_recipes() -> List[str]:
    """Return all registered recipe names."""
    return list(COOKBOOK.keys())


# ============================================================
# XML GENERATION
# ============================================================

_A = "http://schemas.openxmlformats.org/drawingml/2006/main"
_P = "http://schemas.openxmlformats.org/presentationml/2006/main"
_R = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"


def _make_elem(ns: str, tag: str, **attribs) -> etree._Element:
    """Create a namespaced element with attributes."""
    elem = etree.Element(f"{{{ns}}}{tag}")
    for k, v in attribs.items():
        elem.set(k, str(v))
    return elem


def build_text_box_xml(spec: TextBoxSpec, shape_id: int) -> etree._Element:
    """Create a complete p:sp element for a text box from a TextBoxSpec.

    Args:
        spec: The text box specification
        shape_id: Unique shape ID within the slide

    Returns:
        lxml Element representing a p:sp shape
    """
    sp = _make_elem(_P, "sp")

    # nvSpPr
    nvSpPr = etree.SubElement(sp, f"{{{_P}}}nvSpPr")
    cNvPr = etree.SubElement(nvSpPr, f"{{{_P}}}cNvPr")
    cNvPr.set("id", str(shape_id))
    cNvPr.set("name", spec.name)

    cNvSpPr = etree.SubElement(nvSpPr, f"{{{_P}}}cNvSpPr")
    cNvSpPr.set("txBox", "1")

    nvPr = etree.SubElement(nvSpPr, f"{{{_P}}}nvPr")

    # spPr
    spPr = etree.SubElement(sp, f"{{{_P}}}spPr")
    xfrm = etree.SubElement(spPr, f"{{{_A}}}xfrm")

    off = etree.SubElement(xfrm, f"{{{_A}}}off")
    off.set("x", str(spec.position.x))
    off.set("y", str(spec.position.y))

    ext = etree.SubElement(xfrm, f"{{{_A}}}ext")
    ext.set("cx", str(spec.position.cx))
    ext.set("cy", str(spec.position.cy))

    prstGeom = etree.SubElement(spPr, f"{{{_A}}}prstGeom")
    prstGeom.set("prst", "rect")
    etree.SubElement(prstGeom, f"{{{_A}}}avLst")

    etree.SubElement(spPr, f"{{{_A}}}noFill")

    # txBody
    txBody = etree.SubElement(sp, f"{{{_P}}}txBody")

    bodyPr = etree.SubElement(txBody, f"{{{_A}}}bodyPr")
    bodyPr.set("wrap", "square")
    bodyPr.set("anchor", spec.vertical_anchor)
    bodyPr.set("rtlCol", "0")

    etree.SubElement(txBody, f"{{{_A}}}lstStyle")

    p_elem = etree.SubElement(txBody, f"{{{_A}}}p")

    pPr = etree.SubElement(p_elem, f"{{{_A}}}pPr")
    pPr.set("algn", spec.alignment)

    r_elem = etree.SubElement(p_elem, f"{{{_A}}}r")

    rPr = etree.SubElement(r_elem, f"{{{_A}}}rPr")
    rPr.set("lang", "en-US")
    rPr.set("sz", str(spec.font_size_pt * 100))
    rPr.set("dirty", "0")
    if spec.bold:
        rPr.set("b", "1")
    if spec.italic:
        rPr.set("i", "1")

    # Font color
    solidFill = etree.SubElement(rPr, f"{{{_A}}}solidFill")
    srgbClr = etree.SubElement(solidFill, f"{{{_A}}}srgbClr")
    srgbClr.set("val", spec.font_color)

    t_elem = etree.SubElement(r_elem, f"{{{_A}}}t")
    t_elem.text = ""  # Placeholder text — caller will set actual content

    return sp


def apply_recipe_to_slide(
    slide_root: etree._Element,
    recipe: LayoutRecipe,
    title: str = "",
    body: str = "",
) -> None:
    """Add positioned text boxes from a recipe to an existing slide XML.

    This is the fallback path — used when template placeholders are missing.

    Args:
        slide_root: Root element of the slide XML (p:sld)
        recipe: The layout recipe to apply
        title: Title text to insert
        body: Body text to insert
    """
    spTree = slide_root.find(".//p:cSld/p:spTree", namespaces=NSMAP)
    if spTree is None:
        return

    # Find max existing shape ID
    max_id = 1
    for cNvPr in spTree.xpath(".//p:cNvPr", namespaces=NSMAP):
        try:
            sid = int(cNvPr.get("id", "0"))
            max_id = max(max_id, sid)
        except ValueError:
            pass

    next_id = max_id + 1

    if recipe.title and title:
        sp = build_text_box_xml(recipe.title, next_id)
        next_id += 1
        # Set actual text
        t_elem = sp.find(f".//{{{_A}}}t")
        if t_elem is not None:
            t_elem.text = title
        spTree.append(sp)

    if recipe.body and body:
        sp = build_text_box_xml(recipe.body, next_id)
        next_id += 1
        t_elem = sp.find(f".//{{{_A}}}t")
        if t_elem is not None:
            t_elem.text = body
        spTree.append(sp)

    for extra in recipe.extra_text_boxes:
        sp = build_text_box_xml(extra, next_id)
        next_id += 1
        spTree.append(sp)


def build_slide_from_recipe(
    recipe: LayoutRecipe,
    title: str = "",
    body: str = "",
) -> etree._Element:
    """Create a complete slide XML element from scratch using a recipe.

    This is the standalone mode — creates a full slide without needing a template.

    Args:
        recipe: The layout recipe
        title: Title text
        body: Body text

    Returns:
        Root p:sld element
    """
    NSMAP_FULL = {
        "a": _A,
        "p": _P,
        "r": _R,
    }

    sld = etree.Element(f"{{{_P}}}sld", nsmap=NSMAP_FULL)
    cSld = etree.SubElement(sld, f"{{{_P}}}cSld")
    spTree = etree.SubElement(cSld, f"{{{_P}}}spTree")

    # Group shape properties (required)
    nvGrpSpPr = etree.SubElement(spTree, f"{{{_P}}}nvGrpSpPr")
    cNvPr = etree.SubElement(nvGrpSpPr, f"{{{_P}}}cNvPr")
    cNvPr.set("id", "1")
    cNvPr.set("name", "")
    etree.SubElement(nvGrpSpPr, f"{{{_P}}}cNvGrpSpPr")
    etree.SubElement(nvGrpSpPr, f"{{{_P}}}nvPr")

    grpSpPr = etree.SubElement(spTree, f"{{{_P}}}grpSpPr")
    xfrm = etree.SubElement(grpSpPr, f"{{{_A}}}xfrm")
    for tag in ["off", "ext"]:
        elem = etree.SubElement(xfrm, f"{{{_A}}}{tag}")
        elem.set("x" if tag == "off" else "cx", "0")
        elem.set("y" if tag == "off" else "cy", "0")
    chOff = etree.SubElement(xfrm, f"{{{_A}}}chOff")
    chOff.set("x", "0")
    chOff.set("y", "0")
    chExt = etree.SubElement(xfrm, f"{{{_A}}}chExt")
    chExt.set("cx", "0")
    chExt.set("cy", "0")

    next_id = 2

    if recipe.title and title:
        sp = build_text_box_xml(recipe.title, next_id)
        next_id += 1
        t_elem = sp.find(f".//{{{_A}}}t")
        if t_elem is not None:
            t_elem.text = title
        spTree.append(sp)

    if recipe.body and body:
        sp = build_text_box_xml(recipe.body, next_id)
        next_id += 1
        t_elem = sp.find(f".//{{{_A}}}t")
        if t_elem is not None:
            t_elem.text = body
        spTree.append(sp)

    for extra in recipe.extra_text_boxes:
        sp = build_text_box_xml(extra, next_id)
        next_id += 1
        spTree.append(sp)

    return sld
