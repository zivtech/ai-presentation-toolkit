"""
Content Document Models

Pydantic v2 models defining the validated intermediate representation
between parsing and generation. Mirrors the JSON Schema in schemas/content.schema.json.
"""

import json
from datetime import datetime, timezone
from enum import Enum
from pathlib import Path
from typing import Any, Dict, List, Optional, Union

from pydantic import BaseModel, Field, field_validator, model_validator


class ContentType(str, Enum):
    """Content type classification for a slide."""
    auto = "auto"
    statistic = "statistic"
    stats_dashboard = "stats_dashboard"
    quote = "quote"
    numbered_step = "numbered_step"
    bullet_list = "bullet_list"
    comparison = "comparison"
    section_header = "section_header"
    case_study = "case_study"
    case_study_full = "case_study_full"
    statement = "statement"
    feature = "feature"
    detailed_content = "detailed_content"
    title_opening = "title_opening"
    closing = "closing"


class ImagePlacement(str, Enum):
    """Placement hint for an image on a slide."""
    auto = "auto"
    background = "background"
    left = "left"
    right = "right"
    center = "center"
    inline = "inline"


class SlideImage(BaseModel):
    """An image associated with a slide."""
    path: str
    width: int = 0
    height: int = 0
    ext: str = ""
    alt_text: str = ""
    placement: ImagePlacement = ImagePlacement.auto


class StatZone(BaseModel):
    """A single statistic in a stats dashboard."""
    number: str
    label: str


class StatsDashboardZones(BaseModel):
    """Zone data for a stats_dashboard layout."""
    type: str = Field("stats_dashboard", frozen=True)
    stats: List[StatZone] = Field(min_length=1, max_length=6)


class CaseStudyZones(BaseModel):
    """Zone data for a case_study_full layout."""
    type: str = Field("case_study_full", frozen=True)
    company_name: str = ""
    description: str = ""
    bullets: str = ""
    quote: str = ""
    attribution: str = ""


class SlideContent(BaseModel):
    """A single slide in the content document."""
    number: int = Field(ge=1)
    title: str
    body: str
    speaker_notes: str = ""
    content_type: ContentType = ContentType.auto
    layout_hint: str = ""
    images: List[SlideImage] = Field(default_factory=list)
    zones: Optional[Union[StatsDashboardZones, CaseStudyZones]] = None
    extraction_notes: List[str] = Field(default_factory=list)

    @model_validator(mode="after")
    def zones_must_match_content_type(self) -> "SlideContent":
        """Validate that zones type matches content_type when both are set."""
        if self.zones is not None and self.content_type != ContentType.auto:
            if isinstance(self.zones, StatsDashboardZones) and self.content_type != ContentType.stats_dashboard:
                raise ValueError(
                    f"StatsDashboardZones requires content_type='stats_dashboard', got '{self.content_type.value}'"
                )
            if isinstance(self.zones, CaseStudyZones) and self.content_type != ContentType.case_study_full:
                raise ValueError(
                    f"CaseStudyZones requires content_type='case_study_full', got '{self.content_type.value}'"
                )
        return self


class ContentMetadata(BaseModel):
    """Metadata about the content document."""
    source_file: str
    source_format: str
    generated_at: str = ""
    generator: str = ""

    @field_validator("source_format", mode="before")
    @classmethod
    def normalize_format(cls, v: str) -> str:
        return v.lower().lstrip(".")


class ContentDocument(BaseModel):
    """Top-level content document representing an entire presentation."""
    version: str = "1.0"
    metadata: ContentMetadata
    slides: List[SlideContent] = Field(min_length=1)


# ============================================================
# CONVERSION FUNCTIONS
# ============================================================

def slides_to_content_document(
    slides: List[Dict[str, Any]],
    source_file: str = "",
    source_format: str = "pptx",
) -> ContentDocument:
    """Convert the old List[Dict] slide format to a ContentDocument.

    This bridges the legacy pipeline output to the new validated format.
    """
    content_slides = []
    for slide in slides:
        images = []
        for img in slide.get("images", []):
            if isinstance(img, str):
                images.append(SlideImage(path=img))
            elif isinstance(img, dict):
                images.append(SlideImage(
                    path=img.get("path", ""),
                    width=img.get("width", 0),
                    height=img.get("height", 0),
                    ext=img.get("ext", ""),
                    alt_text=img.get("alt_text", ""),
                    placement=img.get("placement", "auto"),
                ))

        extraction_notes = slide.get("_extraction_notes", [])

        content_slides.append(SlideContent(
            number=slide.get("number", len(content_slides) + 1),
            title=slide.get("title", ""),
            body=slide.get("body", ""),
            speaker_notes=slide.get("speaker_notes", ""),
            content_type=slide.get("_content_type", "auto"),
            layout_hint=slide.get("layout_hint", slide.get("layout", "")),
            images=images,
            extraction_notes=extraction_notes,
        ))

    metadata = ContentMetadata(
        source_file=source_file,
        source_format=source_format,
        generated_at=datetime.now(timezone.utc).isoformat(),
        generator="ai-presentation-toolkit",
    )

    return ContentDocument(version="1.0", metadata=metadata, slides=content_slides)


def content_document_to_slides(doc: ContentDocument) -> List[Dict[str, Any]]:
    """Convert a ContentDocument back to the old List[Dict] format.

    Preserves content_type and zones as special keys so the migration
    engine can skip re-detection and re-parsing.
    """
    slides = []
    for sc in doc.slides:
        slide: Dict[str, Any] = {
            "number": sc.number,
            "title": sc.title,
            "body": sc.body,
            "layout": sc.layout_hint or "DEFAULT",
            "images": [
                {
                    "path": img.path,
                    "width": img.width,
                    "height": img.height,
                    "ext": img.ext,
                }
                for img in sc.images
            ],
            "image_count": len(sc.images),
        }

        if sc.content_type != ContentType.auto:
            slide["_content_type"] = sc.content_type.value

        if sc.zones is not None:
            if isinstance(sc.zones, StatsDashboardZones):
                slide["_zones"] = {
                    "type": "stats_dashboard",
                    "stats": [{"number": s.number, "label": s.label} for s in sc.zones.stats],
                }
            elif isinstance(sc.zones, CaseStudyZones):
                slide["_zones"] = {
                    "type": "case_study_full",
                    "company_name": sc.zones.company_name,
                    "description": sc.zones.description,
                    "bullets": sc.zones.bullets,
                    "quote": sc.zones.quote,
                    "attribution": sc.zones.attribution,
                }

        if sc.extraction_notes:
            slide["_extraction_notes"] = sc.extraction_notes

        if sc.speaker_notes:
            slide["speaker_notes"] = sc.speaker_notes

        slides.append(slide)

    return slides


# ============================================================
# SERIALIZATION
# ============================================================

def save_content_document(doc: ContentDocument, path: Union[str, Path]) -> None:
    """Save a ContentDocument to a JSON file."""
    path = Path(path)
    data = doc.model_dump(mode="json", exclude_none=True)
    with open(path, "w", encoding="utf-8") as f:
        json.dump(data, f, indent=2, ensure_ascii=False)


def load_content_document(path: Union[str, Path]) -> ContentDocument:
    """Load a ContentDocument from a JSON file."""
    path = Path(path)
    with open(path, "r", encoding="utf-8") as f:
        data = json.load(f)
    return ContentDocument(**data)


# ============================================================
# VALIDATION
# ============================================================

def validate_content_json(path: Union[str, Path]) -> List[str]:
    """Validate a content JSON file and return a list of error strings.

    Uses jsonschema if available, otherwise falls back to Pydantic validation.
    Returns an empty list when the document is valid.
    """
    path = Path(path)
    errors: List[str] = []

    try:
        with open(path, "r", encoding="utf-8") as f:
            data = json.load(f)
    except json.JSONDecodeError as exc:
        return [f"Invalid JSON: {exc}"]
    except FileNotFoundError:
        return [f"File not found: {path}"]

    # Try jsonschema first
    try:
        import jsonschema
        from .schemas import get_content_schema_path

        with open(get_content_schema_path(), "r") as f:
            schema = json.load(f)

        validator = jsonschema.Draft202012Validator(schema)
        for error in validator.iter_errors(data):
            json_path = " -> ".join(str(p) for p in error.absolute_path) if error.absolute_path else "(root)"
            errors.append(f"{json_path}: {error.message}")
        return errors
    except ImportError:
        pass

    # Fallback: Pydantic validation
    try:
        ContentDocument(**data)
    except Exception as exc:
        errors.append(str(exc))

    return errors
