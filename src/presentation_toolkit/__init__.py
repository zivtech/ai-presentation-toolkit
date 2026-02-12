"""
AI Presentation Toolkit

A generic, brand-agnostic toolkit for migrating PowerPoint presentations
to branded templates with intelligent content-type detection and layout variety.
"""

__version__ = "0.1.0"

from .config import (
    BrandConfig,
    load_config,
    save_config,
    create_minimal_config,
)

from .migrate import (
    migrate_presentation,
    migrate_from_content,
    detect_and_parse,
    detect_content_type,
    LayoutSelector,
)

from .analyze import (
    analyze_presentation,
    get_analysis_json,
)

from .extract import (
    extract_pptx_to_markdown,
)

from .content import (
    ContentDocument,
    SlideContent,
    ContentType,
    slides_to_content_document,
    content_document_to_slides,
    load_content_document,
    save_content_document,
)

from .cookbook import (
    LayoutRecipe,
    COOKBOOK,
    get_recipe,
    list_recipes,
)

from .diagnose import (
    diagnose_template,
    DiagnosticReport,
)

__all__ = [
    # Config
    'BrandConfig',
    'load_config',
    'save_config',
    'create_minimal_config',
    # Migration
    'migrate_presentation',
    'migrate_from_content',
    'detect_and_parse',
    'detect_content_type',
    'LayoutSelector',
    # Analysis
    'analyze_presentation',
    'get_analysis_json',
    # Extraction
    'extract_pptx_to_markdown',
    # Content Document
    'ContentDocument',
    'SlideContent',
    'ContentType',
    'slides_to_content_document',
    'content_document_to_slides',
    'load_content_document',
    'save_content_document',
    # Cookbook
    'LayoutRecipe',
    'COOKBOOK',
    'get_recipe',
    'list_recipes',
    # Diagnostics
    'diagnose_template',
    'DiagnosticReport',
]
