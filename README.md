# AI Presentation Toolkit

A generic, brand-agnostic toolkit for migrating PowerPoint presentations to branded templates. Supports intelligent content-type detection, layout variety tracking, and brand compliance analysis.

## Brand-Specific Skills

This toolkit provides the engine for brand-specific presentation skills:

- **[drupal-brand-skill](https://github.com/zivtech/drupal-brand-skill)** - Drupal brand guidelines with presentation migration support

Brand skills use this toolkit's generic capabilities with their specific configurations (colors, fonts, layouts, templates).

## Features

- **Content Extraction**: Extract text and images from PPTX/PDF files
- **Intelligent Migration**: Auto-detect content types (stats, quotes, bullets, etc.) and map to appropriate template layouts
- **Layout Variety**: Prevents consecutive layout repeats for visual interest
- **Brand Compliance Analysis**: Scan presentations for off-brand fonts and colors
- **Configurable**: All brand-specific values (colors, fonts, slide catalog) loaded from YAML config
- **content.json Intermediate Format**: Validated JSON representation between parsing and generation — inspect, edit, and feed content from external tools
- **Template Diagnostics**: Pre-flight checks catch missing placeholders, broken references, and slide count mismatches before migration starts
- **Slide Layout Cookbook**: Explicit positioning recipes with automatic fallback when template placeholders are missing

## Installation

```bash
pip install -e .
```

With optional JSON Schema validation:

```bash
pip install -e ".[schema]"
```

Or use directly:

```bash
python -m presentation_toolkit.cli migrate input.pptx output.pptx --config brand.yaml
```

## Quick Start

### 1. Create a Brand Configuration

See `examples/sample_brand_config.yaml` for a complete example.

```yaml
version: "1.0"
brand_name: "My Brand"

colors:
  primary:
    brand_blue: "0066CC"
  secondary:
    white: "FFFFFF"
    black: "000000"

fonts:
  brand: ["Brand Font", "Noto Sans"]
  replace: ["Arial", "Calibri", "Helvetica"]
```

### 2. Migrate a Presentation

```bash
# Using CLI
pptx-migrate input.pptx output.pptx --config my-brand.yaml --template template.pptx

# Save intermediate content.json for inspection
pptx-migrate input.pptx output.pptx --config my-brand.yaml --save-content content.json

# Migrate from a pre-built content.json (skip parsing)
pptx-migrate --from-content content.json output.pptx --config my-brand.yaml

# Force cookbook mode (absolute positioning, no template placeholders)
pptx-migrate input.pptx output.pptx --config my-brand.yaml --use-cookbook

# Using Python
from presentation_toolkit import migrate_presentation, load_config

config = load_config("my-brand.yaml")
migrate_presentation("input.pptx", "output.pptx", config, "template.pptx")
```

### 3. Analyze Brand Compliance

```bash
pptx-analyze presentation.pptx --config my-brand.yaml
```

### 4. Extract Content

```bash
pptx-extract input.pptx --output content.md --images
```

### 5. Diagnose a Template

```bash
# Human-readable report
pptx-diagnose template.pptx --config brand.yaml

# JSON output for CI/automation
pptx-diagnose template.pptx --config brand.yaml --json

# Strict mode (exit code 1 on blocking errors)
pptx-diagnose template.pptx --strict
```

## content.json Format

The intermediate `content.json` format provides a validated, editable representation of presentation content between the parsing and generation phases.

```json
{
  "version": "1.0",
  "metadata": {
    "source_file": "input.pptx",
    "source_format": "pptx",
    "generated_at": "2025-01-15T12:00:00Z",
    "generator": "ai-presentation-toolkit"
  },
  "slides": [
    {
      "number": 1,
      "title": "Welcome",
      "body": "Subtitle text",
      "content_type": "title_opening",
      "layout_hint": "title_opening",
      "images": [],
      "zones": null,
      "extraction_notes": []
    },
    {
      "number": 2,
      "title": "Key Metrics",
      "body": "",
      "content_type": "stats_dashboard",
      "zones": {
        "type": "stats_dashboard",
        "stats": [
          {"number": "85%", "label": "Customer Satisfaction"},
          {"number": "$2.5M", "label": "Revenue Growth"}
        ]
      }
    }
  ]
}
```

### Supported content types

`auto`, `statistic`, `stats_dashboard`, `quote`, `numbered_step`, `bullet_list`, `comparison`, `section_header`, `case_study`, `case_study_full`, `statement`, `feature`, `detailed_content`, `title_opening`, `closing`

### Working with content.json in Python

```python
from presentation_toolkit import (
    slides_to_content_document,
    content_document_to_slides,
    save_content_document,
    load_content_document,
)

# Convert legacy slides to content document
doc = slides_to_content_document(slides, "input.pptx", "pptx")

# Save to JSON
save_content_document(doc, "content.json")

# Load and modify
doc = load_content_document("content.json")
doc.slides[0].content_type = "title_opening"

# Convert back to slides for migration
slides = content_document_to_slides(doc)
```

## Cookbook Fallback

When template placeholders are missing, the toolkit automatically falls back to the layout cookbook — a set of 13 built-in recipes that position content using absolute coordinates.

Available recipes: `title_opening`, `feature_default`, `content_image_left`, `content_image_right`, `statement_center`, `stat_default`, `stats_dashboard`, `quote_default`, `two_column`, `section_divider`, `case_study_full`, `closing_cta`, `hero_photo`

```python
from presentation_toolkit import get_recipe, list_recipes, COOKBOOK

# List all available recipes
print(list_recipes())

# Get a specific recipe
recipe = get_recipe("stats_dashboard")
print(recipe.description)
```

## Configuration Reference

See the [Configuration Schema](src/presentation_toolkit/config/schema.py) for full details.

### Required Fields

- `brand_name`: Name of the brand
- `colors`: Color palette with hex values (without #)
- `fonts.brand`: List of approved brand fonts
- `fonts.replace`: List of fonts to flag for replacement

### Optional Fields

- `template.default`: Path to default template PPTX
- `slide_catalog`: Mapping of content types to template slide indices
- `text_capacity`: Character limits and font sizes per layout type
- `content_patterns`: Regex patterns for content type detection

## CLI Commands

### `pptx-migrate`

Migrate a presentation to a branded template.

```bash
pptx-migrate input.pptx output.pptx --config brand.yaml [options]

Options:
  --template PATH       Path to template PPTX (overrides config)
  --no-images           Skip image extraction/insertion
  --save-content PATH   Save intermediate content.json before migration
  --from-content PATH   Skip parsing, use pre-built content.json
  --use-cookbook         Force cookbook mode (absolute positioning)
  --verbose             Show detailed progress
```

### `pptx-analyze`

Analyze a presentation for brand compliance.

```bash
pptx-analyze presentation.pptx --config brand.yaml [options]

Options:
  --json             Output results as JSON
  --strict           Fail on any compliance issue
```

### `pptx-extract`

Extract content from a presentation.

```bash
pptx-extract input.pptx [options]

Options:
  --output PATH      Output markdown file (default: input.md)
  --images           Also extract images to folder
```

### `pptx-diagnose`

Run template diagnostics.

```bash
pptx-diagnose template.pptx [options]

Options:
  --config PATH      Brand configuration file (YAML/JSON)
  --strict           Exit with error if blocking issues found
  --json             Output results as JSON
```

Diagnostic codes:

| Code | Severity | Description |
|------|----------|-------------|
| TMPL-001 | ERROR | Template file not found |
| TMPL-002 | ERROR | Not a valid ZIP/PPTX |
| TMPL-003 | ERROR | Missing required OOXML parts |
| TMPL-004 | WARNING | Config references slide index not in template |
| TMPL-010 | WARNING | Slide missing TITLE placeholder |
| TMPL-011 | WARNING | Slide missing BODY placeholder |
| TMPL-012 | INFO | Slide has no PICTURE placeholder |
| TMPL-020 | WARNING | stats_dashboard slide missing named shapes |
| TMPL-021 | WARNING | case_study_full slide missing Quote/Attribution shapes |
| TMPL-030 | WARNING | Template has >3 slide masters |
| TMPL-040 | WARNING | Relationship references missing media file |
| TMPL-050 | INFO | Shape covers >80% of slide area |

## Python API

```python
from presentation_toolkit import (
    load_config,
    migrate_presentation,
    migrate_from_content,
    analyze_presentation,
    extract_pptx_to_markdown,
    diagnose_template,
    # Content document
    ContentDocument,
    SlideContent,
    ContentType,
    slides_to_content_document,
    content_document_to_slides,
    load_content_document,
    save_content_document,
    # Cookbook
    LayoutRecipe,
    COOKBOOK,
    get_recipe,
    list_recipes,
    # Diagnostics
    DiagnosticReport,
)

# Load brand configuration
config = load_config("brand.yaml")

# Migrate a presentation
migrate_presentation(
    input_path="source.pptx",
    output_path="branded.pptx",
    config=config,
    template_path="template.pptx",
    insert_images=True,
    diagnose=True,
)

# Migrate from content.json
migrate_from_content("content.json", "output.pptx", config, "template.pptx")

# Analyze for brand compliance
issues = analyze_presentation("deck.pptx", config)
for slide in issues:
    print(f"Slide {slide['num']}: {slide['issues']}")

# Extract content
slides = extract_pptx_to_markdown("source.pptx", "content.md", extract_images=True)

# Diagnose template
report = diagnose_template("template.pptx", config)
report.print_report()
if report.has_blocking_issues:
    print("Fix template before migration!")
```

## License

MIT
