"""JSON Schema definitions for Presentation Toolkit."""

from pathlib import Path

SCHEMA_DIR = Path(__file__).parent


def get_content_schema_path() -> Path:
    """Return the path to the content.schema.json file."""
    return SCHEMA_DIR / "content.schema.json"
