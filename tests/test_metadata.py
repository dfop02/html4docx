import logging
from datetime import datetime
from io import BytesIO

import pytest
from docx import Document

from html4docx import HtmlToDocx
from html4docx.metadata import Metadata


@pytest.fixture
def empty_doc():
    return Document()

@pytest.fixture
def metadata_obj(empty_doc):
    return Metadata(empty_doc)

def test_set_and_get_standard_metadata(metadata_obj):
    metadata_obj.set_metadata(author="Robert Downey Jr.", title="The Robert Success", revision="3")
    props = metadata_obj.get_metadata()
    assert props["author"] == "Robert Downey Jr."
    assert props["title"] == "The Robert Success"
    assert props["revision"] == 3

def test_invalid_revision_type(metadata_obj, caplog):
    """Invalid revision must emit a WARNING on the 'html4docx.metadata' logger, not print to stdout."""
    with caplog.at_level(logging.WARNING, logger="html4docx.metadata"):
        metadata_obj.set_metadata(revision="not_a_number")
    assert any("Invalid revision number" in r.message for r in caplog.records)

def test_invalid_datetime_string(metadata_obj, caplog):
    """Invalid ISO datetime must emit a WARNING on the 'html4docx.metadata' logger, not print to stdout."""
    with caplog.at_level(logging.WARNING, logger="html4docx.metadata"):
        metadata_obj.set_metadata(modified="2025-18-99T10:00:00")
    assert any("Invalid datetime string" in r.message for r in caplog.records)

def test_valid_datetime_string(metadata_obj):
    metadata_obj.set_metadata(modified="2025-07-18T10:00:00")
    props = metadata_obj.get_metadata()
    assert isinstance(props["modified"], datetime)


def test_unrecognized_property(metadata_obj, caplog):
    """Unrecognized core property must emit a WARNING on the 'html4docx.metadata' logger, not print to stdout."""
    with caplog.at_level(logging.WARNING, logger="html4docx.metadata"):
        metadata_obj.set_metadata(nonexistent="something")
    assert any('Property "nonexistent" not found' in r.message for r in caplog.records)

def test_print_metadata(capsys, metadata_obj):
    """get_metadata(print_result=True) should still print to stdout — this is intentional output."""
    metadata_obj.set_metadata(author="Test Author")
    metadata_obj.get_metadata(print_result=True)
    captured = capsys.readouterr()
    assert "Test Author" in captured.out

def test_get_metadata_returns_dict(metadata_obj, capsys):
    metadata_obj.set_metadata(author="Test User", title="Metadata Title")
    result = metadata_obj.get_metadata()

    assert isinstance(result, dict)
    assert result["author"] == "Test User"
    assert result["title"] == "Metadata Title"

    captured = capsys.readouterr()
    assert captured.out == ""

def test_metadata_integration_with_html4docx(empty_doc):
    docx_obj = HtmlToDocx()
    docx_obj.set_initial_attrs(empty_doc)

    metadata = docx_obj.metadata
    metadata.set_metadata(author="Jane", created="2025-07-18T09:30:00")

    buffer = BytesIO()
    docx_obj.save(buffer)
    buffer.seek(0)

    reloaded_doc = Document(buffer)
    reloaded_props = reloaded_doc.core_properties

    assert reloaded_props.author == "Jane"
    assert isinstance(reloaded_props.created, datetime)
    assert reloaded_props.created.isoformat().startswith("2025-07-18T09:30")

def test_metadata_logger_is_named():
    """The Metadata module must expose a named logger 'html4docx.metadata' so consumers
    can target it in their logging config. Issue #80."""
    import html4docx.metadata as meta_module
    assert meta_module.logger.name == "html4docx.metadata"

def test_package_logger_has_null_handler():
    """The package-level 'html4docx' logger carries a NullHandler so the library
    is silent by default. Child loggers inherit it via propagation — no NullHandler
    is needed on each module logger. Issue #80."""
    package_logger = logging.getLogger("html4docx")
    handler_types = [type(h) for h in package_logger.handlers]
    assert logging.NullHandler in handler_types, (
        "html4docx package logger should have a NullHandler attached"
    )

def test_metadata_warnings_are_silenceable(metadata_obj):
    """Consumers must be able to suppress html4docx.metadata warnings by raising its level
    without touching the root logger. Issue #80."""
    named_logger = logging.getLogger("html4docx.metadata")
    original_level = named_logger.level
    try:
        named_logger.setLevel(logging.ERROR)
        captured = []

        class _Capture(logging.Handler):
            def emit(self, record):
                captured.append(record)

        handler = _Capture(level=logging.WARNING)
        named_logger.addHandler(handler)
        try:
            metadata_obj.set_metadata(revision="bad", nonexistent="x")
        finally:
            named_logger.removeHandler(handler)

        assert captured == [], (
            "Setting html4docx.metadata to ERROR should suppress WARNING records; "
            f"got: {[r.getMessage() for r in captured]}"
        )
    finally:
        named_logger.setLevel(original_level)
