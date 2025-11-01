import pytest
from io import BytesIO
from docx import Document
from html4docx.metadata import Metadata
from html4docx import HtmlToDocx
from datetime import datetime

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

def test_invalid_revision_type(metadata_obj, capsys):
    metadata_obj.set_metadata(revision="not_a_number")
    captured = capsys.readouterr()
    assert "Invalid revision number" in captured.out

def test_invalid_datetime_string(metadata_obj, capsys):
    metadata_obj.set_metadata(modified="2025-18-99T10:00:00")
    captured = capsys.readouterr()
    assert "Invalid datetime string" in captured.out

def test_valid_datetime_string(metadata_obj):
    metadata_obj.set_metadata(modified="2025-07-18T10:00:00")
    props = metadata_obj.get_metadata()
    assert isinstance(props["modified"], datetime)

def test_unrecognized_property(metadata_obj, capsys):
    metadata_obj.set_metadata(nonexistent="something")
    captured = capsys.readouterr()
    assert 'Property "nonexistent" not found' in captured.out

def test_print_metadata(capsys, metadata_obj):
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
