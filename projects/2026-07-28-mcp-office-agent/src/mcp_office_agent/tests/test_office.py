import os
import tempfile
import pytest

from mcp_office_agent.word import write_word_file, read_word_file
from mcp_office_agent.excel import write_excel_file, read_excel_file
from mcp_office_agent.powerpoint import write_powerpoint_file, read_powerpoint_file
from mcp_office_agent.server import inspect_office_file


def test_word_flow():
    """Tests writing and reading a Word (.docx) file."""
    with tempfile.TemporaryDirectory() as tmpdir:
        file_path = os.path.join(tmpdir, "test_doc.docx")
        
        content = [
            {"type": "heading", "level": 1, "text": "Overview Section"},
            {"type": "paragraph", "text": "This is a test paragraph for Word."},
            {"type": "list_item", "style": "List Bullet", "text": "Bullet point 1"},
            {"type": "list_item", "style": "List Bullet", "text": "Bullet point 2"},
            {"type": "table", "headers": ["Header A", "Header B"], "rows": [["Cell A1", "Cell B1"], ["Cell A2", "Cell B2"]]}
        ]
        
        # Write doc
        write_word_file(file_path, content, title="Document Title")
        assert os.path.exists(file_path)
        
        # Read doc
        md_content = read_word_file(file_path)
        
        assert "# Document Title" in md_content
        assert "# Overview Section" in md_content
        assert "This is a test paragraph for Word." in md_content
        assert "- Bullet point 1" in md_content
        assert "| Header A | Header B |" in md_content
        assert "| Cell A1 | Cell B1 |" in md_content
        
        # Test inspection
        inspection = inspect_office_file(file_path)
        assert "Word Document" in inspection
        assert "Total Paragraphs:" in inspection


def test_excel_flow():
    """Tests writing and reading an Excel (.xlsx) file."""
    with tempfile.TemporaryDirectory() as tmpdir:
        file_path = os.path.join(tmpdir, "test_sheet.xlsx")
        
        sheets_data = {
            "Summary": [
                ["Item", "Cost"],
                ["Servers", 1500],
                ["SaaS", 500],
                ["Total", "=SUM(B2:B3)"]
            ],
            "Details": [
                ["ID", "Description"],
                [101, "AWS EC2 Instance"],
                [102, "Google Workspace"]
            ]
        }
        
        # Write sheet
        write_excel_file(file_path, sheets_data)
        assert os.path.exists(file_path)
        
        # Read sheet with formulas (read_formulas=True)
        md_formulas = read_excel_file(file_path, read_formulas=True)
        assert "### Sheet: Summary" in md_formulas
        assert "| Item | Cost |" in md_formulas
        assert "| Total | =SUM(B2:B3) |" in md_formulas
        assert "### Sheet: Details" in md_formulas
        assert "| 101 | AWS EC2 Instance |" in md_formulas
        
        # Test inspection
        inspection = inspect_office_file(file_path)
        assert "Excel Workbook" in inspection
        assert "Sheet 'Summary'" in inspection
        assert "Sheet 'Details'" in inspection


def test_powerpoint_flow():
    """Tests writing and reading a PowerPoint (.pptx) file."""
    with tempfile.TemporaryDirectory() as tmpdir:
        file_path = os.path.join(tmpdir, "test_slides.pptx")
        
        slides_data = [
            {
                "layout": "title_slide",
                "title": "Welcome to my Presentation",
                "subtitle": "Created by the Antigravity Team"
            },
            {
                "layout": "title_and_content",
                "title": "Key Agenda",
                "content": ["Highlight 1", "Highlight 2", "Highlight 3"],
                "notes": "Speaker prompt: emphasize slide highlights."
            }
        ]
        
        # Write presentation
        write_powerpoint_file(file_path, slides_data)
        assert os.path.exists(file_path)
        
        # Read presentation
        md_content = read_powerpoint_file(file_path)
        
        assert "## Slide 1" in md_content
        assert "Welcome to my Presentation" in md_content
        assert "Created by the Antigravity Team" in md_content
        
        assert "## Slide 2" in md_content
        assert "Key Agenda" in md_content
        assert "Highlight 1" in md_content
        assert "Speaker prompt: emphasize slide highlights." in md_content
        
        # Test inspection
        inspection = inspect_office_file(file_path)
        assert "PowerPoint Presentation" in inspection
        assert "Total Slides: 2" in inspection
        assert "Slide 1: Welcome to my Presentation" in inspection
        assert "Slide 2: Key Agenda" in inspection
