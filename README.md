# lnblxj-skills

A collection of specialized skills for OpenClaw AI agents.

This repository contains my experimental and production-ready skills that extend agent capabilities with document generation and other domain-specific tools.

## Skills

| Skill | Description |
|-------|-------------|
| [document-generator](skills/document-generator/) | Creates professional Word, Excel, and PDF documents from templates or structured data with rich formatting, tables, images, and dynamic content |

## Skill Details

### document-generator

A comprehensive document generation system that programmatically creates professional documents.

**Capabilities:**
- **Word (.docx)**: Multi-level headings, tables, images, rich styling, headers/footers
- **Excel (.xlsx)**: Data population, cell formatting, formulas, charts
- **PDF (.pdf)**: HTML→PDF conversion, custom styles, page layout control
- **Templates**: JSON-driven content, HTML templates with Jinja2 variables

**Use cases:** Reports, invoices, manuals, certificates, data-driven documents, newsletters.

**Quick start:**
```bash
# Install dependencies
pip install python-docx openpyxl fpdf2 weasyprint jinja2

# Create a simple Word doc
python3 create_word.py "report.docx" --title "My Report" --content "Hello World"

# Create Excel from JSON data
python3 create_excel.py "data.xlsx" --data '{"Sheet1": [{"A1": "Name", "B1": "Score"}]}'

# Create PDF from HTML
python3 create_pdf.py "output.pdf" --html "<h1>Hello</h1><p>World</p>"
```

**Advanced usage with rich JSON content:**
```json
{
  "sections": [
    {
      "heading": "Section Title",
      "level": 1,
      "content": [
        {
          "type": "paragraph",
          "text": "Paragraph text",
          "bold": true
        },
        {
          "type": "table",
          "headers": ["Col1", "Col2"],
          "data": [["A", "B"], ["C", "D"]]
        }
      ]
    }
  ]
}
```

Run with:
```bash
python3 create_word.py "report.docx" --content-file "content.json" --title "Report Title"
```

## File Structure

```
skills/
└── document-generator/
    ├── SKILL.md              # Skill definition for OpenClaw
    ├── create_word.py        # Word document generator
    ├── create_excel.py       # Excel workbook generator
    ├── create_pdf.py         # PDF generator
    ├── install_dependencies.py
    ├── demo_example.py       # Example with rich features
    └── USER_GUIDE.md         # Detailed documentation
```

## Dependencies

Required Python packages:
- `python-docx` - Word document creation
- `openpyxl` - Excel workbook creation
- `fpdf2` - PDF generation (alternative backend)
- `weasyprint` - HTML to PDF conversion (high quality)
- `jinja2` - Template rendering

Install automatically:
```bash
python3 install_dependencies.py
```

Or manually:
```bash
pip install python-docx openpyxl fpdf2 weasyprint jinja2
```

**System dependencies (Linux):** For WeasyPrint, you may need:
```bash
apt install libpango-1.0-0 libgdk-pixbuf2.0-0 libffi-dev
```

## Changelog

### 1.1.0 (2026-03-15)
- Added multi-level headings, tables with custom styles, image support with captions
- Enhanced rich styling (font size, color, bold, italic, underline, alignment)
- Added headers and footers for Word documents
- Improved template system with Jinja2 variable substitution

### 1.0.0 (2026-03-14)
- Initial release with basic Word, Excel, and PDF creation
- Support for simple text content and basic data tables

---

**Note:** This skill is production-ready. See `skills/document-generator/USER_GUIDE.md` for comprehensive documentation and advanced examples.