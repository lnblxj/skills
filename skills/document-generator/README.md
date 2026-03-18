# Document Generator Skill

> Create professional Word, Excel, and PDF documents from templates or structured data.

A powerful OpenClaw skill that generates formatted documents programmatically. Supports rich styling, tables, images, multi-level headings, and dynamic content injection.

## Features

- **Word (.docx)**: Multi-level headings, tables, images, rich formatting, headers/footers
- **Excel (.xlsx)**: Data population, cell formatting, formulas, charts
- **PDF (.pdf)**: HTML→PDF conversion, custom styles, page layout control
- **Templates**: JSON-driven content, HTML templates with Jinja2 variables

## Installation

```bash
# Install dependencies
pip install python-docx openpyxl fpdf2 weasyprint jinja2

# Or use the provided script
python3 install_dependencies.py
```

## Quick Examples

```bash
# Simple Word doc
python3 create_word.py "report.docx" --title "My Report" --content "Hello World"

# Rich document from JSON
python3 create_word.py "report.docx" --content-file "content.json" --title "Q1 Summary"

# Excel with data
python3 create_excel.py "data.xlsx" --data '{"Sheet1": [{"A1": "Name", "B1": "Score"}]}'

# PDF from HTML
python3 create_pdf.py "output.pdf" --html "<h1>Hello</h1><p>World</p>"
```

## JSON Content Structure (for Word)

```json
{
  "sections": [
    {
      "heading": "Section Title",
      "level": 1,
      "content": [
        { "type": "paragraph", "text": "Paragraph text", "bold": true },
        { "type": "table", "headers": ["Col1","Col2"], "data": [["A","B"]] }
      ]
    }
  ]
}
```

## Use Within OpenClaw

This skill is designed for the OpenClaw agent framework. Once installed in your skills directory, invoke it by asking:

> "Generate a sales report in Word with a table showing region, revenue, and growth."

The agent will automatically use this skill to produce the document.

## File Layout

```
document-generator/
├── SKILL.md           # Skill definition (for OpenClaw)
├── create_word.py     # Word generator
├── create_excel.py    # Excel generator
├── create_pdf.py      # PDF generator
├── install_dependencies.py
├── demo_example.py    # Example demonstrating rich features
└── USER_GUIDE.md      # Detailed documentation
```

## Requirements

- Python 3.8+
- Packages: `python-docx`, `openpyxl`, `fpdf2`, `weasyprint`, `jinja2`
- (Optional) System libraries for WeasyPrint: `libpango-1.0-0`, `libgdk-pixbuf2.0-0`, `libffi-dev`

## License

MIT © 2026

## Changelog

- **1.1.0** (2026-03-15): Added multi-level headings, tables, images, rich styling
- **1.0.0** (2026-03-14): Initial release (basic Word/Excel/PDF creation)
