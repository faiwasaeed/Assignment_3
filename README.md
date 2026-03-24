# Jupyter Notebook to PDF AI Agent

Convert any Jupyter Notebook (`.ipynb`) into a polished, professionally formatted PDF report in one command.

---

## Features

| Feature | Details |
|---|---|
| **Markdown rendering** | Headings (H1–H4), bold, italic, lists, blockquotes, tables, inline code |
| **Code cells** | Syntax-highlighted Python blocks with `In [N]` labels |
| **Cell outputs** | Styled `stdout`/`stderr`, display data, embedded images |
| **Error tracebacks** | Clearly formatted, ANSI codes stripped |
| **Cover page** | Auto-generated with title, kernel, date, filename |
| **Table of Contents** | Auto-built from notebook headings |
| **Page headers/footers** | Title, kernel, page numbers, date |
| **Professional layout** | A4, consistent typography, colour-coded sections |

---

## Requirements

```
Python 3.8+
reportlab>=4.0
Pillow>=9.0
pygments>=2.14
```

### Install dependencies

```bash
pip install reportlab Pillow pygments
```

---

## Usage

### Basic conversion

```bash
python nb2pdf_agent.py my_notebook.ipynb
# Output: my_notebook.pdf (same directory)
```

### Specify output path

```bash
python nb2pdf_agent.py my_notebook.ipynb -o reports/lab_report.pdf
```

### Generate demo (no notebook needed)

```bash
python nb2pdf_agent.py --demo
# Creates: sample_notebook.ipynb + sample_output.pdf
```

### Full help

```bash
python nb2pdf_agent.py --help
```

---

## CLI Reference

```
usage: nb2pdf_agent.py [-h] [-o OUTPUT] [--demo] [notebook]

positional arguments:
  notebook              Path to .ipynb file

options:
  -h, --help            Show help and exit
  -o OUTPUT, --output OUTPUT
                        Output PDF path (default: <notebook_name>.pdf)
  --demo                Generate a sample notebook and convert it
```

---

## PDF Structure

```
Page 1   ── Cover page (title, kernel, timestamp, filename)
Page 2   ── Table of Contents (auto-generated from headings)
Page 3+  ── Notebook content
             ├── Markdown cells → formatted prose, headings, lists, tables
             ├── Code cells     → syntax-highlighted code blocks (In [N])
             └── Outputs        → styled output blocks (Out [N])
```

---

## Output Styling Guide

| Content type | Visual treatment |
|---|---|
| H1 headings | Large navy text + red underline rule |
| H2 headings | Blue text + grey underline rule |
| Code cells | Grey background, left blue accent bar, `In [N]` label |
| stdout output | Cream background, amber left bar, `Out [N]` label |
| stderr / errors | Pink background, red left bar, error type label |
| Images | Embedded, auto-scaled to page width |
| Tables | Striped rows, navy header row |
| Blockquotes | Left accent bar, italic grey text |

---

## Architecture

```
nb2pdf_agent.py
├── NotebookParser      — Reads .ipynb JSON, extracts cells + metadata
├── MarkdownConverter   — Converts markdown to ReportLab Paragraph flowables
├── PDFBuilder          — Assembles the PDF with cover, TOC, headers/footers
│   ├── add_markdown_cell()
│   ├── add_code_cell()
│   └── _render_output()
└── main()              — CLI argument parsing + orchestration
```

---

## Supported Output Types

| Output type | Handled |
|---|---|
| `stream` (stdout/stderr) | ✅ |
| `execute_result` (text/plain) | ✅ |
| `display_data` (image/png) | ✅ |
| `display_data` (image/jpeg) | ✅ |
| `display_data` (text/html tables) | ✅ (text extracted) |
| `error` (traceback) | ✅ |

---

## Limitations

- Very large outputs (> 60 lines) are truncated to keep PDFs readable
- HTML cell outputs are converted to plain text (CSS/JS stripped)
- Interactive widgets (ipywidgets) are rendered as placeholder text
- LaTeX math in markdown is rendered as plain text (no equation rendering)

---

## Example

```bash
# Convert a real notebook
python nb2pdf_agent.py analysis.ipynb -o Analysis_Report.pdf

# Run the built-in demo
python nb2pdf_agent.py --demo
# → sample_notebook.ipynb
# → sample_output.pdf
```
