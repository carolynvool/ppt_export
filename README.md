# ppt_export

## Overview

`ppt_export` is a small utility module that enables exporting Plotly figures from Streamlit applications into formatted PowerPoint (`.pptx`) presentations.

The module is designed to support reporting and presentation workflows where charts displayed in Streamlit need to be reused in slides, emails, or meetings without manual screenshotting or recreation.

---

## Features

- exports Plotly figures directly to PowerPoint slides
- supports multiple figures across multiple slide sections
- automatically generates:
  - title slide
  - section headers
  - consistent slide layout and formatting
- preserves chart resolution using static image export (via `kaleido`)
- integrates cleanly with Streamlit via a single export button

---

## Design Principles

- **Non-intrusive**  
  export functionality is optional and does not affect existing page behavior

- **Separation of concerns**  
  chart creation, UI rendering, and export logic are kept separate to comply with Streamlit caching rules and to improve maintainability

- **Reusability**  
  the module is page-agnostic and can be reused across different Streamlit pages with minimal setup

- **Future-proofing**  
  designed to be easily extended (e.g. PDF export, different templates, branding)

---

## Usage

Typical usage pattern in a Streamlit page:

```python
from utils.ppt_export import create_export_button

# figuresCol1 and figuresCol2 are lists of Plotly figures
create_export_button(
    figures_col1=figuresCol1,
    figures_col2=figuresCol2,
    page_name="Performance"
)
