# Customer Satisfaction Analysis Tool 📊

A web application for analyzing customer satisfaction data, with an interactive dashboard and automated report generation.

## Features
- 📤 Upload your own customer satisfaction dataset
- 📈 Interactive dashboard with bar, line, radar, pie, doughnut, polar area charts, and a correlation heatmap (Chart.js)
- 📝 Auto-generates a Word report (.docx) with embedded chart images
- 📊 Auto-generates a PowerPoint presentation with embedded charts
- 🌐 Simple web interface built with Flask

## Tech Stack
- **Backend:** Python, Flask
- **Dashboard charts:** Chart.js (runs client-side in the browser)
- **Report charts:** Matplotlib (server-side, Agg backend, rendered to in-memory PNG buffers)
- **Word report generation:** python-docx, with charts inserted via `doc.add_picture()`
- **Presentation generation:** `presentation.py`

## Project Structure
- `app.py` — Main Flask application
- `analysis.py` — Core data analysis logic
- `report.py` — Word report generation with embedded Matplotlib charts
- `presentation.py` — PowerPoint generation
- `generate_sample.py` — Generates sample data for testing

## Getting Started

1. Clone the repository