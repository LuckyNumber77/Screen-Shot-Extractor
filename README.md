# Screenshot Extractor

Extracts video information from screenshots using OCR (Optical Character Recognition) and saves it to an Excel file.

## Features

- Processes PNG/JPG screenshots containing video information
- Extracts:
  - Video titles/descriptions (e.g., "Top 10 bump set")
  - PXL file numbers (e.g., "PXL_20250524_160854612.mp4")
- Maintains a running Excel database of extracted information
- Handles duplicate detection to prevent repeated entries

## Requirements

- Python 3.8+
- Tesseract OCR (install from [UB Mannheim](https://github.com/UB-Mannheim/tesseract/wiki))
- Python packages:
