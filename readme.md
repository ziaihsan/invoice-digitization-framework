# Full FATURA Dataset Source

This directory is intended to act as the container for the complete **FATURA Dataset** (10,000+ invoice images and annotations).

Due to repository size constraints, the full dataset is **not included** in this repository. The experiments in this paper rely on a controlled subset (Template 1) which is already provided in the parent `../images` and `../annotations` directories.

## How to Enable Full Dataset Processing

If you wish to use the `dataset_refresher.py` script to sample new random data for training or validation, you must download the full dataset and extract it into this directory.

**Download Source:**
* **Provider:** Zenodo
* **Link:** https://zenodo.org/records/8261508
* **DOI:** 10.5281/zenodo.8261508

## Expected Directory Structure

After downloading and extracting the dataset, ensure the structure matches the tree below so that `dataset_refresher.py` can locate the files correctly:

```
data/
├── Full_invoices_dataset/       <-- (You are here)
│   ├── Annotations/
│   │   └── Original_Format/     <-- IMPORTANT: JSON files go here
│   │       ├── Template10_Instance0.json
│   │       └── ...
│   └── images/                  <-- JPG files go here
│       ├── Template10_Instance0.jpg
│       └── ...
├── annotations/                 <-- (Active subset used by the framework)
├── images/                      <-- (Active subset used by the framework)
├── dataset_refresher.py         <-- Script to sample from Full_invoices_dataset
└── convert_json_to_csv.py




It should be like this after:
.
├── data
│   ├── Full_invoices_dataset
│   │   ├── Annotations
│   │   │   ├── .virtual_documents
│   │   │   └── Original_Format
│   │   │       ├── Template10_Instance0.json
│   │   │       ├── Template10_Instance1.json
│   │   │       ├── Template10_Instance10.json
│   │   │       ├── Template10_Instance100.json
│   │   │       ├── Template10_Instance101.json
│   │   │       └── ... (skipped 9995 more .json files)
│   │   ├── images
│   │   │   ├── Template10_Instance0.jpg
│   │   │   ├── Template10_Instance1.jpg
│   │   │   ├── Template10_Instance10.jpg
│   │   │   ├── Template10_Instance100.jpg
│   │   │   ├── Template10_Instance101.jpg
│   │   │   └── ... (skipped 9995 more .jpg files)
│   │   └── source.txt
│   ├── annotations
│   │   ├── Template1_Instance0.json
│   │   ├── Template1_Instance1.json
│   │   ├── Template1_Instance10.json
│   │   ├── Template1_Instance100.json
│   │   ├── Template1_Instance101.json
│   │   └── ... (skipped 195 more .json files)
│   ├── convert_json_to_csv.py
│   ├── dataset_refresher.py
│   ├── images
│   │   ├── Template1_Instance0.jpg
│   │   ├── Template1_Instance1.jpg
│   │   ├── Template1_Instance10.jpg
│   │   ├── Template1_Instance100.jpg
│   │   ├── Template1_Instance101.jpg
│   │   └── ... (skipped 195 more .jpg files)
│   │
│   └── trash <-- Archival folder for previous batch files displaced by dataset_refresher.py
│
├── excel_rpa.py
├── ocr_to_csv.py
├── requirements.txt
└── result
    ├── excel_rpa.log
    ├── ocr_results.csv
    └── validate
        ├── ground_truth
        │   └── annotations.csv
        ├── log.txt
        ├── validate_ocr.py
        ├── validation_report.html
        └── validation_report.json
```


## Installation & Setup

### Prerequisites

- **Python 3.13.9** or higher
- **Tesseract OCR** engine (required for OCR functionality)
- **Microsoft Excel** (required for RPA/automation features)

### 1. Install Tesseract OCR

**macOS (Homebrew):**
```bash
brew install tesseract
```

**Ubuntu/Debian:**
```bash
sudo apt update
sudo apt install tesseract-ocr
```

**Windows:**
Download the installer from [UB-Mannheim/tesseract](https://github.com/UB-Mannheim/tesseract/wiki) and add to PATH.

### 2. Install Python Dependencies

Using requirements.txt (recommended):
```bash
pip3 install -r requirements.txt
```

Or install manually:
```bash
pip3 install Pillow==10.4.0 pytesseract==0.3.13 xlwings==0.30.14
```

**macOS additional dependency:**
```bash
pip3 install appscript==1.2.5
```

### 3. Verify Installation

```bash
# Check Python version
python3 --version

# Check Tesseract installation
tesseract --version

# Check installed packages
pip3 list | grep -E "(Pillow|pytesseract|xlwings|appscript)"
```

### Platform Notes

| Dependency | Version | Notes |
|------------|---------|-------|
| Python | 3.13.9 | Tested on macOS Tahoe 26.1 (Apple Silicon) |
| Pillow | 10.4.0 | Image processing for OCR preprocessing |
| pytesseract | 0.3.13 | Python wrapper for Tesseract OCR |
| xlwings | 0.30.14 | Excel automation bridge |
| appscript | 1.2.5 | macOS only (`sys_platform == "darwin"`) |

**Important:**
- For `xlwings` to work properly, use Microsoft Excel downloaded from [microsoft.com](https://www.microsoft.com), **not** the Mac App Store version.
- Tested with Microsoft Excel Version 16.102.3 (25110228).
