# PDF to WooCommerce Product Converter

This automation script converts PDF documents into web-ready images and generates a CSV file formatted specifically for WooCommerce product imports.

It is designed for educational resources, parsing filenames to automatically populate product attributes like Subject, Grade, Term, and Document Type.

## Features

* **PDF to Image:** Converts PDF pages into high-quality PNGs.
* **Smart Resizing:** Resizes images to a square canvas (default 500x500px) with transparent padding, ideal for web thumbnails.
* **Preview Generation:** Automatically generates a square preview thumbnail from the first page.
* **CSV Generation:** Outputs a WooCommerce-ready CSV string containing SKU, Name, Categories, Attributes, and Image paths.
* **Metadata Parsing:** Extracts product details directly from the filename (e.g., `aeat_gr10...` becomes "Afrikaans EAT, Grade 10").

## Prerequisites

### 1. Python 3.10+
Ensure you have Python installed.

### 2. Poppler (Crucial)
This script uses `pdf2image`, which requires Poppler to be installed on your system.

* **Windows:**
    1.  Download the latest binary from [Poppler for Windows](https://github.com/oschwartz10612/poppler-windows/releases/).
    2.  Extract the ZIP (e.g., to `C:\poppler-xx`).
    3.  Add `C:\poppler-xx\bin` to your System PATH environment variable.
    4.  Verify by running `pdftoppm -h` in CMD.

* **Mac:** `brew install poppler`
* **Linux:** `sudo apt install poppler-utils`

## Installation

1.  **Download:**
    * Click the green **<> Code** button at the top of this page.
    * Select **Download ZIP**.
    * Extract the ZIP file to a folder on your computer.

2.  **Install Dependencies:**
    Open your terminal in the extracted folder and run:

        pip install -r requirements.txt

## Usage

The script takes a source directory of PDFs and outputs images to a destination folder. It prints the CSV data to the console, which you should save to a file using the `>` operator.

    python convert.py <source_folder> <output_folder> <size> > <csv_filename.csv>

### Example

    python convert.py "C:\MyPDFs" "C:\ProcessedImages" 500 > products_import.csv

* **source_folder**: Folder containing your `.pdf` files.
* **output_folder**: Where the `.png` images will be saved (organized by subfolders).
* **size**: Pixel dimension for the square images (default: 500).
* **> filename.csv**: Saves the text output to a CSV file.

## Filename Convention

To ensure the CSV is generated with the correct attributes, your input PDFs must follow this naming structure:

`subject_grade_year_term_type_extra.pdf`

**Example:** `aeat_gr10_2025_t1_ass_lesson1.pdf`
* **Subject:** Afrikaans EAT (`aeat`)
* **Grade:** Gr. 10 (`gr10`)
* **Year:** 2025
* **Term:** Term 1 (`t1`)
* **Type:** Assessment Task (`ass`)
* **Extra:** Lesson 1

### Code Reference

**Subjects:**
* `aeat` / `aat`: Afrikaans
* `efal` / `ehl`: English
* `ma` / `me`: Mathematics
* `nw` / `ns`: Natural Sciences
* `vka`: Visual Arts
* `amus`: Music

**Grades:**
* `gr1` to `gr12`: Grade 1 to 12
* `gr49`: Gr. 4-9
* `junior`: Junior Phase
* `senior`: Senior Phase

**Terms:**
* `t1` / `k1`: Term 1
* `t2` / `k2`: Term 2 (etc.)

**Types:**
* `ass`: Assessment Task
* `wb`: Workbook / Study Guide
* `ppt`: PowerPoint
* `kv`: Short Story
* `her`: Revision
