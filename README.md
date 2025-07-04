# AUTO_PPT_SURAJDESIGNER

Automated PowerPoint Generation from Image-based Questions using OCR

This project provides a Python script to process scanned question images, extract each question using Tesseract OCR, and generate a PowerPoint presentation with each question as a separate slide. It is ideal for educators and content creators who want to digitize and present question papers efficiently.

## Features
- Batch process a folder of question images
- Uses Tesseract OCR for accurate text extraction
- Automatically detects and crops individual questions
- Generates a clean PowerPoint presentation (PPTX) with each question on a separate slide
- Easy to configure and extend

## Requirements
- Python 3.7+
- [Tesseract-OCR](https://github.com/tesseract-ocr/tesseract) (Windows installer recommended)
- Python packages: `opencv-python`, `pillow`, `python-pptx`, `pytesseract`

## Installation
1. **Clone this repository:**
   ```bash
   git clone https://github.com/designersuraj/AUTO_PPT_SURAJDESIGNER.git
   cd AUTO_PPT_SURAJDESIGNER
   ```
2. **Install dependencies:**
   ```bash
   pip install -r requirements.txt
   ```
3. **Install Tesseract-OCR:**
   - Download and install from [here](https://github.com/tesseract-ocr/tesseract)
   - Make sure the path in the script matches your installation location

## Usage
1. Place your `.jpg` question images in the `SUBJECT` folder.
2. Edit the paths in `AUTO_PPT_SURAJDESIGNER.py` if needed.
3. Run the script:
   ```bash
   python AUTO_PPT_SURAJDESIGNER.py
   ```
4. The output PowerPoint will be saved as `PPT OUTPUT BY SUBJECT.pptx`.

## Author
**Suraj (designersuraj)**  
[Website: surajdesigner.com](https://surajdesigner.com)
[GitHub: designersuraj](https://github.com/designersuraj)

---
For more information, visit [surajdesigner.com](https://surajdesigner.com) 