# AI-Powered Attendance Manager

Automated data entry dashboard that extracts tabular attendance data from unstructured documents (PDFs, Images, Word Docs) and securely writes it to a mapped Excel template. 

**Powered by Google Gemini 2.5 Flash.**

## 📖 Project Background
This project was originally built to solve a massive administrative bottleneck for a local church administrator. They were spending hours manually transferring weekly attendance numbers from handwritten photos, messy PDFs, and scattered Word documents into a master Excel database. 

To protect the privacy of the original organization, all specific locations, names, and internal structural references have been scrubbed from this repository and replaced with generic placeholders (`REGION_ALPHA`, `REGION_BETA`, etc.). The core logic, multimodal AI extraction pipeline, and dynamic UI remain completely intact. 

If you manage branch-based data reporting, you can easily fork this project and map it to your own Excel templates.

## ✨ Features
* **Multimodal Extraction:** Drop or paste (Ctrl+V) JPEGs, PNGs, PDFs, or DOCX files directly into the browser.
* **Smart Data Parsing:** Gemini AI reads complex, merged-header tables and normalizes the data into strict JSON.
* **Human-in-the-Loop Review:** An interactive, generated UI allows users to review and edit AI-extracted numbers before committing them to the database.
* **Direct Excel Integration:** Automatically writes confirmed data to specific protected rows/columns in an `.xlsx` template using `openpyxl`.

## 🚀 Quick Start

### 1. Requirements
* Python 3.9+
* A free [Google Gemini API Key](https://aistudio.google.com/app/apikey)

### 2. Installation
Clone the repository and install the dependencies:
```bash
git clone [https://github.com/yourusername/ai-attendance-manager.git](https://github.com/yourusername/ai-attendance-manager.git)
cd ai-attendance-manager
pip install -r requirements.txt