# Pixels to Paperwork

**Managing hundreds of invoices, receipts, ID proofs, or business cards  Images manually is time‑consuming and error‑prone. Extracting specific details like dates, GST numbers, invoice IDs, or contact information from unstructured images often requires hours of repetitive work. This tool solves that problem by combining Windows OCR with Python automation and a Granite LLM model. The workflow scans entire folders of images, extracts text, and then intelligently structures the information into a clean Excel file. Regex handles predictable fields like phone numbers and emails, while Granite interprets contextual data such as names, addresses, or company details. The result is a semi‑automated pipeline that reduces manual effort by up to 90%, delivering structured datasets ready for reporting or analysis.**

**Limitations**
- Works only on Python 3.10.19 (Windows OCR requirement).

- IMG Hyperlinks are local (won’t open if Excel is shared).

- Accuracy is ~70–95%, Depends on Situation.

- For PDFs or any other format, you’d need a preprocessing step (convert PDF/or other format into → JPG).

# Technical_Explanations

**This guide explains the logic behind every code cell so you can easily follow the project from start to finish.**

## 📌 Cell 1 – Environment Setup
This cell checks the Python environment and confirms that the correct version is running.

- **Why it matters:**
  - Windows OCR libraries only work with Python `3.10.19`.  
  - Ensures you are inside the correct Conda environment before running OCR.

- **Where to modify:**
  - No changes needed unless you want to run in a different environment.  
  - If OCR fails, double‑check that your Python version matches `3.10.19`.

```python
import sys
import os

print("Python Version:", sys.version)
print("Environment:", os.environ.get("CONDA_DEFAULT_ENV"))
print("Executable Path:", sys.executable)
```
