# Pixels to Paperwork

**Managing hundreds of invoices, receipts, ID proofs, or business cards  Images manually is time‑consuming and error‑prone. Extracting specific details like dates, GST numbers, invoice IDs, or contact information from unstructured images often requires hours of repetitive work. This tool solves that problem by combining Windows OCR with Python automation and a Granite LLM model. The workflow scans entire folders of images, extracts text, and then intelligently structures the information into a clean Excel file. Regex handles predictable fields like phone numbers and emails, while Granite interprets contextual data such as names, addresses, or company details. The result is a semi‑automated pipeline that reduces manual effort by up to 90%, delivering structured datasets ready for reporting or analysis.**

**Limitations**
- Works only on Python 3.10.19 (Windows OCR requirement).

- IMG Hyperlinks are local (won’t open if Excel is shared).

- Accuracy is ~70–95%, Depends on Situation.

- For PDFs or any other format, you’d need a preprocessing step (convert PDF/or other format into → JPG).

# Technical_Explanations

**This guide explains the logic behind every code cell so you can easily follow the project from start to finish.**

## Cell 1 – Environment Setup
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
##  Cell 2 – Import Libraries
This cell imports all required Python libraries and confirms OCR readiness.

- **Why it matters:**
  - These libraries are the backbone of OCR processing.  
  - Without them, the script cannot read text from images.

- **Where to modify:**
  - Add/remove libraries depending on your project.  
  - Example: If you want to process PDFs, you’ll need an extra library like `pdf2image`.

```python
import numpy as np
import re
import os
import pandas as pd
import asyncio
import ollama
import json
from winrt.windows.media.ocr import OcrEngine
from winrt.windows.graphics.imaging import BitmapDecoder
from winrt.windows.storage import StorageFile
from winrt.windows.globalization import Language

print("GOOD TO GO - WINDOWS OCR READY")
```


##  Cell 3 – OCR Function Definition
This cell defines the **async function** that performs OCR on a single image.

- **What it does:**
  - Loads an image file from the given path.
  - Decodes it into a bitmap format.
  - Passes it to the Windows OCR engine.
  - Returns the recognized text.

- **Why it matters:**
  - This is the **core OCR function**. Every image will be processed through this function.  
  - Without this, you cannot extract text from images.

- **Where to modify:**
  - Change the language code inside `Language("en")` if you need OCR in another language (e.g., `"hi"` for Hindi).  
  - Add error handling if you expect corrupted or unreadable images.

```python
#SETUP EVERYTHING BEFORE WE START WORKING
async def ocr_image_windows(image_path):
    file = await StorageFile.get_file_from_path_async(image_path)
    stream = await file.open_async(1)

    decoder = await BitmapDecoder.create_async(stream)
    software_bitmap = await decoder.get_software_bitmap_async()

    engine = OcrEngine.try_create_from_language(Language("en"))
    result = await engine.recognize_async(software_bitmap)

    return result.text

print("Setup Is Ready")
```


##  Cell 4 – Folder Path & Image Types
This cell sets the folder path and checks supported image formats.

- **What it does:**
  - Defines the folder path where all images are stored.
  - Defines supported extensions (`.jpg`, `.jpeg`, `.png`).
  - Reads all images from the folder and sorts them.
 
- **Why it matters:**
  - Ensures the script knows where to look for images.  
  - Sorting helps keep results consistent.

- **Where to modify:**
  - Update `image_folder` to your actual folder path.  
  - Add more extensions if needed (e.g., `.tiff`, `.bmp`).  
  - If you’re processing invoices, make sure all files are in the same folder.

```python
# 2. IMAGE FOLDER PATH
# ----------------------------
image_folder = r"C:\Users\admin\Desktop\your_path"

# ----------------------------
# 3. SUPPORTED IMAGE TYPES
# ----------------------------
supported_ext = (".jpg", ".jpeg", ".png") # add more if needed

# ----------------------------
# 4. READ & SORT IMAGE FILES
# ----------------------------
all_images = sorted([
    img for img in os.listdir(image_folder)
    if img.lower().endswith(supported_ext)
])

print("GOOD TO GO")
```


##  Cell 5 – Image Count Summary
This cell counts how many images are found and groups them by type.

- **Why it matters:**
  - Quick validation step: confirms that the folder contains the expected number of images.  
  - Helps catch mistakes (e.g., wrong folder path, unsupported file types).

- **Where to modify:**
  - No major changes needed.  
  - If you want more detailed reporting (e.g., file sizes, corrupted files), you can extend this logic.

```python
# 5. IMAGE COUNT SUMMARY
# ----------------------------
jpg_count = sum(1 for i in all_images if i.lower().endswith(".jpg"))
jpeg_count = sum(1 for i in all_images if i.lower().endswith(".jpeg"))
png_count = sum(1 for i in all_images if i.lower().endswith(".png"))

total_images = len(all_images)

print(f"Total {total_images} photos found")
print(f"JPG : {jpg_count}")
print(f"JPEG: {jpeg_count}")
print(f"PNG : {png_count}")
print("-" * 40)
```


##  Cell 6 – OCR Processing for All Images
This cell loops through all images in the folder, applies OCR, and stores results in a list.

- **What it does:**
  - Iterates over each image in the folder.
  - Calls the `ocr_image_windows()` function (defined in Cell 3).
  - Extracts text and stores it in a dictionary with:
    - `card` (serial number),
    - `photo_name` (file name),
    - `all_text_info` (OCR text).
  - Collects all dictionaries into a list (`records_5`).

- **Why it matters:**
  - This is the step where raw text is actually extracted from every image.
  - Without this, the DataFrame in later cells would be empty.

- **Where to modify:**
  - Add error handling if you expect unreadable images.
  - If you want to log OCR failures separately, extend the `except` block.

```python
# ----------------------------
# 6. OCR EACH IMAGE (Windows OCR)
# ----------------------------
records_5 = []

for idx, image_name in enumerate(all_images, start=1):
    card_no = f"card_{idx}"
    image_path = os.path.join(image_folder, image_name)

    print(f"Processing {card_no} -> {image_name}")

    try:
        text = await ocr_image_windows(image_path)
        text = text.strip()
    except Exception as e:
        text = ""
        print(f"Error reading {image_name}: {e}")

    records_5.append({
        "card": card_no,
        "photo_name": image_name,
        "all_text_info": text
    })

print("-" * 40)
print("WINDOWS OCR completed for all images.")
```


##  Cell 7 & 8 – Convert into DataFrame & Preview DataFrame
This step shows the first 10 rows of the DataFrame after Conversation.

- **What it does:**
  - Creates a DataFrame called `business_card` from `records_5`.
  - Displays sample OCR results.
  
- **Why it matters:**
  - Lets you visually inspect OCR output before moving to regex/model extraction.
  - Helps catch formatting issues early.
  - Change the variable name if you’re processing invoices or receipts (e.g., `invoice_data`).

- **Example Output:**

| card   | photo_name                              | all_text_info                                                                 |
|--------|-----------------------------------------|-------------------------------------------------------------------------------|
| card_1 | Visiting Card Kannur_page-0001.jpg      | ALMA THALASSERY and more..                                                    |
| card_2 | Visiting Card Kannur_page-0002.jpg      | ALBERT AUGUSTINE Managing director +91 9961 75…                               |

```python
#CREATE PANDAS DATAFRAME
business_card = pd.DataFrame(records_5)

# just to make sure ocr worked well or not?
business_card.head(10)
```


#  Next Step – Granite Model Setup (Cell 10)
Before moving to regex and structured extraction, we need to set up the Granite model via **Ollama**.

- **Why we use Granite (LLM):**
  - Regular expressions alone cannot reliably extract structured info from un‑uniform images.
  - Granite (or other Open Source LLMs) can interpret text contextually and return structured JSON.
  - This helps bifurcate data into fields like `name`, `address`, `invoice_no`, etc.

- **Installation Instructions:**
  1. Download Ollama for Windows: [Click-Here](https://ollama.com/download/windows)  
  2. Install Ollama by running the installer.  
  3. Open Command Prompt and pull the Granite model:  
     ```bash
     ollama pull granite3.2-vision:2b
     ```
  4. Verify installation by running:  
     ```bash
     ollama run granite3.2-vision:2b
     ```

- **YouTube Tutorial (for reference):**  
  How to Install Ollama on Windows [(Click-Here)](https://www.bing.com/search?q="https%3A%2F%2Fwww.youtube.com%2Fwatch%3Fv%3DQxFJ5zZpV9o")  

- **Code (Cell 10):**

```python
# LET'S BRING GRANITE MODEL TO MAKE SURE IT WORKS
MODEL_NAME = "granite3.2-vision:2b"   # TEXT model (FASTER)
print("Granite text model is ready")
```


##  Cell 11 – Regex Patterns & Contact Extraction
This cell defines regex patterns to extract mobile numbers and emails from OCR text (which is in business_card variable).

- **What it does:**
  - Defines regex for mobile numbers and email addresses and more Variables as per your requirements.
  - instead of blank we will fill "-".

- **Why it matters:**
  - Regex is the first step in structuring raw OCR text.
  - Ensures phone numbers and emails are captured Properly, it Depends on regex patterns.

- **Where to modify:**
  - Update regex patterns if your data requires different formats (e.g., GST numbers, invoice numbers).
  - Add new regex rules for custom fields.

```python
# ----------------------------
# REGEX PATTERNS
# ----------------------------

mobile_pattern = r'(\+?\d[\d\s\-]{8,}\d)'
email_pattern = r'[A-Za-z0-9._%+-]+@[A-Za-z0-9.-]+\.[A-Za-z]{2,}'

def clean_mobile(number):
    digits = re.sub(r'\D', '', number)
    if len(digits) >= 10:
        return digits
    return None

def extract_contact_info(text):
    
    # ---- EMAIL ----
    emails = re.findall(email_pattern, text)
    email = emails[0] if emails else "-"
    
    # ---- MOBILE ----
    raw_numbers = re.findall(mobile_pattern, text)
    
    cleaned_numbers = []
    for num in raw_numbers:
        cleaned = clean_mobile(num)
        if cleaned and cleaned not in cleaned_numbers:
            cleaned_numbers.append(cleaned)
    
    primary = cleaned_numbers[0] if len(cleaned_numbers) > 0 else "-"
    secondary = cleaned_numbers[1] if len(cleaned_numbers) > 1 else "-"
    
    return primary, secondary, email, cleaned_numbers

print("CELL 1: REGEX EXTRACTION READY - THIS TASK IS FINISH PROPERLY")
```


**Cell 12 removes extracted contact info (numbers and emails) from the OCR text, This is an optional step based on your requirement.**


##  Cell 13 – Granite Model Extraction
This cell defines how Granite (via Ollama) extracts structured info from cleaned text.

- **What it does:**
  - give customise prompt to Granite model. 
  - Returns JSON output.

- **Why it matters:**
  - Granite interprets text contextually, unlike regex.
  - This step bifurcates unstructured text into structured fields.

- **Where to modify:**
  - Update the prompt to include fields you need (e.g., `invoice_no`, `gst_no`, `date`).
  - Accuracy depends heavily on prompt design incl JSON output.

```python
def extract_from_granite(cleaned_text):
    
    prompt = f"""
You are extracting information from a business card.

STRICT RULES:
1) Do NOT guess.
2) If unsure return "-".
3) Extract only:
   - name (person name only)
   - address (physical location only)
   - additional_info (designation, company, website etc.)

Return ONLY raw JSON in this format:

{{
"name": "",
"address": "",
"additional_info": ""
}}

TEXT:
{cleaned_text}
"""

    try:
        response = ollama.chat(
            model=MODEL_NAME,
            messages=[{"role": "user", "content": prompt}]
        )

        content = response["message"]["content"]
        content = content.replace("```json", "").replace("```", "").strip()

        return json.loads(content)

    except:
        return {
            "name": "-",
            "address": "-",
            "additional_info": "-"
        }

print("CELL 3: GRANITE EXTRACTION READY - THIS TASK IS FINISH PROPERLY")
```
**final_rows(extracted info by Model) Convert into dataframe and Cell 18 helps you to save the output into .xlsx on your given path.**

- **Example Output:**

```text
<class 'pandas.core.frame.DataFrame'>
RangeIndex: 89 entries, 0 to 88
Data columns (total 10 columns):
 #   Column               Non-Null Count  Dtype 
---  ------               --------------  ----- 
 0   card                 89 non-null     object
 1   photo_name           89 non-null     object
 2   IMG_Link             89 non-null     object
 3   name                 89 non-null     object
 4   mobile_no            89 non-null     object
 5   secondary_mobile_no  89 non-null     object
 6   email                89 non-null     object
 7   address              89 non-null     object
 8   additional_info      89 non-null     object
 9   all_text_info        89 non-null     object
dtypes: object(10)
```
