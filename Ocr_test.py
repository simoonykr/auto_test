import os
import pyautogui
from PIL import Image, ImageEnhance, ImageOps, ImageFilter
import pytesseract
import pandas as pd
import openpyxl

# Tesseract environment variable setup
os.environ['TESSDATA_PREFIX'] = r'C:\Program Files\Tesseract-OCR\tessdata'

# Explicit path for Tesseract
pytesseract.pytesseract.tesseract_cmd = r'C:\Program Files\Tesseract-OCR\tesseract.exe'

# Capture region coordinates adjustment (resolution adjustment)
region = (0, 0, 798, 645)  # Adjust to fit the entire window screen

# Capture the entire screen.
screenshot = pyautogui.screenshot(region=region)

# Save the captured image to a file.
screenshot_path = "screenshot.png"
screenshot.save(screenshot_path)

# Reopen the saved screenshot for preprocessing and OCR.
screenshot_image = Image.open(screenshot_path)

# Image preprocessing steps
# Convert the image to grayscale
gray_image = ImageOps.grayscale(screenshot_image)

# Increase resolution (scale up)
scale_factor = 2
width, height = gray_image.size
resized_image = gray_image.resize((width * scale_factor, height * scale_factor), Image.LANCZOS)

# Adjust contrast
enhancer = ImageEnhance.Contrast(resized_image)
enhanced_image = enhancer.enhance(2)  # Increase contrast

# Noise reduction (applying filters)
filtered_image = enhanced_image.filter(ImageFilter.MedianFilter())

# Extract text from preprocessed image
text = pytesseract.image_to_string(filtered_image, lang='kor')  # Change to Korean language code "kor"

# Save extracted text to a text file
text_file_path = "extracted_text.txt"
with open(text_file_path, "w", encoding="utf-8") as text_file:
    text_file.write(text)

print(f"Extracted Text saved to: {text_file_path}")

# Enter data into Excel file
excel_file = "C:\\Users\\simoony\\auto\\auto.xlsx"
data = {'Extracted_Text': [text]}
df = pd.DataFrame(data)

# Check if 'Sheet1' exists and remove it
wb = openpyxl.load_workbook(excel_file)
if 'Sheet1' in wb.sheetnames:
    std = wb['Sheet1']
    wb.remove(std)

# Add DataFrame to Excel file
with pd.ExcelWriter(excel_file, mode='w', engine='openpyxl') as writer:
    df.to_excel(writer, index=False, sheet_name='Sheet1')
