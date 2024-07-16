import os
import pyautogui
from PIL import Image
import pytesseract

# Tesseract 환경 변수 설정
os.environ['TESSDATA_PREFIX'] = r'C:\Program Files\Tesseract-OCR\tessdata'

# Tesseract 명시적 경로 설정
pytesseract.pytesseract.tesseract_cmd = r'C:\Program Files\Tesseract-OCR\tesseract.exe'

# 캡처할 영역의 좌표를 설정
region = (0, 0, 798, 645)  # 해상도에 맞게 조정

# 전체 화면을 캡처
screenshot = pyautogui.screenshot(region=region)

# 캡처한 이미지를 파일로 저장합니다.
screenshot_path = "screenshot1.png"
screenshot.save(screenshot_path)

# OCR을 통해 한글 텍스트를 추출
text_kor = pytesseract.image_to_string(screenshot, lang='kor')

# OCR을 통해 영어 텍스트를 추출
text_eng = pytesseract.image_to_string(screenshot, lang='eng')

# 한글 텍스트를 파일로 저장 (UTF-8 인코딩)
with open("extracted_text_kor.txt", "w", encoding="utf-8") as file_kor:
    file_kor.write(text_kor)

# 영어 텍스트를 파일로 저장 (UTF-8 인코딩)
with open("extracted_text_eng.txt", "w", encoding="utf-8") as file_eng:
    file_eng.write(text_eng)

print("Korean Text saved to: extracted_text_kor.txt")
print("English Text saved to: extracted_text_eng.txt")
