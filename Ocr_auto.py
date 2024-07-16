import os
import pyautogui
from PIL import Image, ImageEnhance, ImageOps, ImageFilter
import pytesseract
import pandas as pd
import time

# Tesseract 환경 변수 설정
os.environ['TESSDATA_PREFIX'] = r'C:\Program Files\Tesseract-OCR\tessdata'

# Tesseract 명시적 경로 설정
pytesseract.pytesseract.tesseract_cmd = r'C:\Program Files\Tesseract-OCR\tesseract.exe'

# 캡처할 영역의 좌표를 조정 (해상도 조정)
region = (0, 0, 798, 645)  # 윈도우 화면 전체 크기에 맞게 조정

# 캡처할 스크린샷 개수
num_screenshots = 5  # 더 많거나 적게 캡처하려면 이 값을 조정하세요

for i in range(num_screenshots):
    # 전체 화면을 캡처합니다.
    screenshot = pyautogui.screenshot(region=region)

    # 캡처한 이미지를 파일로 저장합니다.
    screenshot_path = f"screenshot_{i}.png"
    screenshot.save(screenshot_path)

    # 저장된 스크린샷을 다시 열어 전처리 후 OCR을 수행합니다.
    screenshot_image = Image.open(screenshot_path)

    # 이미지 전처리 과정 추가
    # 이미지를 흑백으로 변환
    gray_image = ImageOps.grayscale(screenshot_image)

    # 해상도 증가 (배율 확대)
    scale_factor = 2
    width, height = gray_image.size
    resized_image = gray_image.resize((width * scale_factor, height * scale_factor), Image.LANCZOS)

    # 명암비 (contrast) 조정
    enhancer = ImageEnhance.Contrast(resized_image)
    enhanced_image = enhancer.enhance(2)  # 명암비를 증가시킴

    # 노이즈 제거 (필터 적용)
    filtered_image = enhanced_image.filter(ImageFilter.MedianFilter())

    # 전처리된 이미지를 저장 (필터 적용된 이미지)
    preprocessed_image_path = f"preprocessed_screenshot_{i}.png"
    filtered_image.save(preprocessed_image_path)
    print(f"Preprocessed Image saved to: {preprocessed_image_path}")

    # 전처리된 이미지에서 텍스트 추출
    text = pytesseract.image_to_string(filtered_image, lang='kor')  # 한글 언어 코드 "kor"로 변경

    # 텍스트 파일로 저장
    text_file_path = f"extracted_text_{i}.txt"
    with open(text_file_path, "w", encoding="utf-8") as text_file:
        text_file.write(text)

    print(f"Extracted Text saved to: {text_file_path}")

    # 다음 스크린샷을 찍기 전에 5초 대기
    time.sleep(5)
