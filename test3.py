import os
import pyautogui
from PIL import Image, ImageEnhance, ImageOps, ImageFilter
import pytesseract
import openpyxl
import time

# Tesseract 환경 변수 설정
os.environ['TESSDATA_PREFIX'] = r'C:\Program Files\Tesseract-OCR\tessdata'

# Tesseract 명시적 경로 설정
pytesseract.pytesseract.tesseract_cmd = r'C:\Program Files\Tesseract-OCR\tesseract.exe'

# 캡처할 영역의 좌표를 조정 (해상도 조정)
region = (0, 0, 798, 1200)  # 윈도우 화면 전체 크기에 맞게 조정

# 캡처할 스크린샷 개수
num_screenshots = 5  # 더 많거나 적게 캡처하려면 이 값을 조정하세요

try:
    # Excel 파일에서 특정 단어 읽기
    excel_file = "C:\\Users\\simoony\\auto\\auto.xlsx"
    wb = openpyxl.load_workbook(excel_file)
    print(f"Opened Excel file: {excel_file}")

    sheet = wb['Sheet1']
    print("Accessing 'Sheet1'.")

    targets = []
    for box_idx in range(1, 6):  # 'A1'부터 'A5'까지 순회
        target_word = sheet[f'A{box_idx}'].value
        if target_word is not None:
            targets.append((box_idx, target_word))
            print(f"Target word from Excel (A{box_idx}): {target_word}")

except FileNotFoundError:
    print(f"File not found: {excel_file}")
    exit()
except KeyError:
    print("'Sheet1' 시트가 없습니다. Excel 파일을 확인하십시오.")
    exit()
except Exception as e:
    print(f"An error occurred: {e}")
    exit()

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

    # 전처리된 이미지에서 텍스트와 위치 추출
    ocr_data = pytesseract.image_to_data(filtered_image, lang='kor', output_type=pytesseract.Output.DICT)
    num_boxes = len(ocr_data['level'])

    # OCR 텍스트를 가로로 정렬하여 저장
    targets_to_remove = []
    for target in targets[:]:  # targets 리스트를 복사하여 순회
        box_idx, target_word = target
        found = False
        for j in range(num_boxes):
            if ocr_data['text'][j] == target_word:
                x = int(ocr_data['left'][j] / scale_factor)
                y = int(ocr_data['top'][j] / scale_factor)
                width = int(ocr_data['width'][j] / scale_factor)
                height = int(ocr_data['height'][j] / scale_factor)

                # 출력
                print(f"Found target word '{target_word}' at position ({x}, {y})")

                # 마우스를 해당 위치로 이동하여 클릭
                pyautogui.click(region[0] + x + width // 2, region[1] + y + height // 2)
                print(f"Clicked on {target_word} at: ({region[0] + x + width // 2}, {region[1] + y + height // 2})")

                # 매칭 완료 기록
                sheet[f'B{box_idx}'] = "매칭완료"
                print(f"Recorded matching completed for '{target_word}' in cell B{box_idx}")

                found = True
                targets_to_remove.append(target)
                break

        if found:
            # 목표 단어를 찾았으므로 targets 리스트에서 제거
            targets.remove(target)

        # OCR 텍스트를 가로로 정렬하여 저장
        text_data = [ocr_data['text'][idx] for idx in sorted(range(num_boxes), key=lambda x: ocr_data['left'][x])]

        # 텍스트 한 줄에 이어서 저장
        text_file_path_kor = f"extracted_text_kor_{box_idx}.txt"
        with open(text_file_path_kor, "w", encoding="utf-8") as text_file_kor:
            text_file_kor.write(' '.join(text_data))  # 한 줄에 이어서 저장
        print(f"Korean Text saved to: {text_file_path_kor}")

    # 매칭 상태를 엑셀 파일에 저장
    wb.save(excel_file)

    # 다음 스크린샷을 찍기 전에 10초 대기
    time.sleep(10)
