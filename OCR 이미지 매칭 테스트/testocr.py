import os
import pyautogui
from PIL import Image, ImageEnhance, ImageOps, ImageFilter
import pytesseract
import openpyxl
import time
import re
from pynput import mouse

# Tesseract 환경 변수 설정
os.environ['TESSDATA_PREFIX'] = r'C:\\Program Files\\Tesseract-OCR\\tessdata'

# Tesseract 명시적 경로 설정
pytesseract.pytesseract.tesseract_cmd = r'C:\\Program Files\\Tesseract-OCR\\tesseract.exe'

# 캡처할 영역의 좌표를 마우스 드래그로 지정
start_x = start_y = end_x = end_y = 0
dragging = False

def on_click(x, y, button, pressed):
    global start_x, start_y, end_x, end_y, dragging
    if pressed:
        start_x, start_y = x, y
        dragging = True
    else:
        end_x, end_y = x, y
        dragging = False
        # Stop listener
        return False

# Listen to mouse events
with mouse.Listener(on_click=on_click) as listener:
    print("드래그하여 캡처 영역을 지정하세요. 시작하려면 마우스를 클릭하고, 끝내려면 마우스를 놓으세요.")
    listener.join()

region = (min(start_x, end_x), min(start_y, end_y), abs(end_x - start_x), abs(end_y - start_y))
print(f"지정된 캡처 영역: {region}")

# 캡처할 스크린샷 개수
num_screenshots = 50  # 더 많거나 적게 캡처하려면 이 값을 조정하세요

# 이미지 확대 배율
scale_factor = 2

def preprocess_image(image_path):
    # 저장된 스크린샷을 다시 열어 전처리 후 OCR을 수행합니다.
    screenshot_image = Image.open(image_path)

    # 이미지 전처리 과정 추가
    # 이미지를 흑백으로 변환
    gray_image = ImageOps.grayscale(screenshot_image)

    # 해상도 증가 (배율 확대)
    global scale_factor  # 전역 변수를 사용하여 scale_factor를 참조합니다.
    width, height = gray_image.size
    resized_image = gray_image.resize((width * scale_factor, height * scale_factor), Image.LANCZOS)

    # 명암비 (contrast) 조정
    enhancer = ImageEnhance.Contrast(resized_image)
    enhanced_image = enhancer.enhance(2)  # 명암비를 증가시킴

    # 샤프닝 필터와 추가적인 필터 적용
    sharpened_image = enhanced_image.filter(ImageFilter.SHARPEN)
    enhanced_image = sharpened_image.filter(ImageFilter.EDGE_ENHANCE_MORE)

    # 전처리된 이미지를 저장 (필터 적용된 이미지)
    preprocessed_image_path = image_path.replace("screenshot", "preprocessed_screenshot")
    enhanced_image.save(preprocessed_image_path)
    print(f"Preprocessed Image saved to: {preprocessed_image_path}")

    return preprocessed_image_path

def load_excel_targets(excel_file):
    try:
        wb = openpyxl.load_workbook(excel_file)
        sheet = wb['Sheet1']
        targets = []
        for box_idx in range(1, 66):
            target_word = sheet[f'A{box_idx}'].value
            if target_word is not None:
                targets.append((box_idx, target_word))
                print(f"Target word from Excel (A{box_idx}): {target_word}")
        return wb, sheet, targets
    except FileNotFoundError:
        print(f"File not found: {excel_file}")
        exit()
    except KeyError:
        print("'Sheet1' 시트가 없습니다. Excel 파일을 확인하십시오.")
        exit()
    except Exception as e:
        print(f"An error occurred: {e}")
        exit()

def is_nearby_position(pos1, pos2, tolerance=5):
    return all(abs(p1 - p2) <= tolerance for p1, p2 in zip(pos1, pos2))

# 매칭 완료 카운터
match_counter = 1

# 엑셀 데이터 로드
excel_file = "C:\\Users\\simoony\\auto\\auto.xlsx"
wb, sheet, targets = load_excel_targets(excel_file)

# 이미 매칭된 좌표를 추적
matched_positions = []

for i in range(num_screenshots):
    # 전체 화면을 캡처합니다.
    screenshot = pyautogui.screenshot(region=region)

    # 캡처한 이미지를 파일로 저장합니다.
    screenshot_path = f"screenshot_{i}.png"
    screenshot.save(screenshot_path)

    # 이미지 전처리
    preprocessed_image_path = preprocess_image(screenshot_path)

    # Tesseract 옵션 설정
    custom_config = r'--oem 3 --psm 6'  # oem: OCR Engine Mode, psm: Page Segmentation Mode

    # 전처리된 이미지에서 텍스트와 위치 추출
    ocr_data = pytesseract.image_to_data(preprocessed_image_path, lang='kor', config=custom_config, output_type=pytesseract.Output.DICT)
    num_boxes = len(ocr_data['level'])

    # OCR 텍스트를 가로로 정렬하여 저장
    for target in targets[:]:  # targets 리스트를 복사하여 순회
        box_idx, target_word = target
        found = False
        for j in range(num_boxes):
            extracted_text = ocr_data['text'][j]
            # 한글만 필터링
            extracted_text = re.sub('[^ㄱ-ㅎㅏ-ㅣ가-힣]', '', extracted_text)

            if extracted_text == target_word:
                x = int(ocr_data['left'][j] / scale_factor)
                y = int(ocr_data['top'][j] / scale_factor)
                width = int(ocr_data['width'][j] / scale_factor)
                height = int(ocr_data['height'][j] / scale_factor)

                position = (x, y, width, height)
                if any(is_nearby_position(position, pos) for pos in matched_positions):
                    continue  # 이미 매칭된 위치는 건너뜀

                # 출력 및 매칭된 위치 저장
                print(f"Found target word '{target_word}' at position ({x}, {y})")
                matched_positions.append(position)

                # 마우스를 해당 위치로 이동하여 클릭
                click_x = region[0] + x + width // 2
                click_y = region[1] + y + height // 2
                pyautogui.click(click_x, click_y)
                print(f"Clicked on {target_word} at: ({click_x}, {click_y})")

                # 매칭 완료 기록
                sheet[f'B{box_idx}'] = f"매칭완료{match_counter}"
                print(f"Recorded matching completed for '{target_word}' in cell B{box_idx}")

                found = True
                targets.remove(target)
                match_counter += 1
                break

        if found:
            # 매칭 완료 후 5초 대기
            time.sleep(5)

            # OCR 텍스트를 가로로 정렬하여 저장
            text_data = [ocr_data['text'][idx] for idx in sorted(range(num_boxes), key=lambda x: ocr_data['left'][x])]
            korean_text_data = [re.sub('[^ㄱ-ㅎㅏ-ㅣ가-힣]', '', text) for text in text_data]
            korean_text_data = ' '.join([text for text in korean_text_data if text])

            # 텍스트 한 줄에 이어서 저장
            text_file_path_kor = f"extracted_text_kor_{box_idx}.txt"
            with open(text_file_path_kor, "w", encoding="utf-8") as text_file_kor:
                text_file_kor.write(korean_text_data)  # 한 줄에 이어서 저장
            print(f"Korean Text saved to: {text_file_path_kor}")

    # 매칭 상태를 엑셀 파일에 저장하고 닫음
    wb.save(excel_file)
    wb.close()

    # 더 이상 타겟 단어가 없으면 종료
    if not targets:
        break

    # 엑셀 파일을 다시 열어 다음 매칭 준비
    wb, sheet, targets = load_excel_targets(excel_file)

    # 다음 스크린샷을 찍기 전에 10초 대기
    time.sleep(5)

# 최종적으로 엑셀 파일을 저장하고 닫음
wb.save(excel_file)
wb.close()
