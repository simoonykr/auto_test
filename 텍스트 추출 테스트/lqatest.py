import pyautogui
import time
from pynput import mouse
from PIL import ImageGrab, Image, ImageOps, ImageEnhance, ImageFilter
import os
from openpyxl import Workbook
from openpyxl.drawing.image import Image as ExcelImage
import pytesseract

# Tesseract 경로 설정 (이 경로는 실제 Tesseract OCR이 설치된 경로로 변경해야 합니다)
pytesseract.pytesseract.tesseract_cmd = r'C:\Program Files\Tesseract-OCR\tesseract.exe'

# 전역 변수 설정
start_x = start_y = end_x = end_y = 0
dragging = False
scale_factor = 2  # 해상도 증가 배율

# 마우스를 사용한 클릭 콜백 함수
def on_click(x, y, button, pressed):
    global start_x, start_y, end_x, end_y, dragging
    if pressed:
        start_x, start_y = x, y
        dragging = True
    else:
        end_x, end_y = x, y
        dragging = False
        return False

# 캡처 영역 선택 함수
def get_selection_area():
    with mouse.Listener(on_click=on_click) as listener:
        print("드래그하여 캡처 영역을 지정하세요. 시작하려면 마우스를 클릭하고, 끝내려면 마우스를 놓으세요.")
        listener.join()
    left = min(start_x, end_x)
    right = max(start_x, end_x)
    top = min(start_y, end_y)
    bottom = max(start_y, end_y)
    print(f"지정된 영역: ({left}, {top})에서 ({right}, {bottom})")
    return (left, top, right, bottom)

# 영역 캡처 및 저장 함수
def capture_and_save_area(area, filename):
    try:
        screenshot = ImageGrab.grab(bbox=area)
        screenshot.save(filename)
        print(f"영역이 {filename}로 저장되었습니다.")
    except Exception as e:
        print(f"Error: {e}")

# 이미지 전처리 함수
def preprocess_image(image_path):
    screenshot_image = Image.open(image_path)
    gray_image = ImageOps.grayscale(screenshot_image)
    width, height = gray_image.size
    resized_image = gray_image.resize((width * scale_factor, height * scale_factor), Image.LANCZOS)
    enhancer = ImageEnhance.Contrast(resized_image)
    enhanced_image = enhancer.enhance(2)
    sharpened_image = enhanced_image.filter(ImageFilter.SHARPEN)
    final_image = sharpened_image.filter(ImageFilter.EDGE_ENHANCE_MORE)
    preprocessed_image_path = image_path.replace(".png", "_processed.png")
    final_image.save(preprocessed_image_path)
    return final_image, preprocessed_image_path

# OCR로 이미지에서 텍스트 추출 함수
def ocr_image(image_path):
    try:
        preprocessed_image, preprocessed_image_path = preprocess_image(image_path)
        text = pytesseract.image_to_string(preprocessed_image, lang='kor+eng')  # 한글과 영어 동시 인식
        return text, preprocessed_image_path
    except Exception as e:
        print(f"Error in OCR processing: {e}")
        return "", ""

# 이미지 목록을 엑셀에 저장하는 함수
def save_images_to_excel(images_list, excel_path="Image_save.xlsx"):
    wb = Workbook()
    ws = wb.active
    
    for index, image_path in enumerate(images_list, start=1):
        if os.path.exists(image_path):
            img = ExcelImage(image_path)
            img.anchor = f'A{index*10}'
            ws.add_image(img)
            extracted_text, preprocessed_image_path = ocr_image(image_path)
            print(f"Extracted Text from Image {index}: {extracted_text.strip()}")
            ws[f'B{index*10}'] = extracted_text
            if os.path.exists(preprocessed_image_path):
                processed_img = ExcelImage(preprocessed_image_path)
                processed_img.anchor = f'C{index*10}'
                ws.add_image(processed_img)
        else:
            print(f"이미지 파일을 찾을 수 없습니다: {image_path}")

    wb.save(excel_path)
    print(f"이미지가 {excel_path} 파일에 저장되었습니다.")

if __name__ == "__main__":
    images_list = []

    for i in range(1, 5):
        print(f"{i}번째 이미지를 선택하세요.")
        area = get_selection_area()
        image_path = f"temp_screenshot{i}.png"
        capture_and_save_area(area, image_path)
        images_list.append(image_path)

    save_images_to_excel(images_list)
