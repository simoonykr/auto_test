import pyautogui
import cv2
import numpy as np
import time
from openpyxl import load_workbook
import os

# Excel 파일 경로 설정
excel_file_path = "C:\\Users\\simoony\\auto\\ROLDLINE\\loadaotu.xlsx"
wb = load_workbook(excel_file_path)
sheet = wb.active

# 검색을 위한 최대 재시도 횟수
max_retries = 3

# 스크린샷을 저장할 폴더 설정 (필요시 경로 수정)
screenshot_folder = "C:\\Users\\simoony\\auto\\LOH\\screenshots"
os.makedirs(screenshot_folder, exist_ok=True)

# 캡처할 영역의 좌표 (해상도 및 필요에 맞게 조정)
region = (0, 0, 798, 645)  # 고정된 영역 설정

# 스케일링 단계
scales = [1.0, 0.9, 1.1]

def find_image_on_screen(template_path, region, scales):
    # 스크린샷을 지정된 영역에서만 찍기
    screenshot = pyautogui.screenshot(region=region)
    screenshot = cv2.cvtColor(np.array(screenshot), cv2.COLOR_RGB2BGR)
    screenshot_gray = cv2.cvtColor(screenshot, cv2.COLOR_BGR2GRAY)

    template = cv2.imread(template_path, cv2.IMREAD_UNCHANGED)
    
    if template is None:
        print(f"파일을 열 수 없습니다: {template_path}")
        return None
    
    if template.shape[-1] == 4:
        template = cv2.cvtColor(template, cv2.COLOR_BGRA2BGR)
    
    template_gray = cv2.cvtColor(template, cv2.COLOR_BGR2GRAY)

    # 여러 스케일에서 템플릿 매칭 시도
    for scale in scales:
        width = int(template_gray.shape[1] * scale)
        height = int(template_gray.shape[0] * scale)
        scaled_template = cv2.resize(template_gray, (width, height), interpolation=cv2.INTER_AREA)

        res = cv2.matchTemplate(screenshot_gray, scaled_template, cv2.TM_CCOEFF_NORMED)
        min_val, max_val, min_loc, max_loc = cv2.minMaxLoc(res)
        print(f"Scale: {scale}, Max Val: {max_val}")

        # 매칭 임계값 설정
        threshold = 0.8
        if max_val >= threshold:
            return max_loc, width, height

    return None

try:
    # A1부터 A4까지 데이터 경로를 순차적으로 처리
    for row in range(1, 10):  # A1 ~ A4
        # 셀에서 데이터 읽기
        cell_value = sheet[f'A{row}'].value
        print(f"Excel 파일에서 읽은 데이터 (A{row}): {cell_value}")

        # 이미지 경로 존재 여부 확인
        if not cell_value:
            continue  # 경로가 없으면 다음 줄로 이동

        # 스크린샷 찍기 (지정된 영역에서만)
        screenshot_path = os.path.join(screenshot_folder, f'screenshot_A{row}.png')
        pyautogui.screenshot(screenshot_path, region=region)
        print(f"스크린샷이 저장되었습니다: {screenshot_path}")

        # 이미지 검색 및 클릭 시도
        for attempt in range(max_retries):
            try:
                print(f"A{row}의 이미지를 찾는 중... 시도 {attempt + 1}")

                result = find_image_on_screen(cell_value, region, scales)
                
                if result:
                    max_loc, width, height = result

                    print(f"이미지를 찾았습니다: 위치 {max_loc}, 크기 ({width}, {height})")

                    # 이미지의 중심 좌표 계산
                    center_x = max_loc[0] + width // 2
                    center_y = max_loc[1] + height // 2

                    # 마우스 이동 및 클릭
                    pyautogui.moveTo(region[0] + center_x, region[1] + center_y)
                    pyautogui.click()
                    print("이미지를 클릭했습니다.")

                    # Excel 파일 업데이트
                    sheet[f'B{row}'] = "클릭 성공"
                    wb.save(excel_file_path)
                    print(f"엑셀 파일 업데이트 성공: 셀 B{row}")

                    # 클릭 후 "move" 단어가 있는 경우 스크롤 다운
                    if "move" in cell_value:
                        print(f"A{row}에 'move'가 있어서 마우스 스크롤 다운을 실행합니다.")
                        pyautogui.scroll(-500)  # 스크롤 다운 (-는 아래로 스크롤)

                    # 클릭 및 스크롤 후 1초 대기
                    time.sleep(2)
                    break
                else:
                    print(f"A{row}의 이미지를 찾지 못했습니다. 재시도합니다.")
                    time.sleep(1)  # 이미지 탐색 실패 후 잠시 대기
            except Exception as e:
                print(f"A{row}의 이미지를 찾는 중 오류 발생: {str(e)}")
finally:
    wb.close()
