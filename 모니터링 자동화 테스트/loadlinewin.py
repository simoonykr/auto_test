import pyautogui
import cv2
import numpy as np
import time
from openpyxl import load_workbook
import os
import tkinter as tk
from tkinter import messagebox

# Excel 파일 경로 설정
excel_file_path = "C:\\Users\\simoony\\auto\\ROLDLINE\\loadaotu.xlsx"
wb = load_workbook(excel_file_path)
sheet = wb.active

# 검색을 위한 최대 재시도 횟수
max_retries = 5

# 스케일링 단계
scales = [1.0, 0.9, 1.1]

def find_image_on_screen(template_path, scales):
    screenshot = pyautogui.screenshot()
    screenshot = cv2.cvtColor(np.array(screenshot), cv2.COLOR_RGB2BGR)
    screenshot_gray = cv2.cvtColor(screenshot, cv2.COLOR_BGR2GRAY)

    template = cv2.imread(template_path, cv2.IMREAD_UNCHANGED)
    
    if template is None:
        print(f"파일을 열 수 없습니다: {template_path}")
        return None
    
    if template.shape[-1] == 4:
        template = cv2.cvtColor(template, cv2.COLOR_BGRA2BGR)
    
    template_gray = cv2.cvtColor(template, cv2.COLOR_BGR2GRAY)

    for scale in scales:
        width = int(template_gray.shape[1] * scale)
        height = int(template_gray.shape[0] * scale)
        scaled_template = cv2.resize(template_gray, (width, height), interpolation=cv2.INTER_AREA)

        res = cv2.matchTemplate(screenshot_gray, scaled_template, cv2.TM_CCOEFF_NORMED)
        min_val, max_val, min_loc, max_loc = cv2.minMaxLoc(res)
        print(f"Scale: {scale}, Max Val: {max_val}")

        threshold = 0.8
        if max_val >= threshold:
            return max_loc, width, height

    return None

def start_search(interval):
    try:
        interval = int(interval) * 60  # Convert minutes to seconds
        while True:
            found_first_image = False
            for row in range(1, 10):
                cell_value = sheet[f'A{row}'].value
                print(f"Excel 파일에서 읽은 데이터 (A{row}): {cell_value}")

                if not cell_value:
                    continue
                
                for attempt in range(max_retries):
                    try:
                        print(f"A{row}의 이미지를 찾는 중... 시도 {attempt + 1}")

                        result = find_image_on_screen(cell_value, scales)
                        
                        if result:
                            max_loc, width, height = result

                            print(f"이미지를 찾았습니다: 위치 {max_loc}, 크기 ({width}, {height})")

                            center_x = max_loc[0] + width // 2
                            center_y = max_loc[1] + height // 2

                            pyautogui.moveTo(center_x, center_y)
                            pyautogui.click()
                            print("이미지를 클릭했습니다.")

                            sheet[f'B{row}'] = "클릭 성공"
                            wb.save(excel_file_path)
                            print(f"엑셀 파일 업데이트 성공: 셀 B{row}")

                            if "move" in cell_value:
                                print(f"A{row}에 'move'가 있어서 마우스 스크롤 다운을 실행합니다.")
                                pyautogui.scroll(-500)

                            time.sleep(2)

                            if row == 1:
                                found_first_image = True
                            break
                        else:
                            print(f"A{row}의 이미지를 찾지 못했습니다. 재시도합니다.")
                            time.sleep(5)
                    except Exception as e:
                        print(f"A{row}의 이미지를 찾는 중 오류 발생: {str(e)}")
            if found_first_image:
                print("첫 번째 이미지를 성공적으로 찾았습니다. 대기 시간 후 다시 시작합니다.")
            else:
                print("첫 번째 이미지를 찾지 못했습니다. 대기 시간 후 다시 시도합니다.")
            time.sleep(interval)
    finally:
        wb.close()

# GUI 설정
def run_gui():
    root = tk.Tk()
    root.title("이미지 검색 프로그램")

    tk.Label(root, text="검색 주기 (분 단위):").grid(row=0, column=0)

    interval_entry = tk.Entry(root)
    interval_entry.grid(row=0, column=1)

    def on_start():
        try:
            interval = interval_entry.get()
            if interval.isdigit() and int(interval) > 0:
                messagebox.showinfo("프로그램 시작", "이미지 검색을 시작합니다.")
                root.destroy()  # GUI 창 닫기
                start_search(interval)
            else:
                messagebox.showerror("입력 오류", "양의 정수 값을 입력하세요.")
        except Exception as e:
            messagebox.showerror("오류", f"예기치 않은 오류가 발생했습니다: {str(e)}")

    start_button = tk.Button(root, text="시작", command=on_start)
    start_button.grid(row=1, column=0, columnspan=2)

    root.mainloop()

# GUI 실행
run_gui()

