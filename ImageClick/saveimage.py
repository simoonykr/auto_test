import pyautogui
import time
from pynput import mouse
from PIL import ImageGrab, Image
import os
from openpyxl import Workbook
from openpyxl.drawing.image import Image as ExcelImage

# 전역 변수 설정
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
def get_selection_area():
    with mouse.Listener(on_click=on_click) as listener:
        print("드래그하여 캡처 영역을 지정하세요. 시작하려면 마우스를 클릭하고, 끝내려면 마우스를 놓으세요.")
        listener.join()
    print(f"지정된 영역: ({start_x}, {start_y})에서 ({end_x}, {end_y})")
    return (start_x, start_y, end_x, end_y)

# 이미지를 지정된 영역에서 검색 (타임아웃 추가)
def search_image_in_area(image_path, area, timeout=10):
    start_time = time.time()
    while time.time() - start_time < timeout:
        try:
            location = pyautogui.locateCenterOnScreen(image_path, region=area, confidence=0.8)
            if location:
                return location
        except Exception as e:
            print(f"Error: {e}")
        time.sleep(1)  # 1초 대기 후 다시 시도
    return None

def capture_and_save_area(area, filename):
    try:
        # 영역을 캡처한 후 저장합니다.
        screenshot = ImageGrab.grab(bbox=area)
        screenshot.save(filename)
        print(f"영역이 {filename}로 저장되었습니다.")
    except Exception as e:
        print(f"Error: {e}")

def save_images_to_excel(images_list, excel_path="Image_save.xlsx"):
    wb = Workbook()
    ws = wb.active
    
    for index, image_path in enumerate(images_list, start=1):
        if os.path.exists(image_path):  # 이미지 파일이 존재하는지 확인
            img = ExcelImage(image_path)
            img.anchor = f'A{index*10}'  # 이미지를 각 행에 맞추기 위해 *10을 사용
            ws.add_image(img)
        else:
            print(f"이미지 파일을 찾을 수 없습니다: {image_path}")

    wb.save(excel_path)
    print(f"이미지가 {excel_path} 파일에 저장되었습니다.")

def click_image(image_path, area, click_count, image_index):
    location = search_image_in_area(image_path, area)
    if location:
        print(f"{image_path}의 위치를 {location}에서 찾았습니다.")
        time.sleep(1)  # 클릭하기 전에 1초 대기
        for _ in range(click_count):
            pyautogui.click(location)
            time.sleep(1)
        print(f"{image_index} 번째 이미지 클릭 완료")
    else:
        print(f"{image_path} 이미지를 지정된 영역에서 찾을 수 없습니다. 다음 이미지로 넘어갑니다.")

def main():
    # 전체 프로세스를 몇 번 실행할지 사용자로부터 입력받음
    repeat_count = int(input("프로세스를 몇 번 실행할까요? "))

    # 캡처한 이미지 파일 리스트
    images_list = []

    # 첫 번째 이미지 영역 선택
    print("첫 번째 이미지를 선택하세요.")
    area1 = get_selection_area()
    image_path1 = "temp_screenshot1.png"
    capture_and_save_area(area1, image_path1)
    images_list.append(image_path1)

    # 첫 번째 이미지 클릭 횟수를 사용자로부터 입력받음
    click_count1 = int(input("첫 번째 이미지를 몇 번 클릭하시겠습니까? "))

    # 두 번째 이미지 영역 선택
    print("두 번째 이미지를 선택하세요.")
    area2 = get_selection_area()
    image_path2 = "temp_screenshot2.png"
    capture_and_save_area(area2, image_path2)
    images_list.append(image_path2)

    # 두 번째 이미지 클릭 횟수를 사용자로부터 입력받음
    click_count2 = int(input("두 번째 이미지를 몇 번 클릭하시겠습니까? "))

    # 세 번째 이미지 영역 선택
    print("세 번째 이미지를 선택하세요.")
    area3 = get_selection_area()
    image_path3 = "temp_screenshot3.png"
    capture_and_save_area(area3, image_path3)
    images_list.append(image_path3)

    # 세 번째 이미지 클릭 횟수를 사용자로부터 입력받음
    click_count3 = int(input("세 번째 이미지를 몇 번 클릭하시겠습니까? "))

  

    # 반복 실행
    for i in range(repeat_count):
        print(f"\n반복 {i+1}/{repeat_count}")
        click_image(image_path1, area1, click_count1, "첫")
        click_image(image_path2, area2, click_count2, "두")
        click_image(image_path3, area3, click_count3, "세")
 
    # 캡처한 이미지를 Excel 파일에 저장
    save_images_to_excel(images_list)

if __name__ == "__main__":
    main()
