import pyautogui
import time
from pynput import mouse
from PIL import ImageGrab, Image
import os
import ctypes

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

# 이미지를 지정된 영역에서 검색
def search_image_in_area(image_path, area):
    # 이미지의 좌표를 지정된 영역 내에서 검색합니다.
    try:
        location = pyautogui.locateCenterOnScreen(image_path, region=area, confidence=0.8)
        if location:
            return location
    except Exception as e:
        show_error_popup(f"Error: {e}")
    return None

def capture_and_save_area(area, filename):
    try:
        # 영역을 캡처한 후 저장합니다.
        screenshot = ImageGrab.grab(bbox=area)
        screenshot.save(filename)
        screenshot.show()
        print(f"영역이 {filename}로 저장되었습니다.")
    except Exception as e:
        show_error_popup(f"Error: {e}")

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
        show_error_popup(f"{image_path} 이미지를 지정된 영역에서 찾을 수 없습니다.")

def show_error_popup(message):
    ctypes.windll.user32.MessageBoxW(0, message, "Error", 0x10)

def main():
    # 전체 프로세스를 몇 번 실행할지 사용자로부터 입력받음
    repeat_count = int(input("프로세스를 몇 번 실행할까요? "))

    # 첫 번째 이미지 영역 선택
    print("첫 번째 이미지를 선택하세요.")
    area1 = get_selection_area()
    image_path1 = "temp_screenshot1.png"
    capture_and_save_area(area1, image_path1)

    # 첫 번째 이미지 클릭 횟수를 사용자로부터 입력받음
    click_count1 = int(input("첫 번째 이미지를 몇 번 클릭하시겠습니까? "))

    # 두 번째 이미지 영역 선택
    print("두 번째 이미지를 선택하세요.")
    area2 = get_selection_area()
    image_path2 = "temp_screenshot2.png"
    capture_and_save_area(area2, image_path2)

    # 두 번째 이미지 클릭 횟수를 사용자로부터 입력받음
    click_count2 = int(input("두 번째 이미지를 몇 번 클릭하시겠습니까? "))

    # 세 번째 이미지 영역 선택
    print("세 번째 이미지를 선택하세요.")
    area3 = get_selection_area()
    image_path3 = "temp_screenshot3.png"
    capture_and_save_area(area3, image_path3)

    # 세 번째 이미지 클릭 횟수를 사용자로부터 입력받음
    click_count3 = int(input("세 번째 이미지를 몇 번 클릭하시겠습니까? "))

    # 네 번째 이미지 영역 선택
    print("네 번째 이미지를 선택하세요.")
    area4 = get_selection_area()
    image_path4 = "temp_screenshot4.png"
    capture_and_save_area(area4, image_path4)

    # 네 번째 이미지 클릭 횟수를 사용자로부터 입력받음
    click_count4 = int(input("네 번째 이미지를 몇 번 클릭하시겠습니까? "))

    # 다섯 번째 이미지 영역 선택
    print("다섯 번째 이미지를 선택하세요.")
    area5 = get_selection_area()
    image_path5 = "temp_screenshot5.png"
    capture_and_save_area(area5, image_path5)

    # 다섯 번째 이미지 클릭 횟수를 사용자로부터 입력받음
    click_count5 = int(input("다섯 번째 이미지를 몇 번 클릭하시겠습니까? "))

    # 여섯 번째 이미지 영역 선택
    print("여섯 번째 이미지를 선택하세요.")
    area6 = get_selection_area()
    image_path6 = "temp_screenshot6.png"
    capture_and_save_area(area6, image_path6)

    # 여섯 번째 이미지 클릭 횟수를 사용자로부터 입력받음
    click_count6 = int(input("여섯 번째 이미지를 몇 번 클릭하시겠습니까? "))

    # 일곱 번째 이미지 영역 선택
    print("일곱 번째 이미지를 선택하세요.")
    area7 = get_selection_area()
    image_path7 = "temp_screenshot7.png"
    capture_and_save_area(area7, image_path7)

    # 일곱 번째 이미지 클릭 횟수를 사용자로부터 입력받음
    click_count7 = int(input("일곱 번째 이미지를 몇 번 클릭하시겠습니까? "))

    # 반복 실행
    for i in range(repeat_count):
        print(f"\n반복 {i+1}/{repeat_count}")
        click_image(image_path1, area1, click_count1, "첫")
        click_image(image_path2, area2, click_count2, "두")
        click_image(image_path3, area3, click_count3, "세")
        click_image(image_path4, area4, click_count4, "네")
        click_image(image_path5, area5, click_count5, "다섯")
        click_image(image_path6, area6, click_count6, "여섯")
        click_image(image_path7, area7, click_count7, "일곱")


if __name__ == "__main__":
    main()
