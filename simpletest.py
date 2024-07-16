import pyautogui
import time
from pynput import mouse
from PIL import ImageGrab

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
    screenshot = ImageGrab.grab(bbox=area)
    screenshot.save("temp_screenshot.png")  # 임시로 저장
    try:
        location = pyautogui.locateCenterOnScreen(image_path, region=area)
        if location:
            pyautogui.click(location)
            return location
    except pyautogui.ImageNotFoundException:
        print(f"Error: {image_path} 이미지가 지정된 영역에서 발견되지 않았습니다.")
    return None

def main():
    area = get_selection_area()
    image_path = r"C:\Users\simoony\auto\1st.png"  # 찾고자 하는 이미지의 경로
    location = search_image_in_area(image_path, area)
    if location:
        print(f"이미지를 {location} 위치에서 찾았습니다.")
    else:
        print("이미지를 지정된 영역에서 찾을 수 없습니다.")

if __name__ == "__main__":
    main()
