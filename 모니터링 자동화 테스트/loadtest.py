import pyautogui
import cv2
import numpy as np
import matplotlib.pyplot as plt

# 스케일링 단계
scales = [1.0, 0.9, 1.1]

def find_image_on_screen(template_path, scales):
    # 전체 화면 스크린샷 찍기
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

        # 매칭 임계값 설정
        threshold = 0.8
        if max_val >= threshold:
            return max_loc, width, height

    return None

# 테스트용 코드
def test_image_matching():
    # 테스트를 위해 사용할 템플릿 이미지 경로 설정
    template_path = "C:\\Users\\simoony\\auto\\ROLDLINE\\play.png"  # 실제 테스트 이미지 경로로 변경

    # 이미지를 찾고 결과를 출력
    result = find_image_on_screen(template_path, scales)
    if result:
        max_loc, width, height = result
        print(f"이미지를 찾았습니다: 위치 {max_loc}, 크기 ({width}, {height})")

        # 전체 화면 스크린샷 다시 찍기 (시각화를 위해)
        screenshot = pyautogui.screenshot()
        screenshot = cv2.cvtColor(np.array(screenshot), cv2.COLOR_RGB2BGR)
        cv2.rectangle(screenshot, max_loc, (max_loc[0] + width, max_loc[1] + height), (0, 255, 0), 2)

        # 결과 보여주기
        plt.imshow(cv2.cvtColor(screenshot, cv2.COLOR_BGR2RGB))
        plt.title(f'Found: {max_loc}')
        plt.show()
    else:
        print("이미지를 찾지 못했습니다.")

# 테스트 코드 실행
test_image_matching()
