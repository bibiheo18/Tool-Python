import pyautogui
# Chạy script dưới để xem tọa độ thủ công
print("Di chuột vào vị trí cần đo. Nhấn Ctrl+C để thoát.")
try:
    while True:
        x, y = pyautogui.position()
        print(f"Tọa độ hiện tại: ({x}, {y})", end="\r")
except KeyboardInterrupt:
    print("\nKết thúc.")