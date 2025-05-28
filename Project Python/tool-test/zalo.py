##################### TOOL ZALO LẤY DS SDT EXCEL TÌM TRÊN ZALO VÀ CẬP NHẬT TRẠNG THÁI #########################

import pyautogui            # Thư viện điều khiển chuột, bàn phím tự động
import pygetwindow as gw    # Thư viện để lấy danh sách và thao tác cửa sổ ứng dụng
import subprocess           # Dùng để mở ứng dụng (Excel, Zalo)
import time                 # Thư viện dùng để delay giữa các thao tác
import pyperclip            # Thư viện để thao tác với clipboard (copy/paste nhanh, an toàn)

# ========== CẤU HÌNH ==========
excel_path = r"D:\Project Python\tool-test\dsUser.xlsx"  # Đường dẫn file Excel chứa danh sách SĐT
zalo_path = r"C:\Users\bibih\AppData\Local\Programs\Zalo\Zalo.exe"  # Đường dẫn ứng dụng Zalo
start_row = 2        # Dòng bắt đầu (bỏ dòng tiêu đề)
end_row = 30         # Dòng kết thúc (số dòng cần kiểm tra)

# ========== KHỞI ĐỘNG ỨNG DỤNG ==========
subprocess.Popen(["start", "excel", excel_path], shell=True)  # Mở Excel
subprocess.Popen(zalo_path)                                    # Mở Zalo
time.sleep(3)  # Chờ ứng dụng load xong (giảm từ 5s → 3s để nhanh hơn)

# ========== HÀM HỖ TRỢ ==========

# Chuyển sang cửa sổ ứng dụng theo tiêu đề (Excel hoặc Zalo)
def switch_to_window(title):
    windows = gw.getWindowsWithTitle(title)
    if windows:
        win = windows[0]
        win.activate()   # Đưa cửa sổ lên trước
        win.maximize()   # Phóng to toàn màn hình
        time.sleep(0.3)  # Delay ngắn để đảm bảo cửa sổ đã sẵn sàng

# Đi tới và chọn ô Excel cụ thể (vd: C5, F10)
def goto_and_focus_cell(cell_ref):
    pyautogui.hotkey("ctrl", "g")      # Mở hộp thoại "Go To"
    time.sleep(0.1)
    pyperclip.copy(cell_ref)           # Copy địa chỉ ô vào clipboard
    pyautogui.hotkey("ctrl", "v")      # Dán vào hộp thoại
    pyautogui.press("enter")           # Nhảy đến ô chỉ định
    time.sleep(0.1)
    pyautogui.press("esc")             # Thoát khỏi hộp thoại
    time.sleep(0.1)

# Xóa ô tìm kiếm trong Zalo trước khi tìm mới
def clear_zalo_search():
    pyautogui.moveTo(147, 70, duration=0)  # Di chuyển nhanh đến ô tìm kiếm Zalo (tọa độ cố định)
    pyautogui.click()                      # Click vào ô
    pyautogui.hotkey("ctrl", "a")          # Chọn hết nội dung
    pyautogui.press("delete")              # Xóa nội dung cũ
    time.sleep(0.1)

# ========== BẮT ĐẦU KIỂM TRA DANH SÁCH SĐT ==========
for row in range(start_row, end_row + 1):
    cell_sdt = f"C{row}"           # Ô chứa số điện thoại (cột C)
    cell_trang_thai = f"F{row}"    # Ô ghi trạng thái kết quả (cột F)

    # Lấy số điện thoại từ Excel
    switch_to_window("Excel")             # Chuyển sang Excel
    goto_and_focus_cell(cell_sdt)         # Di chuyển tới ô chứa SĐT
    pyautogui.hotkey("ctrl", "c")         # Copy SĐT vào clipboard
    time.sleep(0.1)
    sdt = pyperclip.paste().strip()       # Lấy SĐT và xóa khoảng trắng

    if not sdt:
        continue  # Bỏ qua nếu ô rỗng

    print(f"🔍 Đang kiểm tra: {sdt}")

    # Tìm kiếm trong Zalo
    switch_to_window("Zalo")              # Chuyển sang Zalo
    clear_zalo_search()                   # Xóa tìm kiếm cũ
    pyperclip.copy(sdt)                   # Copy SĐT
    pyautogui.hotkey("ctrl", "v")         # Dán vào ô tìm kiếm Zalo
    time.sleep(1.5)                       # Chờ Zalo hiện kết quả (giảm từ 2s → 1.5s)

    # Kiểm tra hình ảnh báo không tìm thấy hoặc chưa đăng ký
    try:
        not_found = pyautogui.locateOnScreen("not_found.png", confidence=0.6)
    except:
        not_found = None

    try:
        not_registered = pyautogui.locateOnScreen("not_registered.png", confidence=0.6)
    except:
        not_registered = None

    # Xử lý kết quả
    if not_registered:
        print(f"⚠️ SĐT chưa đăng ký hoặc không cho phép tìm kiếm: {sdt}")
        status = "SĐT chưa đăng ký"
    elif not_found:
        print(f"❌ Không tìm thấy người này: {sdt}")
        status = "Không tìm thấy người này"
    else:
        print(f"✅ Tìm thấy: {sdt}")
        # Click vào kết quả tìm thấy để xác nhận
        pyautogui.moveTo(140, 235, duration=0.05)
        pyautogui.click()
        time.sleep(0.5)
        status = "X"

    # Ghi kết quả trở lại Excel
    switch_to_window("Excel")                 # Quay lại Excel
    goto_and_focus_cell(cell_trang_thai)      # Di chuyển đến ô ghi kết quả
    pyperclip.copy(status)                    # Copy nội dung trạng thái
    pyautogui.hotkey("ctrl", "v")             # Dán vào ô
    pyautogui.press("enter")                  # Xác nhận
    print(f"✔️ Đã ghi '{status}' vào {cell_trang_thai}")

    time.sleep(0.3)  # Delay nhẹ trước khi qua dòng tiếp theo

print("🎉 Đã hoàn thành kiểm tra tất cả số điện thoại.")
