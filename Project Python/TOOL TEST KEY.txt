#Comment = Ctrl K + Ctrl C Uncomment = Ctrl K + Ctrl U
from selenium import webdriver
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.common.by import By
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from webdriver_manager.chrome import ChromeDriverManager
from selenium.webdriver.common.action_chains import ActionChains # Thư viện dùng click và nhấn nút
import time

def run_google_search():
    # Cấu hình Chrome giả lập người dùng thật
    options = Options()
    options.add_argument("--start-maximized")
    options.add_argument("--incognito")
    options.add_argument("--disable-blink-features=AutomationControlled")
    options.add_experimental_option("excludeSwitches", ["enable-automation"])
    options.add_experimental_option('useAutomationExtension', False)
    options.add_argument(
        "user-agent=Mozilla/5.0 (Windows NT 10.0; Win64; x64) "
        "AppleWebKit/537.36 (KHTML, like Gecko) Chrome/113.0.0.0 Safari/537.36"
    )

    # Tự động tải ChromeDriver
    service = Service(ChromeDriverManager().install())
    driver = webdriver.Chrome(service=service, options=options)

    # Xóa dấu hiệu 'webdriver' trong navigator (giảm bị phát hiện)
    driver.execute_cdp_cmd("Page.addScriptToEvaluateOnNewDocument", {
        "source": """
            Object.defineProperty(navigator, 'webdriver', {
              get: () => undefined
            })
        """
    })

    try:
        # Truy cập Google và tìm test keyboard
        driver.get("https://www.google.com/?hl=vi&theme=light")
        time.sleep(1)

        search_box = driver.find_element(By.NAME, "q")
        search_box.send_keys("test keyboard")
        search_box.submit()

        testkeyboard_link = WebDriverWait(driver, 10).until(
            EC.element_to_be_clickable((By.XPATH, '//h3[normalize-space()="Test Key | Keyboard Test | Test Keyboard Online"]/ancestor::a'))
        )
        testkeyboard_link.click()

        # ✅ Scroll xuống vị trí Y = 200
        time.sleep(1)  # Chờ trang load trước khi scroll
        driver.execute_script("window.scrollTo(0, 200);")
        print("✅ Đã scroll xuống Y = 200.")

        # ✅ Di chuyển đến tọa độ (380, 650) và click 3 lần
        actions = ActionChains(driver)
        for i in range(3):
            actions.move_by_offset(380, 650).click().perform()
            actions.move_by_offset(-380, -650)  # Reset lại để click lần sau
            time.sleep(0.5)
        print("✅ Đã click tại (380, 650) ba lần.")

        # ✅ Nhấn phím 'A' 3 lần
        for i in range(3):
            actions.send_keys('a')  # hoặc Keys.A
            time.sleep(0.5)
        actions.perform()
        print("✅ Đã nhấn phím 'A' 3 lần.")

        # ✅ Chờ quảng cáo hiển thị và đóng nó nếu có
        try:
            close_ad_button = WebDriverWait(driver, 5).until(
                EC.element_to_be_clickable((By.ID, "dismiss-button"))
            )
            close_ad_button.click()
            print("✅ Đã đóng quảng cáo.")
        except:
            print("ℹ️ Không thấy quảng cáo hoặc quảng cáo không cần đóng.")




        # Chờ thêm nếu cần thao tác tiếp
        time.sleep(180)

    except Exception as e:
        print("❌ Lỗi:", e)

    finally:
        driver.quit()

run_google_search()

        # Đợi và nhấn nút "Log in"
        # login_button = WebDriverWait(driver, 15).until(
        #     EC.element_to_be_clickable((By.CSS_SELECTOR, '[data-testid="login-button"]'))
        # )
        # login_button.click()