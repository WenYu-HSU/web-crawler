from playwright.sync_api import sync_playwright, TimeoutError
import pandas as pd
from bs4 import BeautifulSoup
import time
import random

# 替換為你的 Threads 帳號與密碼
THREADS_USERNAME = "1234abc432199"
THREADS_PASSWORD = "hvkY!StrwRz3@Q7"

# 模擬的 user-agent
USER_AGENT = "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/91.0.4472.124 Safari/537.36"

# 登入 Threads 的函式
def login_to_threads(playwright):
    # 啟動 Chromium 瀏覽器，開啟無痕並最大化
    browser = playwright.chromium.launch(headless=False, args=['--start-maximized', '--incognito'])

    # 設定瀏覽器上下文
    context = browser.new_context(
        user_agent=USER_AGENT,
        viewport={'width': 1280, 'height': 800},
        ignore_https_errors=True,
        java_script_enabled=True,
    )

    page = context.new_page()

    try:
        print("Navigating to Threads login page...")
        page.goto("https://www.threads.net/login", timeout=60000)
        page.wait_for_load_state('networkidle')

        # 輸入帳號
        print("Filling in username...")
        username_input = page.locator('input[placeholder="用戶名稱、手機號碼或電子郵件地址"]')
        username_input.clear()
        username_input.fill(THREADS_USERNAME)
        time.sleep(random.uniform(1, 3))

        # 輸入密碼
        print("Filling in password...")
        password_input = page.locator('input[placeholder="密碼"]')
        password_input.clear()
        password_input.fill(THREADS_PASSWORD)
        time.sleep(random.uniform(1, 3))

        # 點擊登入按鈕
        print("Clicking login button...")
        login_button = page.locator('//div[contains(@class, "xwhw2v2") and contains(text(), "登入")]')
        login_button.click()

        # 等待頁面載入
        page.wait_for_load_state('networkidle')
        page.wait_for_timeout(10000)  # 等待 10 秒

        # 驗證是否登入成功
        current_url = page.url
        if "login_success=true" in current_url or "?" not in current_url:
            print("Login successful!")
            print(f"Current URL: {current_url}")

            # 儲存登入憑證
            storage_state = context.storage_state()
            with open("auth.json", "w") as f:
                f.write(str(storage_state))

            return context
        else:
            print("Login failed, please check your credentials.")
            print(f"Current URL: {current_url}")
            return None

    except Exception as e:
        print(f"An error occurred: {str(e)}")
        page.screenshot(path="error_screenshot.png")

    finally:
        print("Login process completed.")

# 導航到指定網址的函式
def navigate_to_url(context, target_url):
    try:
        page = context.new_page()
        print(f"Navigating to {target_url}")
        page.goto(target_url, timeout=60000)

        # 模擬使用者滑動頁面
        for _ in range(3):
            page.mouse.wheel(0, random.randint(300, 600))
            time.sleep(random.uniform(2, 5))

        # 等待網頁完全載入
        page.wait_for_load_state('networkidle')
        page.wait_for_timeout(10000)

        # 確認網頁是否成功載入
        if page.url == target_url:
            print(f"Successfully navigated to {target_url}")
            return page.content()
        else:
            print(f"Failed to navigate to the target page. Current URL: {page.url}")
            return None

    except TimeoutError as e:
        print(f"Navigation to {target_url} failed due to timeout: {str(e)}")
        return None

    finally:
        page.close()

# 分析 HTML 內容並擷取貼文資料
def parse_threads_content(content):
    if not content:
        print("No content to parse.")
        return []

    soup = BeautifulSoup(content, 'html.parser')

    # 嘗試多個選擇器找出貼文區塊
    selectors = [
        'div[role="article"]',
        'article',
        'div[data-pressable-container="true"]',
    ]

    articles = []
    for selector in selectors:
        articles = soup.select(selector)
        if articles:
            print(f"找到 {len(articles)} 個元素，使用選擇器: {selector}")
            break

    if not articles:
        print("無法找到任何貼文內容")
        return []

    parsed_data = []

    # 擷取每篇文章的內容
    for i, article in enumerate(articles):
        post_data = {
            'user_id': '',
            'time': '',
            'content': '',
            'likes': '0',
            'comments': '0',
            'reposts': '0',
            'shares': '0',
            'type': 'Main Post' if i == 0 else 'Reply'
        }

        # 擷取使用者 ID
        user_element = article.select_one('a[href^="/@"]')
        if user_element:
            post_data['user_id'] = user_element.get('href', '').strip('/@')

        # 擷取時間
        time_element = article.select_one('time')
        if time_element:
            post_data['time'] = time_element.get('datetime', '')

        # 擷取貼文內容
        content_selectors = [
            'h1.x1lliihq.x1plvlek.xryxfnj.x1n2onr6.x1ji0vk5.x18bv5gf.x193iq5w.xeuugli.x1fj9vlw.x13faqbe.x1vvkbs.x1s928wv.xhkezso.x1gmr53x.x1cpjm7i.x1fgarty.x1943h6x.x1i0vuye.xjohtrz.xo1l8bm.xp07o12.x1yc453h.xat24cr.xdj266r',
            'span.x1lliihq.x1plvlek.xryxfnj.x1n2onr6.x1ji0vk5.x18bv5gf.x193iq5w.xeuugli.x1fj9vlw.x13faqbe.x1vvkbs.x1s928wv.xhkezso.x1gmr53x.x1cpjm7i.x1fgarty.x1943h6x.x1i0vuye.xjohtrz.xo1l8bm.xp07o12.x1yc453h.xat24cr.xdj266r',
            'div[data-content-type="text"]'
        ]

        for content_selector in content_selectors:
            content_element = article.select_one(content_selector)
            if content_element:
                post_data['content'] = content_element.get_text(strip=True)
                break

        if not post_data['content']:
            print(f"警告: 無法找到內容 {'主貼文' if i == 0 else f'回覆 #{i}'}")

        # 擷取互動數據（讚、分享等）
        stat_elements = article.select('div.x6s0dn4.xfkn95n.xly138o.xchwasx.xfxlei4.x78zum5.xl56j7k.x1n2onr6.x3oybdh.x12w9bfk.xx6bhzk.x11xpdln.xc9qbxq.x1ye3gou.xn6708d.x14atkfc')

        for stat_element in stat_elements:
            svg = stat_element.find('svg')
            if svg:
                aria_label = svg.get('aria-label')
                print(f"找到 SVG 元素，aria-label: {aria_label}")
                value_span = stat_element.select_one('span span.x17qophe.x10l6tqk.x13vifvy')
                if value_span:
                    value = value_span.get_text(strip=True)
                    print(f"找到值: {value}")
                    if aria_label == '讚':
                        post_data['likes'] = value
                    elif aria_label == '轉發':
                        post_data['reposts'] = value
                    elif aria_label == '分享':
                        post_data['shares'] = value
                    elif aria_label == '回覆':
                        post_data['comments'] = value
                else:
                    print(f"未找到值 span for {aria_label}")
            else:
                print("未找到 SVG 元素")

        print(f"文章 {i+1} 的最終數據: {post_data}")
        parsed_data.append(post_data)

    return parsed_data

# 將資料匯出為 Excel
def export_to_excel(data, filename):
    df = pd.DataFrame(data)
    df.to_excel(filename, index=False)
    print(f"數據已保存到 {filename}")

# 主流程
def main():
    with sync_playwright() as playwright:
        context = login_to_threads(playwright)
        if not context:
            print("Login failed. Exiting program.")
            return

        # 目標網址，可修改為其他 threads 連結
        target_url = "https://www.threads.net/@chalan5946/post/DByRUnAyWby?hl=zh-tw"
        content = navigate_to_url(context, target_url)

        if content:
            parsed_data = parse_threads_content(content)
            if parsed_data:
                export_to_excel(parsed_data, "threads_data.xlsx")
        else:
            print("Failed to retrieve content from the target URL.")

if __name__ == "__main__":
    main()