from playwright.sync_api import sync_playwright, TimeoutError
import pandas as pd
from bs4 import BeautifulSoup
import time

# 替換為你的 Threads 帳號與密碼
THREADS_USERNAME = "1234abc432199"
THREADS_PASSWORD = "hvkY!StrwRz3@Q7"


def login_to_threads(playwright):
    browser = playwright.chromium.launch(
        headless=False,
        args=['--start-maximized']  # 最大化視窗
    )
    context = browser.new_context()
    page = context.new_page()

    try:
        print("Navigating to Threads login page...")
        page.goto("https://www.threads.net/login", timeout=60000)
        page.wait_for_load_state('networkidle')

        # 填寫用戶名和密碼前先清空輸入框
        print("Filling in username...")
        username_input = page.locator('input[placeholder="用戶名稱、手機號碼或電子郵件地址"]')
        username_input.clear()
        username_input.fill(THREADS_USERNAME)
        time.sleep(2)  # 短暫延遲

        print("Filling in password...")
        password_input = page.locator('input[placeholder="密碼"]')
        password_input.clear()
        password_input.fill(THREADS_PASSWORD)
        time.sleep(2)  # 短暫延遲

        # 點擊登入按鈕
        print("Clicking login button...")
        login_button = page.locator('//div[contains(@class, "xwhw2v2") and contains(text(), "登入")]')
        login_button.click()

        # 等待登入處理
        page.wait_for_load_state('networkidle')
        page.wait_for_timeout(10000)  # 延遲 10 秒

        # 檢查登入是否成功，修改邏輯
        current_url = page.url
        if "login_success=true" in current_url or "?" not in current_url:
            print("Login successful!")
            print(f"Current URL: {current_url}")
            
            # 儲存 cookies 和 storage state
            storage_state = context.storage_state()
            with open("auth.json", "w") as f:
                f.write(str(storage_state))
                
            return context  # 返回已登入的上下文，保持登入狀態

        else:
            print("Login failed, please check your credentials.")
            print(f"Current URL: {current_url}")
            return None

    except Exception as e:
        print(f"An error occurred: {str(e)}")
        page.screenshot(path="error_screenshot.png")

    finally:
        print("Login process completed.")

def navigate_to_url(context, target_url):
    try:
        page = context.new_page()
        print(f"Navigating to {target_url}")
        page.goto(target_url, timeout=60000)

        # 增加延遲以確保網頁加載
        page.wait_for_load_state('networkidle')
        page.wait_for_timeout(10000)

        # 滾動頁面以加載更多內容
        print("Scrolling to load more content...")
        for _ in range(10):  # 調整滾動次數
            page.evaluate("window.scrollTo(0, document.body.scrollHeight)")
            time.sleep(5)

        # 確保已成功導航到正確的網址
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

def parse_threads_content(content):
    if not content:
        print("No content to parse.")
        return []

    soup = BeautifulSoup(content, 'html.parser')

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

        user_element = article.select_one('a[href^="/@"]')
        if user_element:
            post_data['user_id'] = user_element.get('href', '').strip('/@')

        time_element = article.select_one('time')
        if time_element:
            post_data['time'] = time_element.get('datetime', '')

        # 嘗試多個選擇器來找到內容
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

        # 修改這部分以使用新的評論提取邏輯
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

def export_to_excel(data, filename='threads.xlsx'):
    df = pd.DataFrame(data)
    df.to_excel(filename, index=False)
    print(f"Data exported to {filename}")

if __name__ == "__main__":
    target_url = "https://www.threads.net/@stupidcooldog/post/DCBHOTsTV57?xmt=AQGzlCLQNWHmLQhXT3iCa2SYfNQEyD2xQu6q0e53ObEtqzM"
    try:
        with sync_playwright() as p:
            # 登入並保持登入狀態
            context = login_to_threads(p)
            if context:
                # 使用已登入的會話導航到目標網頁
                content = navigate_to_url(context, target_url)

                # 解析內容並匯出
                parsed_data = parse_threads_content(content)
                export_to_excel(parsed_data)

    except Exception as e:
        print(f"An error occurred: {str(e)}")
