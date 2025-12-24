from playwright.sync_api import sync_playwright
import pandas as pd
import time
import os

URL = "https://emsvip.linkedlife.cn/"
COMPANY = "xm-lf"
USERNAME = "前台"
PASSWORD = "123"
EXCEL_PATH = "appointments.xlsx"


# ---------------- 工具函数 ----------------

def clean_id(raw: str) -> str:
    """去掉后两位"""
    return raw[:-2] if raw and len(raw) > 2 else raw


def save_to_excel(row: dict):
    df_new = pd.DataFrame([row])

    if os.path.exists(EXCEL_PATH):
        df_old = pd.read_excel(EXCEL_PATH)
        if row["会员号"] in df_old["会员号"].values:
            print("已存在，跳过：", row["会员号"])
            return
        df = pd.concat([df_old, df_new], ignore_index=True)
    else:
        df = df_new

    df.to_excel(EXCEL_PATH, index=False)
    print("写入 Excel:", row["客户"])


# ---------------- 页面行为 ----------------

def login(page):
    page.goto(URL, wait_until="domcontentloaded", timeout=60000)
    page.wait_for_selector("input", timeout=20000)

    inputs = page.locator("input")

    inputs.nth(0).type(COMPANY, delay=100)
    inputs.nth(1).type(USERNAME, delay=100)
    inputs.nth(2).type(PASSWORD, delay=100)

    page.keyboard.press("Tab")
    time.sleep(0.5)

    page.get_by_role("button", name="登 录").click()
    page.wait_for_timeout(3000)


def goto_appointment_center(page):
    page.get_by_text("预约", exact=True).click()
    page.wait_for_timeout(5000)
    page.get_by_text("预约中心", exact=True).click()
    page.wait_for_timeout(5000)


def extract_detail(page) -> dict:
    """
    从【预约详情页】抓取完整字段
    """
    data = {}

    # 会员号（页面顶部 / 高亮区域）
    raw_id = page.locator("span.ng-star-inserted").first.inner_text().strip()
    data["会员号"] = clean_id(raw_id)

    items = page.locator("div.appointment-detail-wrap div.item")

    for i in range(items.count()):
        label = items.nth(i).locator(".label").inner_text().strip().replace("：", "")
        value = items.nth(i).locator(".content").inner_text().strip()
        data[label] = value

    # 关键字段标准化
    data["客户"] = data.get("客户", "")
    data["预约时间"] = data.get("预约时间", "")
    data["医生"] = data.get("医生", "")
    data["咨询师"] = data.get("咨询师", "")
    data["项目"] = data.get("项目", "")

    return data


def process_all_cards(page):
    """
    遍历预约卡片 → 点击 → 抓详情 → 返回
    """
    cards = page.locator("div[class*='event'], div[class*='appointment']")
    print("检测到预约卡片数量：", cards.count())

    for i in range(cards.count()):
        card = cards.nth(i)

        # 滚动到可见
        card.scroll_into_view_if_needed()
        time.sleep(0.5)

        # 点击进入详情
        card.click()
        page.wait_for_selector("div.appointment-detail-wrap", timeout=15000)

        # 抓详情
        data = extract_detail(page)
        save_to_excel(data)

        # 返回预约中心
        page.go_back(wait_until="domcontentloaded")
        page.wait_for_timeout(1500)


# ---------------- 主程序 ----------------

def main():
    with sync_playwright() as p:
        browser = p.chromium.launch(
            headless=False,
            args=["--disable-blink-features=AutomationControlled"]
        )
        page = browser.new_page()

        login(page)
        goto_appointment_center(page)
        process_all_cards(page)

        browser.close()


if __name__ == "__main__":
    main()
