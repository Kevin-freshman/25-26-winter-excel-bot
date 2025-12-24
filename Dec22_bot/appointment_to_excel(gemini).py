from playwright.sync_api import sync_playwright
import pandas as pd
import time
import os
import re
from datetime import datetime

# ---------------- 配置信息 ----------------
URL = "https://emsvip.linkedlife.cn/"
COMPANY = "xm-lf"
USERNAME = "前台"
PASSWORD = "123"
EXCEL_PATH = "appointments.xlsx"

# ---------------- 工具函数 ----------------

def parse_date_time(raw_time_str):
    """解析时间字符串"""
    if not raw_time_str:
        return "", ""
    try:
        parts = raw_time_str.split("-")[0].strip()
        dt = datetime.strptime(parts, "%Y/%m/%d %H:%M")
        date_str = dt.strftime("%m月%d日").lstrip("0").replace("月0", "月")
        time_str = dt.strftime("%H:%M")
        return date_str, time_str
    except Exception as e:
        return raw_time_str, ""

def get_next_index():
    """获取 Excel 下一个序号"""
    if not os.path.exists(EXCEL_PATH): return 1
    try:
        df = pd.read_excel(EXCEL_PATH)
        if "序号" in df.columns and not df.empty:
            return int(df["序号"].max()) + 1
        return 1
    except:
        return 1

def save_to_excel(raw_data: dict):
    """保存到 Excel"""
    date_str, time_str = parse_date_time(raw_data.get("预约时间", ""))
    new_row = {
        "序号": get_next_index(),
        "上门日期": date_str,
        "具体时间": time_str,
        "顾客姓名": raw_data.get("姓名", ""),
        "病历号/会员卡号": raw_data.get("会员号", ""),
        "来源渠道": raw_data.get("客户来源", "")
    }

    df_new = pd.DataFrame([new_row])

    if os.path.exists(EXCEL_PATH):
        df_old = pd.read_excel(EXCEL_PATH)
        if "病历号/会员卡号" in df_old.columns:
            df_old["病历号/会员卡号"] = df_old["病历号/会员卡号"].astype(str)
        df = pd.concat([df_old, df_new], ignore_index=True)
    else:
        df = df_new
        cols = ["序号", "上门日期", "具体时间", "顾客姓名", "病历号/会员卡号", "来源渠道"]
        df = df[cols]
    
    df.to_excel(EXCEL_PATH, index=False)
    print(f"✅ [写入成功] 序号: {new_row['序号']} | 姓名: {new_row['顾客姓名']}")

def is_blue_card(rgb_string: str) -> bool:
    """颜色判断逻辑"""
    if not rgb_string: return False
    match = re.search(r"rgb\((\d+),\s*(\d+),\s*(\d+)\)", rgb_string)
    if not match: return False
    r, g, b = map(int, match.groups())
    if b > r and b > 200: return True
    if b > 220 and r < 230: return True
    return False

def already_exists(member_id: str, date_check: str) -> bool:
    """防止重复录入"""
    if not os.path.exists(EXCEL_PATH): return False
    df = pd.read_excel(EXCEL_PATH)
    if "病历号/会员卡号" in df.columns and "上门日期" in df.columns:
        filtered = df[df["病历号/会员卡号"].astype(str) == str(member_id)]
        if not filtered.empty:
            if date_check in filtered["上门日期"].values:
                return True
    return False

# ---------------- 页面行为 ----------------

def login(page):
    print("正在登录...")
    # --- 增强点 1: 网络波动重试机制 ---
    max_retries = 3
    for attempt in range(max_retries):
        try:
            print(f"尝试连接网站 (第 {attempt+1} 次)...")
            page.goto(URL, wait_until="domcontentloaded", timeout=30000)
            break # 如果成功，跳出循环
        except Exception as e:
            print(f"⚠️ 连接失败: {e}")
            if attempt < max_retries - 1:
                print("等待 3 秒后重试...")
                time.sleep(3)
            else:
                raise Exception("无法连接到网站，请检查网络设置。")

    try:
        page.locator("input[type='text']").nth(0).fill(COMPANY)
        page.locator("input[type='text']").nth(1).fill(USERNAME)
        page.locator("input[type='password']").fill(PASSWORD)
        page.get_by_role("button", name="登 录").click()
        
        print("等待跳转...")
        page.wait_for_selector("text=预约", timeout=30000)
        page.wait_for_selector("text=预约中心", timeout=30000)
        print("登录成功。")
    except Exception as e:
        print(f"登录过程出错: {e}")

def goto_appointment_center(page):
    print("正在跳转到预约中心...")
    try:
        menu_btn = page.locator("li").filter(has_text="预约").first
        menu_btn.click()
        time.sleep(1)
        
        sub_menu_btn = page.locator("li").filter(has_text="预约中心").first
        if sub_menu_btn.is_visible():
            sub_menu_btn.click()
        else:
            sub_menu_btn.click(force=True)

        # 只要日历框架加载出来就算成功，内容可能还没出来
        page.wait_for_selector(".fc-view-container", timeout=20000)
        print("日历框架已加载，准备读取数据。")
        
    except Exception as e:
        print(f"跳转导航失败: {e}")

def extract_detail_from_modal(page) -> dict:
    """提取数据"""
    data = {}
    page.locator(".ant-modal-content").first.wait_for(timeout=5000)
    modal_text = page.locator(".ant-modal-body").inner_text()
    
    try:
        header = page.locator(".header-info").inner_text()
        lines = header.split('\n')
        name_candidate = lines[0].strip()
        data["姓名"] = re.sub(r'[^\u4e00-\u9fa5a-zA-Z]', '', name_candidate)
        id_match = re.search(r"\d{6,}", header)
        data["会员号"] = id_match.group(0) if id_match else ""
    except:
        data["姓名"] = "未知"
        data["会员号"] = ""

    for line in modal_text.split('\n'):
        if "：" in line:
            key, val = line.split("：", 1)
            data[key.strip()] = val.strip()
    return data

def process_appointments(page):
    # --- 增强点 2: 智能轮询等待 ---
    card_selector = "a.fc-day-grid-event"
    print("正在等待卡片渲染 (最多等待 100 秒)...")
    
    found_cards = False
    max_wait = 100  # 最大等待秒数
    
    for _ in range(max_wait // 2):
        count = page.locator(card_selector).count()
        if count > 0:
            found_cards = True
            break
        time.sleep(2) # 每 2 秒检查一次
        print("...加载中...")
    
    if not found_cards:
        print("⚠️ 100秒内未检测到任何预约卡片，可能是因为：")
        print("1. 今天确实没有预约。")
        print("2. 网速过慢导致加载超时。")
        print("程序结束。")
        return

    cards = page.locator(card_selector)
    count = cards.count()
    print(f"检测到 {count} 个预约卡片，开始处理。")

    for i in range(count):
        card = cards.nth(i)
        
        # 颜色判断
        bg_color = card.evaluate("el => window.getComputedStyle(el).backgroundColor")
        if not is_blue_card(bg_color):
            continue
        
        try:
            print(f"处理第 {i+1} 个卡片...")
            card.click()
            
            raw_data = extract_detail_from_modal(page)
            date_check, _ = parse_date_time(raw_data.get("预约时间", ""))
            
            if already_exists(raw_data.get("会员号"), date_check):
                print(f"   -> 跳过: {raw_data.get('姓名')} (已存在)")
            else:
                save_to_excel(raw_data)
            
            page.keyboard.press("Escape")
            time.sleep(0.5)
            
        except Exception as e:
            print(f"   -> 处理出错: {e}")
            page.keyboard.press("Escape")

# ---------------- 主程序 ----------------

def main():
    with sync_playwright() as p:
        # --- 增强点 3: 启动参数优化 ---
        browser = p.chromium.launch(
            headless=False, 
            args=[
                "--start-maximized", 
                "--disable-blink-features=AutomationControlled" # 防反爬
            ]
        )
        context = browser.new_context(no_viewport=True)
        page = context.new_page()

        # 设置页面默认超时时间为 30秒
        page.set_default_timeout(30000)

        login(page)
        goto_appointment_center(page)
        process_appointments(page)
        
        print("任务完成，3秒后退出...")
        time.sleep(3)
        browser.close()

if __name__ == "__main__":
    main()