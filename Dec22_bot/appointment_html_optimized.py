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
    
    # 构建数据行
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
    try:
        page.goto(URL, wait_until="domcontentloaded", timeout=30000)
    except Exception as e:
        print(f"⚠️ 首次连接超时，正在重试... ({e})")
        page.goto(URL, wait_until="domcontentloaded", timeout=30000)

    try:
        page.locator("input[type='text']").nth(0).fill(COMPANY)
        page.locator("input[type='text']").nth(1).fill(USERNAME)
        page.locator("input[type='password']").fill(PASSWORD)
        page.get_by_role("button", name="登 录").click()
        
        # 等待左侧菜单加载
        page.wait_for_selector("text=预约", timeout=30000)
        page.wait_for_selector("text=预约中心", timeout=30000)
        print("登录成功。")
    except Exception as e:
        print(f"登录过程出错: {e}")

def goto_appointment_center(page):
    print("正在跳转到预约中心...")
    try:
        # 1. 点击一级菜单 "预约"
        menu_btn = page.locator("li").filter(has_text="预约").first
        menu_btn.click()
        time.sleep(1)
        
        # 2. 点击二级菜单 "预约中心"
        sub_menu_btn = page.locator("li").filter(has_text="预约中心").first
        if sub_menu_btn.is_visible():
            sub_menu_btn.click()
        else:
            sub_menu_btn.click(force=True)
        
        time.sleep(5)

        # 3. 关键修复：检查是否在“预约列表”视图，如果是，点击“预约视图”切换到日历模式
        # 这一步是为了解决你提到的“必须再点一次才能加载”的问题
        view_tab = page.locator("div").filter(has_text="预约视图").last
        if view_tab.is_visible():
            print("正在切换到【预约视图】(日历模式)...")
            view_tab.click()
        else:
            print("未找到视图切换按钮，尝试直接寻找卡片...")

        # 4. 等待卡片容器加载 (基于你提供的 HTML class)
        # 等待任意一个卡片容器出现，或者等待日历网格出现
        page.wait_for_selector(".appointment-block-container, .fc-view-container", timeout=15000)
        print("日历视图加载完成。")
        
    except Exception as e:
        print(f"跳转导航警告: {e}")
        print("尝试继续执行...")

def extract_detail_from_modal(page) -> dict:
    """提取弹窗数据"""
    data = {}
    # 等待弹窗内容
    page.locator(".ant-modal-content").first.wait_for(timeout=5000)
    modal_text = page.locator(".ant-modal-body").inner_text()
    
    try:
        # 头部信息
        header = page.locator(".header-info").inner_text()
        # 提取会员号
        id_match = re.search(r"\d{6,}", header)
        data["会员号"] = id_match.group(0) if id_match else ""
        
        # 提取姓名 (简单逻辑：取头部第一行非数字部分)
        lines = header.split('\n')
        name_candidate = lines[0].strip()
        data["姓名"] = re.sub(r'[^\u4e00-\u9fa5a-zA-Z]', '', name_candidate)
        
    except:
        data["姓名"] = "未知"
        data["会员号"] = ""

    # 提取键值对
    for line in modal_text.split('\n'):
        if "：" in line:
            key, val = line.split("：", 1)
            data[key.strip()] = val.strip()
            
    return data

def process_appointments(page):
    # --- 核心修复：直接使用 CSS Class 定位蓝色卡片 ---
    # 你提供的 HTML 显示蓝色卡片 class 为 "appointment-block-container blue"
    # 我们直接定位它，不需要再算颜色了！
    
    blue_card_selector = "div.appointment-block-container.blue"
    
    print("正在扫描【蓝色/已到店】卡片...")
    
    # 智能等待：给页面一点时间渲染
    try:
        # 等待至少一个蓝色卡片出现，或者 5秒后超时
        page.wait_for_selector(blue_card_selector, timeout=50000)
    except:
        print("⚠️ 未检测到任何蓝色(已完成)卡片。")
        # 可能是还没加载出来，再给一点宽限时间
        time.sleep(10)

    cards = page.locator(blue_card_selector)
    count = cards.count()
    print(f"--> 发现 {count} 个蓝色卡片待处理。")

    if count == 0:
        return

    for i in range(count):
        # 重新定位防止元素过期
        card = cards.nth(i)
        
        try:
            # 获取卡片上的名字方便日志输出 (从你给的HTML结构看，名字在 h3.user-name)
            try:
                card_name = card.locator(".user-name").inner_text().strip()
            except:
                card_name = "未知客户"

            print(f"正在处理: {card_name} ...")
            
            # 点击卡片
            card.click()
            
            # 提取详情
            raw_data = extract_detail_from_modal(page)
            
            # 查重逻辑
            date_check, _ = parse_date_time(raw_data.get("预约时间", ""))
            
            if already_exists(raw_data.get("会员号"), date_check):
                print(f"   -> 跳过 (已存在)")
            else:
                save_to_excel(raw_data)
            
            # 关闭弹窗
            page.keyboard.press("Escape")
            # 等待弹窗消失动画
            time.sleep(0.5)
            
        except Exception as e:
            print(f"   -> 处理出错: {e}")
            # 出错后尝试按 ESC 复位
            page.keyboard.press("Escape")
            time.sleep(1)

# ---------------- 主程序 ----------------

def main():
    with sync_playwright() as p:
        browser = p.chromium.launch(
            headless=False, 
            args=["--start-maximized", "--disable-blink-features=AutomationControlled"]
        )
        context = browser.new_context(no_viewport=True)
        page = context.new_page()
        page.set_default_timeout(30000)

        login(page)
        goto_appointment_center(page)
        process_appointments(page)
        
        print("\n所有任务完成，程序将在 5 秒后关闭...")
        time.sleep(5)
        browser.close()

if __name__ == "__main__":
    main()