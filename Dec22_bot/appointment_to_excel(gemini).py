from playwright.sync_api import sync_playwright
import pandas as pd
import time
import os
import re

# ---------------- 配置信息 ----------------
URL = "https://emsvip.linkedlife.cn/"
COMPANY = "xm-lf"
USERNAME = "前台"
PASSWORD = "123"
EXCEL_PATH = "appointments.xlsx"

# ---------------- 工具函数 ----------------

def clean_id(raw: str) -> str:
    """清洗会员ID，去除可能的干扰字符"""
    return raw.strip() if raw else ""

def is_blue_card(rgb_string: str) -> bool:
    """
    判断背景颜色是否为浅蓝色。
    逻辑：解析 'rgb(r, g, b)' 格式。
    黄色/米色通常是 Red 和 Green 很高，Blue 较低 (如 255, 255, 200)。
    蓝色通常是 Blue 值最高，或者 Red 值明显较低。
    """
    if not rgb_string:
        return False
    
    # 使用正则提取数字
    match = re.search(r"rgb\((\d+),\s*(\d+),\s*(\d+)\)", rgb_string)
    if not match:
        return False # 无法识别颜色，默认跳过
    
    r, g, b = map(int, match.groups())

    # --- 颜色判断阈值 (根据经验调整) ---
    # 浅蓝色特征: Blue 分量通常很高 (例如 > 200)，且通常大于 Red
    # 浅黄色特征: Red 和 Green 很高，Blue 相对较低
    
    # 这里的逻辑是：如果 蓝色分量 > 红色分量，或者 B > 200 且 R < 240，倾向于认为是蓝色卡片
    # 你可以根据实际打印出来的数值微调这里
    if b > r and b > 200:
        return True
    
    # 额外的蓝色判断：如果背景偏冷色调
    if b > 220 and r < 230:
        return True

    return False

def already_exists(member_id: str, date_str: str) -> bool:
    """
    判断是否已存在。建议同时校验 会员号 和 预约时间，防止同一人多次预约被漏掉。
    """
    if not os.path.exists(EXCEL_PATH):
        return False
    df = pd.read_excel(EXCEL_PATH)
    
    # 简单校验会员号，如果需要更精准，可以加 try-except 校验日期
    if "会员号" in df.columns:
        return member_id in df["会员号"].astype(str).values
    return False

def save_to_excel(row: dict):
    df_new = pd.DataFrame([row])
    if os.path.exists(EXCEL_PATH):
        df_old = pd.read_excel(EXCEL_PATH)
        # 转换会员号为字符串，防止科学计数法
        if "会员号" in df_old.columns:
            df_old["会员号"] = df_old["会员号"].astype(str)
        df = pd.concat([df_old, df_new], ignore_index=True)
    else:
        df = df_new
    
    df.to_excel(EXCEL_PATH, index=False)
    print(f"✅ [保存成功] 客户: {row.get('姓名', '未知')} | 项目: {row.get('项目', '未知')}")

# ---------------- 页面行为 ----------------

def login(page):
    print("正在登录...")
    page.goto(URL, wait_until="networkidle", timeout=60000)
    
    # 填写登录表单
    page.locator("input[type='text']").nth(0).fill(COMPANY)
    page.locator("input[type='text']").nth(1).fill(USERNAME)
    page.locator("input[type='password']").fill(PASSWORD)

    page.get_by_role("button", name="登 录").click()
    page.wait_for_url("**/dashboard_v2", timeout=200000) # 等待跳转
    print("登录成功！")

def goto_appointment_center(page):
    print("正在跳转到预约中心...")
    # 1. 点击左侧 "预约" (确保侧边栏展开)
    # 使用更精准的定位，防止点到其他文字
    page.locator("li.menu-item:has-text('预约')").first.click()
    time.sleep(1)
    
    # 2. 点击 "预约中心"
    page.locator("li.submenu-item:has-text('预约中心')").click()
    
    # 3. 等待日历视图加载完毕
    # 观察图1，日历中包含 "fc-view" 或 "appointment-content" 等元素
    page.wait_for_selector(".fc-view-container", timeout=20000)
    time.sleep(3) # 额外等待数据渲染
    print("已进入预约视图。")

def extract_detail_from_modal(page) -> dict:
    """从弹出的详情框（图3）中提取数据"""
    data = {}
    
    # 等待弹窗内容加载
    modal = page.locator(".ant-modal-content, .detail-modal").first
    modal.wait_for(timeout=5000)

    # 1. 提取头部基础信息 (图3上半部分)
    # 姓名通常是最大的标题 h3, h4 或 class="name"
    try:
        # 尝试寻找包含性别图标旁边的名字
        name_loc = page.locator("div.header-info .name, h4, .title").first
        data["姓名"] = name_loc.inner_text().strip()
    except:
        data["姓名"] = "未知"

    try:
        # 会员号通常在皇冠图标或者 ID 图标旁边
        # 假设它是纯数字或特定格式
        text_content = page.locator(".header-info").inner_text()
        # 简单粗暴正则提取 ID (假设是 10位以上数字)
        id_match = re.search(r"\d{8,}", text_content)
        data["会员号"] = id_match.group(0) if id_match else "未知"
    except:
        data["会员号"] = "未知"

    # 2. 提取列表详情 (图3下半部分: 预约时间, 医生, 项目等)
    # 遍历所有的 label 和 content
    items = page.locator("div.detail-item, div.row-item") # 需要根据实际情况调整class
    
    # 如果找不到特定class，尝试遍历所有包含冒号的行
    text_lines = page.locator(".ant-modal-body").inner_text().split('\n')
    for line in text_lines:
        if "：" in line:
            parts = line.split("：", 1)
            if len(parts) == 2:
                key = parts[0].strip()
                val = parts[1].strip()
                data[key] = val

    # 补充提取时间戳，作为去重依据
    data["抓取时间"] = time.strftime("%Y-%m-%d %H:%M:%S")
    
    return data

def process_appointments(page):
    """
    核心逻辑：
    1. 遍历日历上的所有卡片
    2. 判断背景色：蓝色 -> 点击; 黄色 -> 跳过
    3. 点击后提取 -> 关闭弹窗 -> 下一个
    """
    
    # 定位日历中的所有预约块
    # 根据图1，卡片通常有 class "fc-event" 或 "appointment-card"
    card_selector = "a.fc-day-grid-event, div.event-item" 
    
    # 等待至少一个卡片出现，或者超时（可能当天无预约）
    try:
        page.wait_for_selector(card_selector, timeout=10000)
    except:
        print("未检测到预约卡片，结束。")
        return

    cards = page.locator(card_selector)
    count = cards.count()
    print(f"当前视图共检测到 {count} 个预约卡片。")

    for i in range(count):
        # 重新获取卡片句柄，防止页面DOM刷新导致 stale element
        card = cards.nth(i)
        
        # ---------------- 关键步骤：颜色识别 ----------------
        # 获取计算后的背景颜色
        bg_color = card.evaluate("el => window.getComputedStyle(el).backgroundColor")
        
        # 调试输出颜色，方便你微调
        # print(f"卡片 {i} 颜色: {bg_color}") 
        
        if not is_blue_card(bg_color):
            # 如果不是蓝色（即黄色/待确认），跳过
            # print(f"卡片 {i} 是黄色/未到店，跳过。")
            continue
        
        print(f"⚡ 发现【到店】卡片 (索引 {i})，准备处理...")

        # ---------------- 点击与提取 ----------------
        try:
            card.click()
            
            # 等待弹窗出现 (图3)
            page.wait_for_selector("text=预约概览", timeout=5000)
            
            # 提取数据
            data = extract_detail_from_modal(page)
            
            # 保存 (带去重)
            if already_exists(data.get("会员号", ""), data.get("预约时间", "")):
                print(f"   -> 跳过: {data.get('姓名')} (已存在)")
            else:
                save_to_excel(data)
                
            # ---------------- 关键步骤：关闭弹窗 ----------------
            # 不用 go_back()，而是按 ESC 关闭弹窗，速度快且稳定
            page.keyboard.press("Escape")
            time.sleep(0.5) # 等待动画消失
            
        except Exception as e:
            print(f"   -> 处理卡片 {i} 时出错: {e}")
            # 如果出错，尝试按 ESC 恢复状态，防止阻塞后续
            page.keyboard.press("Escape")

# ---------------- 主程序 ----------------

def main():
    with sync_playwright() as p:
        browser = p.chromium.launch(
            headless=False, # 设为 True 可后台运行
            args=["--start-maximized"] # 最大化窗口，方便查看日历
        )
        context = browser.new_context(no_viewport=True)
        page = context.new_page()

        try:
            login(page)
            goto_appointment_center(page)
            process_appointments(page)
            print("所有卡片处理完毕！")
        except Exception as e:
            print(f"运行出错: {e}")
        finally:
            time.sleep(3) # 观察一下结果再关闭
            browser.close()

if __name__ == "__main__":
    main()