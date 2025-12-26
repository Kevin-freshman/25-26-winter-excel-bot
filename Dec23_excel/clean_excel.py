import pandas as pd
import numpy as np
from datetime import datetime

def process_data(file_path):
    # 读取 Excel 文件
    try:
        df = pd.read_excel(file_path)
        print(f"成功读取文件: {file_path}, 共 {len(df)} 行数据")
    except Exception as e:
        print(f"读取文件失败: {e}")
        return

    # --- 配置区域 ---
    # 图片规则提到：消费产品必须是9个标签中的一个，且文字必须一模一样
    # 由于图片中没有列出具体9个标签，这里请你手动替换为真实的标签列表
    VALID_PRODUCTS = [
        "保妥适单次", "乔雅登", "保妥适疗程卡", "热玛吉", "超声刀", 
        "动能素", "自定义项目1", "自定义项目2", "其他项目"
    ]
    
    # 用于存储错误信息的列表
    error_logs = []

    # --- 数据处理与校验函数 ---
    def validate_row(row, index):
        errors = []
        
        # 1. 消费日期 (必填, 格式 yyyy/mm/dd)
        date_val = row.get('消费日期')
        if pd.isna(date_val):
            errors.append("消费日期为空")
        else:
            try:
                # 尝试转换为日期格式
                pd.to_datetime(date_val)
            except:
                errors.append("消费日期格式错误 (需为 yyyy/mm/dd)")

    # 2. 客户卡号 (必填, 长度<=50)
        card_no = str(row.get('客户卡号', ''))
        if pd.isna(row.get('客户卡号')) or card_no.strip() == '':
            errors.append("客户卡号为空")
        elif len(card_no) > 50:
            errors.append(f"客户卡号长度超过50位 (当前{len(card_no)}位)")

        # 3. 年龄 (非必填, 0-100)
        age = row.get('年龄 (非必填)')
        if not pd.isna(age):
            try:
                age_num = float(age)
                if not (0 <= age_num <= 100):
                    errors.append("年龄必须在 0-100 之间")
            except ValueError:
                errors.append("年龄必须是数字")

        # 4. 消费产品 (必填, 必须匹配指定标签)
        product = row.get('消费产品')
        if pd.isna(product) or str(product).strip() == '':
            errors.append("消费产品为空")
        elif str(product).strip() not in VALID_PRODUCTS:
            errors.append(f"消费产品名称不匹配 (必须属于指定9个标签)")

        # 5. 业绩金额 (非必填, -100万到+100万, 保留两位小数)
        amount = row.get('业绩金额 (非必填)')
        if not pd.isna(amount):
            try:
                amt_num = float(amount)
                if not (-1000000 <= amt_num <= 1000000):
                    errors.append("业绩金额超出范围 (-100万 到 +100万)")
            except ValueError:
                errors.append("业绩金额必须是数字")

        # 6. 咨询师 (非必填, 长度<=10)
        consultant = row.get('咨询师 (非必填)')
        if not pd.isna(consultant):
            if len(str(consultant)) > 10:
                errors.append("咨询师名称长度超过10位")

        # 7. 渠道来源 (非必填, 长度<=50)
        source = row.get('渠道来源 (非必填)')
        if not pd.isna(source):
            if len(str(source)) > 50:
                errors.append("渠道来源内容长度超过50位")

        return "; ".join(errors)

    # --- 执行校验 ---
    print("正在校验数据...")
    # 创建一个新列 '数据校验结果' 来存储每行的检查情况
    df['数据校验结果'] = df.apply(lambda row: validate_row(row, row.name), axis=1)

    # --- 数据格式化 (Formatting) ---
    # 即使数据有错，我们也尝试对合规的数据进行格式化，以便后续使用
    
    # 1. 格式化日期为字符串 'yyyy/mm/dd'
    df['消费日期'] = pd.to_datetime(df['消费日期'], errors='coerce').dt.strftime('%Y/%m/%d')
    
    # 5. 格式化金额，保留两位小数
    df['业绩金额 (非必填)'] = pd.to_numeric(df['业绩金额 (非必填)'], errors='coerce').round(2)

    # --- 输出结果 ---
    # 分离出有问题的数据和没问题的数据
    valid_rows = df[df['数据校验结果'] == ""]
    invalid_rows = df[df['数据校验结果'] != ""]

    print("-" * 30)
    print(f"校验完成。")
    print(f"通过行数: {len(valid_rows)}")
    print(f"失败行数: {len(invalid_rows)}")

    # 保存结果到新文件
    output_filename = 'processed_a.xlsx'
    
    # 使用 ExcelWriter 可以设置列宽等格式，这虽然不是必须的，但体验更好
    with pd.ExcelWriter(output_filename, engine='xlsxwriter') as writer:
        df.to_excel(writer, index=False, sheet_name='处理结果')
        
        # 获取 workbook 和 worksheet 对象
        workbook = writer.book
        worksheet = writer.sheets['处理结果']
        
        # 设置红色高亮格式用于标记错误列
        red_format = workbook.add_format({'font_color': '#9C0006', 'bg_color': '#FFC7CE'})
        
        # 如果有错误信息，可以根据条件格式化高亮 (这里简单高亮最后一列)
        if len(df) > 0:
             worksheet.set_column(len(df.columns)-1, len(df.columns)-1, 40) # 拓宽错误信息列

    print(f"结果已保存至: {output_filename}")
    if len(invalid_rows) > 0:
        print("请打开生成的 Excel 查看最后一列的具体错误原因。")

# 运行函数
if __name__ == "__main__":
    # 确保目录下有 a.xlsx 文件
    process_data('a.xlsx')