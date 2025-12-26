import pandas as pd
import numpy as np

def process_data(file_path):
    print(f"正在读取文件: {file_path} ...")
    
    try:
        # 读取 Excel 文件
        df = pd.read_excel(file_path)
        
        # 1. 清洗表头：去除表头可能存在的空格，防止 '消费日期 ' 这种匹配不到的情况
        df.columns = df.columns.str.strip()
        print(f"成功读取，包含列名: {list(df.columns)}")
        
    except Exception as e:
        print(f"读取文件失败: {e}")
        return

    # --- 配置区域 ---
    # 根据你的截图，示例产品为"其他"，我已将其加入列表。
    # 请在此处补全所有合规的9个产品名称
    VALID_PRODUCTS = [
        "其他", "保妥适单次", "乔雅登", "酷塑", "标签5", 
        "标签6", "标签7", "标签8", "标签9"
    ]
    
    # --- 校验逻辑函数 ---
    def validate_row(row):
        errors = []
        
        # 1. 消费日期 (必填, 格式转换)
        # 你的源数据带时间(18:46:31)，这里主要检查能否转换为日期
        date_val = row.get('消费日期')
        if pd.isna(date_val):
            errors.append("消费日期为空")
        else:
            try:
                pd.to_datetime(date_val)
            except:
                errors.append("消费日期格式错误")

        # 2. 业绩金额 (非必填, 范围检查)
        amount = row.get('业绩金额')
        if not pd.isna(amount):
            try:
                amt_num = float(amount)
                if not (-1000000 <= amt_num <= 1000000):
                    errors.append("业绩金额超出范围 (-100万 到 +100万)")
            except ValueError:
                errors.append("业绩金额必须是数字")

        # 3. 客户卡号 (必填, 长度<=50)
        # 注意：Excel读取长数字可能会变成科学计数法或数字类型，需强制转字符串
        card_no = str(row.get('客户卡号', ''))
        # 如果是 NaN 或者转成字符串是 'nan'
        if pd.isna(row.get('客户卡号')) or card_no.lower() == 'nan' or card_no.strip() == '':
            errors.append("客户卡号为空")
        elif len(card_no) > 50:
            errors.append(f"客户卡号长度超过50位")

        # 4. 渠道来源 (非必填, 长度<=50)
        source = row.get('渠道来源')
        if not pd.isna(source):
            if len(str(source)) > 50:
                errors.append("渠道来源长度超过50位")

        # 5. 咨询师 (非必填, 长度<=10)
        consultant = row.get('咨询师')
        if not pd.isna(consultant):
            if len(str(consultant)) > 10:
                errors.append("咨询师名称长度超过10位")

        # 6. 消费产品 (必填, 必须在白名单内)
        product = row.get('消费产品')
        if pd.isna(product) or str(product).strip() == '':
            errors.append("消费产品为空")
        elif str(product).strip() not in VALID_PRODUCTS:
            errors.append(f"产品名称不合规")

        return "; ".join(errors)

    # --- 执行校验 ---
    print("正在校验数据...")
    df['数据校验结果'] = df.apply(validate_row, axis=1)

    # --- 数据清洗与格式化 (Formatting) ---
    
    # 1. 日期格式化：无论原数据是 "2025-11-30 18:46:31" 还是其他，统一转为 "yyyy/mm/dd"
    # errors='coerce' 会把无法转换的变成 NaT，避免报错
    df['消费日期'] = pd.to_datetime(df['消费日期'], errors='coerce').dt.strftime('%Y/%m/%d')
    
    # 2. 金额格式化：保留两位小数
    df['业绩金额'] = pd.to_numeric(df['业绩金额'], errors='coerce').round(2)

    # 3. 客户卡号：防止变成 2.50822E+11 这种形式，去掉 .0
    def format_card(x):
        if pd.isna(x): return ""
        s = str(x)
        if s.endswith('.0'): return s[:-2]
        return s
    df['客户卡号'] = df['客户卡号'].apply(format_card)

    # --- 输出统计 ---
    valid_count = len(df[df['数据校验结果'] == ""])
    invalid_count = len(df) - valid_count
    print(f"校验完成: 通过 {valid_count} 行, 失败 {invalid_count} 行")

    # --- 保存结果 ---
    output_filename = '处理结果_a.xlsx'
    
    try:
        with pd.ExcelWriter(output_filename, engine='xlsxwriter') as writer:
            df.to_excel(writer, index=False, sheet_name='处理结果')
            
            workbook = writer.book
            worksheet = writer.sheets['处理结果']
            
            # 设置列宽，方便阅读
            worksheet.set_column('A:A', 15) # 消费日期
            worksheet.set_column('B:B', 12) # 业绩金额
            worksheet.set_column('C:C', 20) # 客户卡号
            worksheet.set_column('G:G', 40) # 校验结果列(假设在G列)

            # 标记错误的行：如果有错误，最后一列标红
            red_format = workbook.add_format({'bg_color': '#FFC7CE', 'font_color': '#9C0006'})
            # 如果想让最后一列文字变红，ExcelWriter需要更复杂的条件格式，
            # 这里我们简单一点：告诉用户直接看最后一列。

        print(f"处理完毕！结果已保存至: {output_filename}")
        
    except Exception as e:
        print(f"保存文件失败，请检查文件是否被占用: {e}")

if __name__ == "__main__":
    # 请确保你的文件名为 a.xlsx
    process_data('a.xlsx')