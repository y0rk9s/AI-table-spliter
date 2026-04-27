import pandas as pd
import re
import sys
import os

def clean_drug_name(raw_name):
    """
    清洗药物名称：移除日期、用法（皮下、肌肉等），保留英文注释。
    """
    # 1. 移除日期格式如 (2023-08-07) 或 2023-08-07
    name = re.sub(r'\(?\d{4}-\d{2}-\d{2}\)?', '', raw_name)
    # 2. 移除用法信息及分隔符
    tags = ["皮下注射", "肌肉注射", "口服", "注射", "外用"]
    for tag in tags:
        name = name.replace(tag, "")
    # 3. 清理多余标点和空格
    name = name.strip(",.，。 /")
    return name

def parse_drug_column(cell_content):
    """
    解析第二列内容，处理换行、缺失逗号、截断及剂量合并。
    """
    if pd.isna(cell_content):
        return {}

    content = str(cell_content).replace('×', '*')
    # 1. 统一分隔符：由于“天”是基础格式的逻辑终点，且存在换行符代替逗号的情况
    # 将换行符转换为逗号，方便后续分割
    content = content.replace('\n', ',')
    
    # 2. 切分条目
    items = [i.strip() for i in re.split(r'[,，]', content) if i.strip()]
    
    row_results = {}
    
    # 正则逻辑：
    # (.*?) 匹配药物名称
    # (\d+(?:\.\d+)?[a-zA-Z]+) 匹配剂量（数值+单位）
    # (?:.*?) 非捕获组跳过用法（如qd, bid）
    # (?:\*(\d+)天)? 捕获天数（可选，应对截断）
    pattern = re.compile(r'^(.*?)\s*(\d+(?:\.\d+)?[a-zA-Z]+)(?:.*?(?:\*(\d+)天))?', re.IGNORECASE)

    for item in items:
        match = pattern.search(item)
        if match:
            raw_name, dose, days = match.groups()
            
            # 清理药物名
            drug_name = clean_drug_name(raw_name)
            if not drug_name: continue
            
            # 统一剂量格式（不含QD等频率词）
            dose_clean = dose.upper()
            column_key = f"{drug_name} {dose_clean}"
            
            # 检查是否有天数（应对截断情况）
            if days:
                day_val = int(days)
                if drug_name != raw_name.strip():
                    print(f"  [处理] 原始词条: '{item}' -> 识别为: '{column_key}'")
                
                # 同行内相同药量天数相加
                row_results[column_key] = row_results.get(column_key, 0) + day_val
            else:
                print(f"  [异常] 丢弃截断条目: '{item}' (无法提取有效天数)")
        else:
            print(f"  [异常] 格式不匹配，跳过: '{item}'")
                
    return row_results

def main(input_path, output_path):
    if not os.path.exists(input_path):
        print(f"错误: 找不到输入文件 {input_path}")
        return

    # 读取Excel，不设置表头名称以兼容不同标题，假设第一、二列位置固定
    df = pd.read_excel(input_path)
    
    all_rows_data = []
    seen_drug_cols = [] # 记录药物列出现的顺序
    
    print(f"--- 开始处理: {input_path} ---")

    for idx, row in df.iterrows():
        p_id = row.iloc[0]
        drug_cell = row.iloc[1]
        other_data = row.iloc[2:].to_dict() # 其它相关信息列
        
        print(f"第 {idx+1} 行 (ID: {p_id}) 处理中...")
        
        parsed_drugs = parse_drug_column(drug_cell)
        
        # 整合数据
        new_entry = {"编号": p_id}
        new_entry.update(other_data)
        
        for drug_key, days in parsed_drugs.items():
            if drug_key not in seen_drug_cols:
                print(f"  [新种类] 发现新药物配置: {drug_key}")
                seen_drug_cols.append(drug_key)
            else:
                # 输出复用信息（可选）
                pass
            new_entry[drug_key] = days
            
        all_rows_data.append(new_entry)

    # 构造新的DataFrame
    # 确定列顺序：编号 + 原有其它列 + 按发现顺序排序的药物列
    other_cols = list(df.columns[2:])
    final_cols = ["编号"] + other_cols + seen_drug_cols
    
    output_df = pd.DataFrame(all_rows_data, columns=final_cols)
    
    # 填充缺失值为0或空（这里填充为空，代表该病人未服用此药）
    output_df.to_excel(output_path, index=False)
    
    print(f"\n--- 处理完成 ---")
    print(f"总计识别药物种类: {len(seen_drug_cols)} 种")
    print(f"结果已保存至: {output_path}")

if __name__ == "__main__":
    if len(sys.argv) != 3:
        print("使用方法: python table_splitter.py <输入路径> <输出路径>")
    else:
        main(sys.argv[1], sys.argv[2])
