import pandas as pd
import re
import sys

def normalize_drug_name(name):
    """
    标准化药物名称：移除用法前缀和日期括号，保留英文注释
    """
    usage_keywords = ['口服', '皮下注射', '肌肉注射', '静脉注射']
    for keyword in usage_keywords:
        name = name.replace(keyword, "")
    
    # 移除日期括号 (如 2023-08-07)，保留英文注释括号 (如 HCG)
    name = re.sub(r'[（\(]\d{4}-\d{2}-\d{2}[）\)]', '', name)
    return name.strip()

def parse_medicine_column(cell_content):
    """
    解析第二列内容，返回 { "药物+剂量": 天数 }
    """
    if pd.isna(cell_content):
        return {}

    # 处理分隔符：兼容逗号和换行符
    raw_entries = re.split(r'[,\n]+', str(cell_content))
    parsed_data = {}
    
    # 正则：捕获药物、剂量、忽略qd等、捕获天数
    pattern = re.compile(
        r'^(?P<drug>.*?)\s*'
        r'(?P<dose>\d+\.?\d*[a-zA-Z]+)\s*'
        r'(?:qd|oncem|onc|once|bid)?\s*'
        r'[*×]\s*(?P<days>\d+)',
        re.IGNORECASE
    )

    for entry in raw_entries:
        entry = entry.strip()
        if not entry: continue
        
        match = pattern.search(entry)
        if match:
            drug_raw = match.group('drug')
            dose = match.group('dose').lower()
            days = int(match.group('days'))
            
            drug_clean = normalize_drug_name(drug_raw)
            if drug_raw != drug_clean:
                print(f"  [清洗] 药物名: '{drug_raw}' -> '{drug_clean}'")

            key = f"{drug_clean} {dose}"
            parsed_data[key] = parsed_data.get(key, 0) + days
        else:
            print(f"  [丢弃/截断] 无法识别条目: '{entry}'")
            
    return parsed_data

def main(input_file, output_file):
    print(f"正在读取: {input_file}...")
    df = pd.read_excel(input_file)
    
    new_rows = []
    # 用于统计总药物种类的计数器
    total_drugs_counter = {}

    for index, row in df.iterrows():
        patient_id = row.iloc[0]
        print(f"处理第 {index+1} 行 (编号: {patient_id})")
        
        medicine_dict = parse_medicine_column(row.iloc[1])
        
        # 基础数据
        new_row_dict = row.to_dict()
        original_col2_name = df.columns[1]
        del new_row_dict[original_col2_name]
        
        # 填充新列并记录全局药物
        for drug_key, days in medicine_dict.items():
            new_row_dict[drug_key] = days
            total_drugs_counter[drug_key] = total_drugs_counter.get(drug_key, 0) + 1
            if total_drugs_counter[drug_key] > 1:
                print(f"  [复用] 药物列 '{drug_key}' 已存在")

        new_rows.append(new_row_dict)

    # 输出结果文件
    output_df = pd.DataFrame(new_rows)
    # 保持原表其他列在前，新药物列在后
    cols = [c for c in df.columns if c != df.columns[1]] + sorted(list(total_drugs_counter.keys()))
    output_df[cols].to_excel(output_file, index=False)

    # --- 输出总的药物种类统计 ---
    print("\n" + "="*30)
    print(f"任务完成！总计识别出 {len(total_drugs_counter)} 种药物(含剂量)组合：")
    for i, (drug_name, count) in enumerate(sorted(total_drugs_counter.items()), 1):
        print(f"{i}. {drug_name} (在 {count} 行中出现)")
    print("="*30)
    print(f"结果已保存至: {output_file}")

if __name__ == "__main__":
    if len(sys.argv) != 3:
        print("用法: python table_splitter.py <输入Excel> <输出Excel>")
    else:
        main(sys.argv[1], sys.argv[2])
