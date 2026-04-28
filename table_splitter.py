import pandas as pd
import re
import sys

def normalize_drug_name(name):
    """标准化药物名称"""
    if not name: return ""
    name = str(name).replace('\n', ' ')
    usages = ['口服', '皮下注射', '肌肉注射', '静脉注射']
    for usage in usages:
        name = re.sub(rf'{usage}[,\s]*', '', name)
    # 剔除日期括号 (2023-08-07)
    name = re.sub(r'[\(（]\d{4}-\d{2}-\d{2}[\)）]', '', name)
    return name.strip()

def parse_medicine_column(cell_content):
    """分段解析：处理粘连、换行及截断"""
    if pd.isna(cell_content): return {}

    text = str(cell_content).replace('，', ',').replace('×', '*').replace('✕', '*')
    # 在“天”字后插入标记以强制切分
    temp_text = re.sub(r'(\d+天)', r'\1|||', text)
    parts = re.split(r'\|\|\||[,，\n]', temp_text)
    
    parsed_data = {}
    item_pattern = re.compile(
        r'(?P<drug>.+?)\s*(?P<dose>\d+\.?\d*[a-zA-Z]*)\s*(?:qd|oncem|onc|once|bid|tid|\s)*'
        r'(?:\*\s*(?P<days>\d+))?',
        re.IGNORECASE | re.DOTALL
    )

    for part in parts:
        part = part.strip()
        if not part: continue
        
        m = item_pattern.search(part)
        if m:
            drug_raw = m.group('drug').strip()
            # 清理前导编号（如 133.）
            drug_raw = re.sub(r'^[.\d\s,]+', '', drug_raw)
            dose = m.group('dose').lower()
            days_match = m.group('days')
            days = int(days_match) if days_match else 0
            
            drug_clean = normalize_drug_name(drug_raw)
            if not drug_clean: drug_clean = "未知药物"
            key = f"{drug_clean} {dose}"
            
            if days_match is None:
                display_part = part.replace('\n', ' ')
                print(f"  [截断保留] 识别到剂量但无周期: '{display_part}' -> 记 0 天")

            parsed_data[key] = parsed_data.get(key, 0) + days
        else:
            display_fail = part.replace('\n', ' ')
            if display_fail:
                print(f"  [丢弃] 无法识别有效剂量信息: '{display_fail}'")

    return parsed_data

def main(input_path, output_path):
    print(f">>> 任务启动。读取文件: {input_path}")
    try:
        # 读取源文件
        df = pd.read_excel(input_path)
    except Exception as e:
        print(f"错误: {e}")
        return

    source_cols = list(df.columns) # 记录源文件所有列名
    new_rows = []
    global_drug_columns = set()

    for index, row in df.iterrows():
        # 获取第一列（编号）和第二列（药物内容）的值
        p_id = row.iloc[0]
        medicine_cell = row.iloc[1]
        
        # 规则3：编号为空停止
        if pd.isna(p_id):
            print(f"\n[停止] 检测到第 {index+2} 行第一列为空，终止解析。")
            break
            
        print(f"\n>>> 处理编号: {p_id}")
        medicine_info = parse_medicine_column(medicine_cell)
        
        # 构造新行数据：保留源行所有内容
        new_row = row.to_dict()
        
        for drug_key, days in medicine_info.items():
            if drug_key not in global_drug_columns:
                print(f"  [新增列] '{drug_key}'")
                global_drug_columns.add(drug_key)
            new_row[drug_key] = days
            
        new_rows.append(new_row)

    # 排序新增药物列
    drug_cols_sorted = sorted(list(global_drug_columns))
    
    print("\n" + "="*60 + "\n解析摘要 (保存前确认):\n" + "="*60)
    for i, name in enumerate(drug_cols_sorted, 1):
        print(f"{i:02d}. {name}")
    print("="*60)

    # 生成 DataFrame，未使用药品的位置将默认为空 (NaN)
    result_df = pd.DataFrame(new_rows)
    
    # 确定列顺序：源文件所有原始列 + 追加的药物列
    final_order = source_cols + drug_cols_sorted
    
    # 过滤掉在结果中不存在的列（防止因空行终止导致的问题）
    final_order = [c for c in final_order if c in result_df.columns]
    
    # 保存，不填充 0，Excel 中将显示为空白单元格
    result_df[final_order].to_excel(output_path, index=False)
    print(f"结果已成功输出至: {output_path}")

if __name__ == "__main__":
    if len(sys.argv) != 3:
        print("用法: python table_splitter.py <输入.xlsx> <输出.xlsx>")
    else:
        main(sys.argv[1], sys.argv[2])
