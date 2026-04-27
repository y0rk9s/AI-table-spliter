import pandas as pd
import re
import sys

def normalize_drug_name(name):
    """
    标准化药物名称：
    1. 剔除用法信息及随后的空格或逗号
    2. 剔除带日期的括号，保留英文注释括号
    """
    if not name: return ""
    # 替换掉所有换行符为空格，确保名称连贯
    name = name.replace('\n', ' ')
    
    # 剔除用法
    usages = ['口服', '皮下注射', '肌肉注射', '静脉注射']
    for usage in usages:
        name = re.sub(rf'{usage}[,\s]*', '', name)
    
    # 剔除日期括号 (如 2023-08-07)，保留其他括号 (如 HCG)
    name = re.sub(r'[\(（]\d{4}-\d{2}-\d{2}[\)）]', '', name)
    
    return name.strip()

def parse_medicine_column(cell_content):
    """
    利用正则表达式捕获 [药名 剂量 频度] 结构。
    即使药名和剂量间有换行，或条目间没逗号，只要符合“剂量+*+天数”即可识别。
    """
    if pd.isna(cell_content): return {}

    # 预处理：统一中文符号
    text = str(cell_content).replace('，', ',').replace('×', '*')
    
    # 核心正则表达式：
    # (?P<drug>.*?) -> 药物名称（非贪婪匹配）
    # (?P<dose>\d+\.?\d*[a-zA-Z]+) -> 剂量（数字+单位），例如 150iu, 0.05mg
    # (?:qd|oncem|onc|once|bid|\s)* -> 频次单词或空格（忽略）
    # \*\s*(?P<days>\d+)\s*天 -> 周期（锚点：*数字天）
    pattern = re.compile(
        r'(?P<drug>.*?)\s*'
        r'(?P<dose>\d+\.?\d*[a-zA-Z]+)\s*'
        r'(?:qd|oncem|onc|once|bid|\s)*'
        r'\*\s*(?P<days>\d+)\s*天',
        re.IGNORECASE | re.DOTALL
    )

    parsed_data = {}
    # 使用 finditer 逐个寻找符合“剂量+天数”结构的有效条目
    for m in pattern.finditer(text):
        drug_raw = m.group('drug').strip()
        # 清除开头可能残余的逗号或换行
        drug_raw = re.sub(r'^[,\s\n]+', '', drug_raw)
        
        dose = m.group('dose').lower()
        days = int(m.group('days'))
        
        # 标准化处理
        drug_clean = normalize_drug_name(drug_raw)
        if not drug_clean: drug_clean = "未知药物"
        
        if drug_raw != drug_clean:
            # 这里的 drug_raw 可能包含换行，打印时简单处理
            clean_print = drug_raw.replace('\n', ' ')
            print(f"  [清洗] 发现附带信息: '{clean_print}' -> '{drug_clean}'")

        key = f"{drug_clean} {dose}"
        
        # 同行天数累加
        if key in parsed_data:
            print(f"  [合并] 同行重复药物 '{key}'，天数累加: {parsed_data[key]} + {days}")
            parsed_data[key] += days
        else:
            parsed_data[key] = days

    return parsed_data

def main(input_path, output_path):
    print(f"读取输入文件: {input_path}")
    try:
        df = pd.read_excel(input_path)
    except Exception as e:
        print(f"读取失败: {e}")
        return

    new_rows = []
    global_drug_kinds = set()

    for index, row in df.iterrows():
        # 规则 4: 编号为空则停止处理
        if pd.isna(row.iloc[0]):
            print(f"\n[提示] 第 {index+1} 行编号为空，停止解析。")
            break
            
        p_id = row.iloc[0]
        print(f"\n>>> 正在处理病人: {p_id}")
        
        # 解析第二列
        medicine_results = parse_medicine_column(row.iloc[1])
        
        if not medicine_results:
            print(f"  [提醒] 未能识别到有效药物条目（可能数据为空或截断严重）。")

        # 构造新行，保留原表其他列
        new_row = row.to_dict()
        col2_name = df.columns[1]
        del new_row[col2_name]
        
        for drug_key, days in medicine_results.items():
            if drug_key in global_drug_kinds:
                print(f"  [复用] 匹配到已有药物种类: '{drug_key}'")
            else:
                print(f"  [新增] 发现全新药物种类: '{drug_key}'")
                global_drug_kinds.add(drug_key)
            
            new_row[drug_key] = days
            
        new_rows.append(new_row)

    # 结果整合
    result_df = pd.DataFrame(new_rows)
    
    # 排序：非药列在前，药列按字母排序在后
    other_cols = [c for i, c in enumerate(df.columns) if i != 1]
    drug_cols = sorted(list(global_drug_kinds))
    
    # 确保列存在（处理全空表情况）
    final_cols = [c for c in other_cols if c in result_df.columns] + drug_cols
    result_df = result_df[final_cols]

    result_df.to_excel(output_path, index=False)
    
    print("\n" + "="*50)
    print(f"任务完成！总计处理 {len(new_rows)} 条记录。")
    print(f"总计识别唯一药物种类: {len(global_drug_kinds)}")
    print("-" * 50)
    for i, name in enumerate(drug_cols, 1):
        print(f"{i:02d}. {name}")
    print("="*50)
    print(f"结果保存至: {output_path}")

if __name__ == "__main__":
    if len(sys.argv) != 3:
        print("用法: python table_splitter.py <输入.xlsx> <输出.xlsx>")
    else:
        main(sys.argv[1], sys.argv[2])
