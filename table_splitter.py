import pandas as pd
import re
import sys
import os

def parse_medication(cell_text):
    """
    解析药物字符串。
    逻辑：
    1. 严格按换行符和逗号切分，防止行间粘连。
    2. 药物名称完整保留（包含括号注释），将其作为独立列。
    """
    if pd.isna(cell_text) or not str(cell_text).strip():
        return {}
    
    # 1. 预处理：统一乘号，将换行符和逗号都视为分隔符
    text = str(cell_text).replace('×', '*').replace('，', ',').replace('\r', '\n')
    # 按逗号或换行符分割，并去除每个片段前后的空白
    parts = [p.strip() for p in re.split(r'[,\n]+', text) if p.strip()]
    
    med_map = {}
    
    # 2. 核心正则：
    # Group 1: ([^0-9*]+) -> 药物名称（保留所有非数字和非乘号字符，包括括号）
    # Group 2: (\d+(?:\.\d+)?\s*[a-zA-Z]+\s*(?:qd|bid|oncem|onc|tid|qn|QD|BID|TID)?) -> 剂量单位频率
    # Group 3: \*(\d+) -> 周期天数
    pattern = r"([^0-9*]+)(\d+(?:\.\d+)?\s*[a-zA-Z]+\s*(?:qd|bid|oncem|onc|tid|qn|QD|BID|TID)?)\s*[*]\s*(\d+)"
    
    for part in parts:
        match = re.search(pattern, part, re.IGNORECASE)
        if match:
            # 完整保留名称，包括括号内的注释
            name = match.group(1).strip()
            dosage = match.group(2).strip()
            days = int(match.group(3))
            
            # 药物名和剂量共同构成独立列名
            key = f"{name} {dosage}"
            # 同一行内的相同药物剂量组合，天数累加
            med_map[key] = med_map.get(key, 0) + days
            
    return med_map

def process_excel(input_file, output_file):
    if not os.path.exists(input_file):
        print(f"错误：找不到文件 {input_file}")
        return

    print(f"正在读取文件: {input_file}")
    try:
        df = pd.read_excel(input_file)
        orig_cols = df.columns.tolist()
        
        # 定义列位置
        id_col = orig_cols[0]    # 第一列：编号
        med_col = orig_cols[1]   # 第二列：药物汇总
        other_cols = orig_cols[2:] # 第三列及之后
        
        all_processed_data = []
        
        print("开始解析数据...")
        for _, row in df.iterrows():
            med_dict = parse_medication(row[med_col])
            
            # 构造新行：编号 + 其他原始列 + 拆分后的药物列
            record = {col: row[col] for col in orig_cols}
            record.update(med_dict)
            all_processed_data.append(record)
            
        result_df = pd.DataFrame(all_processed_data)
        
        # 提取新生成的药物列并排序
        new_med_cols = sorted([c for c in result_df.columns if c not in orig_cols])
        
        # 重新排序列顺序
        final_column_order = [id_col] + other_cols + new_med_cols
        result_df = result_df[final_column_order]
        
        # 缺失值填 0
        result_df[new_med_cols] = result_df[new_med_cols].fillna(0)
        
        result_df.to_excel(output_file, index=False)
        print(f"处理成功！输出至: {output_file}")
        print(f"共生成了 {len(new_med_cols)} 个独立的药物剂量列。")

    except Exception as e:
        print(f"运行失败: {e}")
        import traceback
        traceback.print_exc()

if __name__ == "__main__":
    if len(sys.argv) < 3:
        print("使用说明: python3 test.py <输入文件.xlsx> <输出文件.xlsx>")
    else:
        process_excel(sys.argv[1], sys.argv[2])
