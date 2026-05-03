import sys
import os
import re
import pandas as pd
from openpyxl import Workbook

# --- 配置区域 ---
DEBUG_MODE = True  # 设置为 False 可关闭详细日志，仅显示进度

def parse_args():
    """解析命令行参数"""
    if len(sys.argv) < 2:
        print("❌ 错误: 请提供输入文件路径作为第一个参数。")
        print("用法: python table_splitter.py <输入文件> [输出文件]")
        sys.exit(1)
    
    input_file = sys.argv[1]
    output_file = sys.argv[2] if len(sys.argv) > 2 else f"{os.path.splitext(input_file)[0]}-o.xlsx"
    
    return input_file, output_file

def clean_data(text):
    """清洗数据：去除换行符、首尾空格，模拟Excel CLEAN"""
    if not isinstance(text, str):
        return ""
    # 去除换行符和回车
    text = re.sub(r'[\n\r]', ' ', text)
    # 去除多余空白
    text = ' '.join(text.split())
    return text

def normalize_unit(unit_str):
    """单位标准化"""
    if not unit_str:
        return ""
    unit = unit_str.lower().strip()
    if unit in ['i', 'u', 'iu']:
        return 'mg' 
    return unit

def parse_dosage_part(part_text, debug_logs):
    """
    解析单份药物的 [药物名称 剂量 频度]
    """
    part_text = part_text.strip()
    if not part_text:
        return None, None

    # 1. 提取频度 (*数字天)
    freq_match = re.search(r'[*×]\s*(\d+)\s*天', part_text)
    days = int(freq_match.group(1)) if freq_match else 0
    
    text_no_freq = re.sub(r'[*×]\s*\d+\s*天', '', part_text).strip()
    
    if not text_no_freq:
        debug_logs.append(f"  ⚠️ 无法识别天数，丢弃片段: '{part_text}'")
        return None, None

    # 2. 提取剂量数值和单位
    dosage_match = re.search(r'(\d+\.?\d*)\s*(mg|iu|i|u|片)?', text_no_freq)
    
    dosage_val = 0.0
    unit = ""
    multiplier = 1 
    
    if dosage_match:
        dosage_val = float(dosage_match.group(1))
        unit = normalize_unit(dosage_match.group(2))
        
        after_dosage = text_no_freq[dosage_match.end():].lower().strip()
        
        if re.search(r'\bbid\b', after_dosage):
            multiplier = 2
        elif re.search(r'\btid\b', after_dosage):
            multiplier = 3
            
        final_dosage = dosage_val * multiplier
        
        # 3. 提取药物名称
        drug_name = text_no_freq[:dosage_match.start()].strip()
        
        if not drug_name:
            debug_logs.append(f"  ⚠️ 无法识别药名，丢弃片段: '{part_text}'")
            return None, None
            
        drug_key = f"{drug_name}_{final_dosage}_{unit}"
        
        # --- 调试信息生成 ---
        info = f"解析成功: [{drug_name}] 剂量[{final_dosage}{unit}] 天数[{days}]"
        if multiplier > 1:
            info += f" (频度倍率x{multiplier})"
        debug_logs.append(f"  ✅ {info}")
            
        return drug_key, days
    else:
        debug_logs.append(f"  ⚠️ 未找到剂量数值，丢弃片段: '{part_text}'")
        return None, None

def split_and_process_row(row_data, row_idx):
    """
    处理单行数据
    """
    # 检查第一列编号是否为空
    if pd.isna(row_data.iloc[0]) or str(row_data.iloc[0]).strip() == "":
        return None, None

    col2_content = row_data.iloc[1]
    if pd.isna(col2_content):
        print(f"📋 [行 {row_idx}] 第2列为空，跳过。")
        return "无药物信息", {}

    raw_text = str(col2_content)
    clean_text = clean_data(raw_text)
    
    row_logs = [] # 用于收集本行的调试日志
    drug_map = {} 

    # 预处理文本：连续逗号变一个，去除首尾逗号
    clean_text = re.sub(r',+', ',', clean_text)
    if clean_text.startswith(','): clean_text = clean_text[1:]
    if clean_text.endswith(','): clean_text = clean_text[:-1]
    
    segments = clean_text.split(',')
    
    valid_segments_count = 0
    
    for seg in segments:
        seg = seg.strip()
        if not seg: continue
        
        # 检查是否包含频度特征
        if re.search(r'[*×]\s*\d+\s*天', seg):
            drug_key, days = parse_dosage_part(seg, row_logs)
            if drug_key:
                valid_segments_count += 1
                if drug_key in drug_map:
                    drug_map[drug_key] += days
                    row_logs.append(f"  ➕ 药物已存在，天数累加 -> {drug_map[drug_key]}天")
                else:
                    drug_map[drug_key] = days
        else:
            # 没有频度特征，视为截断或无效
            row_logs.append(f"  ❌ 丢弃无效/截断片段: '{seg}'")

    # 生成处理信息字符串（写入Excel用）
    info_str = f"处理{valid_segments_count}种药物"
    if len(row_logs) > 0:
        # 如果日志太多，只取前几个摘要，避免Excel单元格爆炸，但控制台会全打出来
        info_str += "; " + "; ".join([l for l in row_logs if "✅" in l or "➕" in l][:3]) 

    # --- 控制台详细输出 ---
    print("-" * 30)
    print(f"📋 [行 {row_idx}] 正在处理编号: {row_data.iloc[0]}")
    if DEBUG_MODE:
        print(f"   原始内容: {raw_text[:50]}...") # 截取前50字符防止刷屏
        for log in row_logs:
            print(log)
    print(f"   最终结果: 共 {len(drug_map)} 种药物组合")
    
    return info_str, drug_map

def main():
    input_file, output_file = parse_args()
    
    print(f"🚀 启动脚本...")
    print(f"📂 输入文件: {input_file}")
    print(f"💾 输出文件: {output_file}")
    
    try:
        df = pd.read_excel(input_file)
    except Exception as e:
        print(f"❌ 读取文件失败: {e}")
        sys.exit(1)

    print(f"📊 加载完成，共 {len(df)} 行。开始逐行解析...\n")

    all_drug_keys = set()
    processed_rows = []

    for index, row in df.iterrows():
        info, drug_map = split_and_process_row(row, index + 2) # 行号从2开始（考虑表头）
        
        if info is None: 
            print(f"ℹ️ [行 {index+2}] 编号为空，停止处理后续行。")
            break
            
        processed_rows.append((row, info, drug_map))
        if drug_map:
            all_drug_keys.update(drug_map.keys())

    # 排序药物列名
    sorted_drug_keys = sorted(list(all_drug_keys))
    
    print("\n" + "="*40)
    print(f"✨ 数据处理结束。发现 {len(sorted_drug_keys)} 种 Unique 药物组合。")
    print(f"💾 正在保存文件 (稀疏化模式)...")

    # 使用 write_only 模式优化性能
    wb = Workbook(write_only=True)
    ws = wb.create_sheet()

    # 写入表头
    header = list(df.columns) + ["处理过程信息"] + sorted_drug_keys
    ws.append(header)

    # 写入数据行
    for row, info, drug_map in processed_rows:
        new_row = []
        for val in row:
            new_row.append(val)
        
        new_row.append(info)
        
        for key in sorted_drug_keys:
            if key in drug_map:
                new_row.append(drug_map[key])
            else:
                new_row.append(None)
                
        ws.append(new_row)

    wb.save(output_file)
    print(f"✅ 成功！文件已保存至: {output_file}")

if __name__ == "__main__":
    main()
