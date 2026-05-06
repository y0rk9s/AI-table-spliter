import pandas as pd
import re
import sys
import os
from openpyxl import Workbook

def normalize_drug_name(name):
    """标准化药物名称逻辑：用于Key构造，剔除用法和日期"""
    if not name: return ""
    name = str(name).replace('\n', ' ').strip()
    name = re.sub(r'达英\s*-\s*35', '##DY35##', name)
    name = re.sub(r'[\(（]\s*\d{4}[-/]\d{1,2}[-/]\d{1,2}\s*[\)）]', '', name)
    usages = ['口服', '皮下注射', '肌肉注射', '静脉注射', '塞阴道']
    for u in usages: name = re.sub(rf'{u}[,\s，]*', '', name)
    name = name.replace('##DY35##', '达英-35')
    return re.sub(r'^[,，\s\.\-]+', '', name).strip()

def parse_medicine_column(cell_content, global_cols):
    """字符级镜像映射解析逻辑"""
    if pd.isna(cell_content): return {}, [], "", ""

    # 1. CLEAN 清洗换行
    raw_text = str(cell_content).replace('\n', '').replace('\r', '')
    # 统一符号用于内部匹配逻辑
    norm_text = raw_text.replace('×', '*').replace('✕', '*').replace('x', '*').replace('，', ',')
    
    # 初始化镜像对齐列表
    aligned_chars = list(' ' * len(raw_text))
    
    # 2. 物理定位切分
    parts_meta = []
    # 寻找以“天”结尾的片段或末尾片段
    for m in re.finditer(r'[^,，]+?(\d+\s*天|$)', norm_text):
        if m.group().strip():
            parts_meta.append({'text': m.group(), 'start': m.start(), 'end': m.end()})

    parsed_data = {}
    logs = []
    
    # 3. 核心提取正则（增加各元素位置捕获）
    # 药名 + 剂量数值 + 单位 + 频次(可选) + * + 天数
    item_pattern = re.compile(
        r'(?P<drug>.+?)\s*(?P<val>\d+\.?\d*)(?P<unit>mg|iu|i|u|片|[\u4e00-\u9fa5]*)\s*'
        r'(?P<freq>qd|q3d|oncem|onc|once|w3d|bid|tid|\s)*(?:\*\s*(?P<days>\d+))', 
        re.IGNORECASE
    )

    for meta in parts_meta:
        m = item_pattern.search(meta['text'])
        if m:
            try:
                # A. 提取位置与原始文本
                drug_raw = m.group('drug')
                val_raw = m.group('val')
                unit_raw = m.group('unit')
                days_raw = m.group('days')

                # B. 数据清洗与换算
                drug_raw_clean = re.sub(r'^[.\d\s,天\-]+', '', drug_raw)
                if not drug_raw_clean and "达英-35" in meta['text']: drug_raw_clean = "达英-35"
                if not drug_raw_clean: continue

                val_num = float(val_raw)
                unit_norm = (unit_raw or '').lower()
                if unit_norm in ['i', 'u', 'iu']: unit_norm = 'mg'
                
                freq = (m.group('freq') or '').lower().strip()
                multiplier = 2 if freq == 'bid' else (3 if freq == 'tid' else 1)
                final_val = val_num * multiplier
                v_str = f"{int(final_val)}" if final_val.is_integer() else f"{final_val}"
                
                drug_clean = normalize_drug_name(drug_raw_clean)
                dose_key = f"{drug_clean} {v_str}{unit_norm}"
                days_num = int(days_raw)

                # C. 维测状态判定
                if dose_key in parsed_data: status = "合并"
                elif dose_key in global_cols: status = "复用"
                else: status = "新增"
                logs.append(f"[{status}] {dose_key} (天数:{days_num})")

                # D. 字符级镜像锚定填充 (维测需求1)
                # 锚定1: 药名
                d_start = meta['start'] + m.start('drug')
                for i, char in enumerate(drug_clean[:len(drug_raw)]):
                    aligned_chars[d_start + i] = char
                
                # 锚定2: 剂量数值
                v_start = meta['start'] + m.start('val')
                for i, char in enumerate(v_str[:len(val_raw)]):
                    aligned_chars[v_start + i] = char
                
                # 锚定3: 单位
                u_start = meta['start'] + m.start('unit')
                for i, char in enumerate(unit_norm[:len(unit_raw or " ")]):
                    aligned_chars[u_start + i] = char

                # 锚定4: 天数数值 (在 * 号后面寻找数字位置)
                days_pos_in_part = m.start('days')
                t_start = meta['start'] + days_pos_in_part
                t_str = f"*{days_num}天"
                for i, char in enumerate(t_str):
                    if t_start - 1 + i < len(aligned_chars): # -1 补偿星号位置
                        aligned_chars[t_start - 1 + i] = char

                # E. 累加天数
                parsed_data[dose_key] = parsed_data.get(dose_key, 0) + days_num
            except:
                logs.append(f"[丢弃] {meta['text'][:10]}... (解析异常)")
        else:
            if len(meta['text'].strip()) > 1:
                logs.append(f"[丢弃] {meta['text'].strip()[:10]} (格式不符/截断)")

    return parsed_data, logs, "".join(aligned_chars), raw_text

def main():
    if len(sys.argv) < 2:
        print("用法: python table_splitter.py <输入.xlsx> [输出.xlsx]"); return

    input_file = sys.argv[1]
    output_file = sys.argv[2] if len(sys.argv) >= 3 else os.path.splitext(input_file)[0] + "-o.xlsx"

    print(f">>> 任务启动。处理: {input_file}")
    try:
        df = pd.read_excel(input_file)
    except Exception as e:
        print(f"读取错误: {e}"); return

    processed_rows, global_drug_cols = [], []

    print("\n" + "="*40 + " 1. 逐行维测对比 (位置锚定) " + "="*40)
    for index, row in df.iterrows():
        p_id = row.iloc[0]
        if pd.isna(p_id) or str(p_id).strip() == "":
            print(f"\n[停止] 第 {index+2} 行编号为空。")
            break
            
        medicine_info, logs, aligned_str, clean_input = parse_medicine_column(row.iloc[1], global_drug_cols)
        
        # 维测输出
        print(f"行{index+2} [编号:{p_id}]")
        print(f"原文: {clean_input}")
        print(f"对比: {aligned_str}")
        if logs: print(f"过程: {' | '.join(logs)}")
        print("-" * 80)

        for k in medicine_info.keys():
            if k not in global_drug_cols: global_drug_cols.append(k)
        
        row_dict = row.to_dict()
        row_dict['处理过程信息'] = " | ".join(logs)
        for k, v in medicine_info.items():
            row_dict[k] = v
        processed_rows.append(row_dict)

    global_drug_cols.sort()
    print("\n" + "="*40 + " 2. 药品种类统计 " + "="*40)
    print(f"总计识别种类: {len(global_drug_cols)}")
    for i, col in enumerate(global_drug_cols, 1):
        print(f"{i:03d}. {col}")
    print("="*60 + "\n")

    # 流式写入优化
    print(f">>> 正在执行流式保存至: {output_file}")
    wb = Workbook(write_only=True)
    ws = wb.create_sheet()
    header = list(df.columns) + ['处理过程信息'] + global_drug_cols
    ws.append(header)
    for r_data in processed_rows:
        ws.append([r_data.get(col) for col in header])
    wb.save(output_file)
    print("任务成功完成！")

if __name__ == "__main__":
    main()
