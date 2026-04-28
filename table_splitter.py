import pandas as pd
import re
import sys

def normalize_drug_name(name):
    """标准化药物名称：保护达英-35，移除用法和日期"""
    if not name: return ""
    name = str(name).replace('\n', ' ')
    
    # 保护“达英-35”
    if re.search(r'达英\s*-\s*35', name):
        name = re.sub(r'达英\s*-\s*35', '##DY35##', name)

    # 移除用法词汇
    usages = ['口服', '皮下注射', '肌肉注射', '静脉注射', '塞阴道']
    for usage in usages:
        name = re.sub(rf'{usage}[,\s]*', '', name)
    
    # 移除日期括号 (如 2023-08-07)
    name = re.sub(r'[\(（]\s*\d{4}[-/]\d{1,2}[-/]\d{1,2}\s*[\)）]', '', name)
    
    # 恢复达英-35
    name = name.replace('##DY35##', '达英-35')
    return name.strip()

def parse_medicine_column(cell_content):
    """解析第二列：处理剂量换算、单位等同、物理切分"""
    if pd.isna(cell_content): return {}, []

    # 预处理：统一乘号、处理头部逗号
    text = str(cell_content).replace('，', ',').replace('×', '*').replace('✕', '*').replace('x', '*')
    text = re.sub(r'^[,]+', '', text)
    
    # 保护“达英-35”粘连并强制隔离后续数值
    text = re.sub(r'达英\s*[-－]\s*35', '##DY35## ', text)

    # 物理分段：利用“天”字强制切分，防止截断数据污染前项
    temp_text = re.sub(r'(\d+\s*天)', r'\1|||', text)
    parts = re.split(r'\|\|\||[,]+|\n', temp_text)
    
    parsed_data = {}
    logs = []
    
    # 正则提取：药名、数值、单位、频次、天数
    item_pattern = re.compile(
        r'(?P<drug>.+?)\s*'
        r'(?P<val>\d+\.?\d*)(?P<unit>mg|iu|[iumg片\u4e00-\u9fa5]*)\s*'
        r'(?P<freq>qd|q3d|w3d|oncem|onc|once|bid|tid|\s)*'
        r'\*\s*(?P<days>\d+)', 
        re.IGNORECASE | re.DOTALL
    )

    for part in parts:
        part = part.strip()
        if not part: continue
        
        m = item_pattern.search(part)
        if m:
            drug_raw = m.group('drug').strip()
            # 清理药名前导干扰符
            drug_raw = re.sub(r'^[.\d\s,天\-]+', '', drug_raw)
            if not drug_raw and "##DY35##" not in part: continue
            if not drug_raw and "##DY35##" in part: drug_raw = "##DY35##"

            try:
                val = float(m.group('val'))
                unit = (m.group('unit') or '').lower()
                # 单位处理：i/u/iu/mg 统一转为 mg
                if any(u in unit for u in ['i', 'u', 'iu', 'mg']):
                    unit = 'mg'
                
                freq = (m.group('freq') or '').lower().strip()
                days = int(m.group('days'))

                # 剂量换算：bid * 2, tid * 3
                multiplier = 1
                if freq == 'bid': multiplier = 2
                elif freq == 'tid': multiplier = 3
                
                final_val = val * multiplier
                # 格式化数值：如果是整数则去掉小数点
                dose_val_str = f"{int(final_val)}" if final_val.is_integer() else f"{final_val}"
                dose_str = f"{dose_val_str}{unit}"
                
                drug_clean = normalize_drug_name(drug_raw)
                key = f"{drug_clean} {dose_str}"

                log_entry = f"提取: {key}({days}天)"
                if multiplier > 1:
                    log_entry += f" [换算自{val}{freq}]"
                if drug_raw != drug_clean:
                    log_entry += f" [清洗附带信息]"
                logs.append(log_entry)
                
                # 同行累加天数
                parsed_data[key] = parsed_data.get(key, 0) + days
            except Exception:
                continue
        else:
            log_p = part.replace('\n', ' ')
            if len(log_p) > 1:
                logs.append(f"丢弃截断数据: {log_p}")
                
    return parsed_data, logs

def main(input_path, output_path):
    print(f">>> 任务启动。读取文件: {input_path}")
    try:
        df = pd.read_excel(input_path)
    except Exception as e:
        print(f"读取失败: {e}")
        return

    source_cols = list(df.columns)
    new_rows = []
    global_drug_columns = set()

    for index, row in df.iterrows():
        # 规则 3：第一列编号为空则停止
        p_id = row.iloc[0]
        if pd.isna(p_id) or str(p_id).strip() == "":
            print(f"\n[停止] 第 {index+2} 行编号为空，任务结束。")
            break
            
        print(f"\n>>> 处理编号: {p_id}")
        # row.iloc[1] 为第二列
        medicine_info, process_logs = parse_medicine_column(row.iloc[1])
        
        new_row = row.to_dict()
        new_row['处理过程信息'] = " | ".join(process_logs)
        
        for drug_key, days in medicine_info.items():
            if drug_key in global_drug_columns:
                print(f"  [复用] 匹配列: '{drug_key}'")
            else:
                print(f"  [新增] 发现新药物: '{drug_key}'")
                global_drug_columns.add(drug_key)
            new_row[drug_key] = days
        new_rows.append(new_row)

    # 排序新列
    drug_cols_sorted = sorted(list(global_drug_columns))
    
    print("\n" + "="*60 + "\n汇总摘要 (保存前预览):\n" + "="*60)
    for i, name in enumerate(drug_cols_sorted, 1):
        print(f"{i:02d}. {name}")
    print("="*60)

    # 合并结果
    result_df = pd.DataFrame(new_rows)
    # 组合列顺序：源列 + 处理过程信息 + 排序后的新药列
    final_order = source_cols + ['处理过程信息'] + drug_cols_sorted
    final_order = [c for c in final_order if c in result_df.columns]
    
    result_df[final_order].to_excel(output_path, index=False)
    print(f"\n成功！结果已保存至: {output_path}")

if __name__ == "__main__":
    if len(sys.argv) != 3:
        print("用法: python table_splitter.py <输入.xlsx> <输出.xlsx>")
    else:
        main(sys.argv[1], sys.argv[2])
