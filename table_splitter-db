import pandas as pd
import re
import sys
from collections import OrderedDict

# ===================== 全局配置 =====================
# 忽略的频度词汇（不区分大小写）
IGNORE_FREQ_WORDS = {'qd', 'oncem', 'onc', 'once', 'bid'}
# 等效天数符号
DAY_MARKS = {'*', '×', '×'}
# 正则：匹配剂量（数字+可选单位）
DOSE_REGEX = re.compile(r'(\d+\.?\d*)([a-zA-Z]*)', re.IGNORECASE)
# 正则：匹配 *5天 / ×5天
DAY_REGEX = re.compile(r'([*××]\s*\d+\s*天)', re.IGNORECASE)
# 正则：提取天数数字
DAY_NUM_REGEX = re.compile(r'(\d+)\s*天', re.IGNORECASE)

# 全局唯一药物列：key=药名|剂量，value=列名
GLOBAL_DRUGS = OrderedDict()
LOG_ENABLE = True


# ===================== 日志工具 =====================
def log(msg):
    if LOG_ENABLE:
        print(f"[INFO] {msg}")


# ===================== 核心解析函数 =====================
def clean_freq_tokens(text):
    """清除频度词：qd、bid等"""
    for word in IGNORE_FREQ_WORDS:
        text = re.sub(rf'\b{word}\b', '', text, flags=re.IGNORECASE)
    return text


def split_drug_entries(raw_text):
    """
    拆分药物条目：支持逗号、换行、无逗号换行分隔
    """
    if pd.isna(raw_text):
        return []

    # 统一换行符
    text = raw_text.replace('\r\n', '\n').replace('\r', '\n')
    # 按 逗号 或 换行 分割
    parts = re.split(r'[,|\n]', text)
    entries = [p.strip() for p in parts if p.strip()]
    return entries


def parse_entry(entry):
    """
    解析单条药物
    返回：(drug_name, dose_str, days) | None
    """
    original_entry = entry
    entry = clean_freq_tokens(entry)

    # 1. 提取天数
    days = 0
    content = entry
    day_match = DAY_REGEX.search(entry)
    if day_match:
        day_str = day_match.group(1)
        num_match = DAY_NUM_REGEX.search(day_str)
        if num_match:
            days = int(num_match.group(1))
        content = entry.replace(day_str, '').strip()

    # 2. 从后向前匹配剂量（最准确）
    dose_match = DOSE_REGEX.search(content[::-1])
    if not dose_match:
        log(f"无法识别剂量，丢弃：{original_entry}")
        return None

    # 反转恢复剂量
    num_val = dose_match.group(1)[::-1]
    unit_val = dose_match.group(2)[::-1]
    dose_str = f"{num_val}{unit_val}".lower().strip()

    # 3. 提取药物名称
    split_pos = len(content) - dose_match.end()
    drug_name = content[:split_pos].strip()

    if not drug_name:
        log(f"无药物名称，丢弃：{original_entry}")
        return None

    log(f"解析成功 → 药物={drug_name}  剂量={dose_str}  天数={days}")
    return drug_name, dose_str, days


def get_drug_key(name, dose):
    return f"{name}|{dose}"


def get_column_name(name, dose):
    """列名 = 药物名称 + 剂量"""
    return f"{name} {dose}"


# ===================== 行处理 =====================
def process_row(row_idx, row):
    """处理单行数据"""
    patient_id = str(row.iloc[0]).strip() if not pd.isna(row.iloc[0]) else ''

    # 编号为空 → 停止
    if not patient_id:
        log(f"第{row_idx + 1}行编号为空，终止后续所有行")
        return 'STOP'

    log(f"\n===== 处理第{row_idx + 1}行 | 编号={patient_id} =====")
    drug_text = row.iloc[1]
    log(f"原始用药信息：{drug_text}")

    entries = split_drug_entries(drug_text)
    log(f"拆分得到 {len(entries)} 条药物条目")

    # 同一行相同药物合并天数
    current_row_drugs = {}
    for entry in entries:
        parsed = parse_entry(entry)
        if not parsed:
            continue
        name, dose, day = parsed
        key = get_drug_key(name, dose)

        if key in current_row_drugs:
            old_day = current_row_drugs[key]
            current_row_drugs[key] = old_day + day
            log(f"→ 同行合并：{name} {dose} | 天数 {old_day}+{day}={current_row_drugs[key]}")
        else:
            current_row_drugs[key] = day

    # 初始化所有全局药物列为 0
    row_result = {col: 0 for col in GLOBAL_DRUGS.values()}

    # 填充当前行药物天数
    for key, day_val in current_row_drugs.items():
        name, dose = key.split('|', 1)
        col_name = get_column_name(name, dose)

        if key not in GLOBAL_DRUGS:
            GLOBAL_DRUGS[key] = col_name
            log(f"✅ 新增药物列：{col_name}")
        else:
            log(f"♻️  复用已有列：{col_name}")

        row_result[col_name] = day_val

    return row_result


# ===================== 主函数 =====================
def main():
    if len(sys.argv) != 3:
        print("用法：python table_splitter.py 输入文件.xlsx 输出文件.xlsx")
        sys.exit(1)

    input_path = sys.argv[1]
    output_path = sys.argv[2]

    try:
        log(f"读取输入文件：{input_path}")
        df = pd.read_excel(input_path)
        log(f"读取完成，总行数：{len(df)}")

        output_rows = []
        stop_flag = False

        for idx, row in df.iterrows():
            if stop_flag:
                break
            res = process_row(idx, row)
            if res == 'STOP':
                stop_flag = True
                continue

            # 合并原始数据 + 药物列
            row_dict = row.to_dict()
            row_dict.update(res)
            output_rows.append(row_dict)

        # 生成输出
        log("\n===== 处理完成 =====")
        out_df = pd.DataFrame(output_rows)
        log(f"✅ 全局唯一药物（名称+剂量）总数：{len(GLOBAL_DRUGS)}")
        log(f"✅ 最终药物列：{list(GLOBAL_DRUGS.values())}")

        out_df.to_excel(output_path, index=False)
        log(f"✅ 输出文件已保存：{output_path}")

    except Exception as e:
        log(f"❌ 处理失败：{str(e)}")
        sys.exit(1)


if __name__ == '__main__':
    main()
