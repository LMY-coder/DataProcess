import pandas as pd
import re

def clean_excel_simple(input_file, output_file):
    """
    简化版的Excel特殊字符清理：
    - 移除换行/制表符
    - 移除控制字符与不可见空白（NBSP、零宽空格、BOM等）
    - 可选移除常见emoji
    - 折叠多余空格
    """
    # 组合需要清理的字符集合：换行/制表符/控制字符/不可见空白
    INVISIBLE_PATTERN = re.compile(r"[\r\n\t\u0000-\u001F\u007F\u00A0\u200B-\u200D\uFEFF]")
    # 常见 emoji 区间（可按需扩展）
    EMOJI_PATTERN = re.compile(r"[\U0001F300-\U0001FAFF\U0001F1E6-\U0001F1FF]", flags=re.UNICODE)

    # Excel 有时会把控制字符保存为字面量如 "_x000D_"，这里一并清理
    EXCEL_ENCODED_CTRL = re.compile(r"_x[0-9A-Fa-f]{4}_")

    def clean_text(value):
        if pd.isna(value):
            return value
        text = str(value)
        # 去除不可见字符
        text = INVISIBLE_PATTERN.sub(" ", text)
        # 去除 Excel 编码的控制字符（如 _x000D_ / _x000A_ / _x0009_ 等）
        text = EXCEL_ENCODED_CTRL.sub(" ", text)
        # 去除 emoji
        try:
            text = EMOJI_PATTERN.sub("", text)
        except re.error:
            # 某些环境下re不支持高位unicode范围，忽略emoji处理
            pass
        # 折叠空格
        text = re.sub(r"\s{2,}", " ", text).strip()
        return text
    # 一次性读取所有sheet为字典，确保不遗漏
    all_sheets = pd.read_excel(input_file, sheet_name=None, dtype=object)

    with pd.ExcelWriter(output_file, engine='openpyxl') as writer:
        for sheet_name, df in all_sheets.items():
            print(f"处理sheet: {sheet_name}")

            # 对所有单元格执行清理（包括原数值/日期字段，会被安全地字符串化清理后再写出）
            df_cleaned = df.applymap(clean_text)

            # 保存到对应sheet
            df_cleaned.to_excel(writer, sheet_name=sheet_name, index=False)
    
    print(f"处理完成! 保存到: {output_file}")

# 使用简化版
clean_excel_simple("input.xlsx", "input_cleaned.xlsx")