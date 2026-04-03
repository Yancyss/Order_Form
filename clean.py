"""
=========================
数据清洗逻辑说明（核心规则）
=========================

一、数据背景问题
--------------------------------
原始 Excel 数据存在以下问题：
1. 列是按“空格分隔”导出的，导致字段错位（尤其是布套名称、图号）
2. 某些字段（如布套名称）被拆分到多列
3. 图号位置不固定，需要动态识别
4. 存在“和”、“、”等分隔符导致字段断裂
5. 部分行接收人为空或异常，需要过滤

--------------------------------

二、整体处理流程
--------------------------------
1. 读取 Excel（无表头）
2. 强制指定标准列名（不足补列，多余截断）
3. 对每一行进行逐行清洗（核心）
4. 删除无效数据（接收人为空）
5. 输出干净数据

--------------------------------

三、核心清洗逻辑（merge_bs_name_columns）
--------------------------------

【逻辑1：布套名称合并】
--------------------------------
问题：
布套名称可能被拆成多列，例如：
列7: "垣曲3#水轮机"
列8: "-座环装配"
列9: "Z1a012985"（实际是图号）

解决：
- 从“图号列”开始向右扫描
- 找到第一个“符合图号规则”的值（通过 is_valid_graph 判断）
- 将“布套名称列 → 图号前一列”全部拼接
- 再把后面的数据整体左移（对齐字段）

--------------------------------

【逻辑2：处理“和 / 、”断裂】
--------------------------------
问题：
数据可能像：
"零件A、"   "零件B"

解决：
- 如果当前列以“和”或“、”结尾
- 则与下一列拼接
- 后续所有列整体左移一位

--------------------------------

【逻辑3：跳过特殊人员】
--------------------------------
- 对部分指定接收人（如杨雪、徐雷等）
- 不执行清洗（认为其数据已规范）

--------------------------------

四、辅助函数说明
--------------------------------

1. is_valid_graph(val)
   判断是否为“图号”
   规则：是否以指定字符开头（如 L/Z/T/数字等）

2. ends_with_he_or_dun(val)
   判断字段是否以：
   - “和”
   - “、”
   结尾，用于触发合并

--------------------------------

五、数据过滤规则
--------------------------------

最终数据中：
- 删除“接收人为空”的行
- 保留有效业务数据

--------------------------------

六、输出结果
--------------------------------

生成 cleaned_data.xlsx：
- 列结构标准化
- 布套名称完整
- 图号位置正确
- 无空接收人

--------------------------------

七、核心思想总结（很重要）
--------------------------------

✔ 不依赖固定列位置  
✔ 基于“特征规则”动态纠正  
✔ 按行处理，避免全局污染  
✔ 出错数据尽量修复，无法修复则剔除  

--------------------------------
"""
import pandas as pd
import numpy as np

file_path = "data.xlsx"

# =====================
# ✅ 关键修复1：禁止自动解析时间
# =====================
df = pd.read_excel(file_path, header=None, dtype=str)

# =====================
# 列名
# =====================
cols = [
    "序号","委托单编号","制表人","联系电话","委托单位","单位主管",
    "布套号","布套名称","图号","工件名称","版次",
    "使用设备","接收人","委托时间","接收时间","流程状态"
]
for i in range(df.shape[1] - len(cols)):
    cols.append(f"额外列{i+1}")
df.columns = cols

# =====================
# 判断有效图号（略增强鲁棒性）
# =====================
def is_valid_graph(val):
    val = str(val).strip()
    if val == "" or val.lower() == "nan":
        return False

    valid_starts = (
        'L','Z','T','固','无','9999','/','R',
        '0','1','2','3','4','5','6','7','8','9',
        'z','l','附','工','通','C','见','A','N','八','J','F'
    )
    return val.startswith(valid_starts)

# =====================
# 判断结尾为“和”或“、”
# =====================
def ends_with_he_or_dun(val):
    val = str(val).strip()
    return val.endswith("和") or val.endswith("、")

# =====================
# 核心清洗函数
# =====================
def merge_bs_name_columns(df):
    df_clean = df.copy()

    exclude_receivers = {"李树伟","薛爱迪","杨雪","徐雷","季嗣珉","朱强","孙波","戴鼎章"}

    for idx, row in df_clean.iterrows():

        # =====================
        # ✅ 关键修复2：强制整行转字符串（彻底杜绝 datetime）
        # =====================
        row_values = [str(x).strip() if pd.notna(x) else "" for x in row.tolist()]

        bs_idx = 7
        graph_idx = 8
        receiver_idx = 12

        receiver_val = row_values[receiver_idx]

        # 跳过指定人员
        if receiver_val in exclude_receivers:
            continue

        # =====================
        # 逻辑1：布套名称合并
        # =====================
        target_idx = None
        for i in range(graph_idx, len(row_values)):
            if is_valid_graph(row_values[i]):
                target_idx = i
                break

        if target_idx is not None and target_idx > graph_idx:
            merged_str = " ".join([
                row_values[j] for j in range(bs_idx, target_idx)
                if row_values[j] != ""
            ])
            row_values[bs_idx] = merged_str

            new_idx = bs_idx + 1
            for j in range(target_idx, len(row_values)):
                row_values[new_idx] = row_values[j]
                new_idx += 1

            for j in range(new_idx, len(row_values)):
                row_values[j] = ""

        # =====================
        # 逻辑2：处理“和 / 、”
        # =====================
        current_idx = bs_idx
        while current_idx < len(row_values) - 1:
            val = row_values[current_idx]
            next_val = row_values[current_idx + 1]

            if val and next_val and ends_with_he_or_dun(val):
                row_values[current_idx] = val + next_val

                for j in range(current_idx + 1, len(row_values) - 1):
                    row_values[j] = row_values[j + 1]

                row_values[-1] = ""
            else:
                current_idx += 1

        df_clean.loc[idx] = row_values

    return df_clean


# =====================
# 应用清洗
# =====================
df_clean = merge_bs_name_columns(df)
df_clean = df_clean.iloc[:, :16]

# =====================
# 删除接收人为空的行
# =====================
df_clean = df_clean[
    df_clean['接收人'].notna() &
    (df_clean['接收人'].astype(str).str.strip() != "")
]

# =====================
# ✅ 关键修复3：最终全列再保险（防止任何异常类型）
# =====================
df_clean = df_clean.astype(str)

# =====================
# 保存结果
# =====================
df_clean.to_excel("cleaned_data.xlsx", index=False)

print("✅ 清洗完成，结果已保存到 cleaned_data.xlsx")