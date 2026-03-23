import pandas as pd
import numpy as np

file_path = "data.xlsx"
df = pd.read_excel(file_path, header=None)

# 列名
cols = [
    "序号","委托单编号","制表人","联系电话","委托单位","单位主管",
    "布套号","布套名称","图号","工件名称","版次",
    "使用设备","接收人","委托时间","接收时间","流程状态"
]
for i in range(df.shape[1] - len(cols)):
    cols.append(f"额外列{i+1}")
df.columns = cols

# 判断有效图号
def is_valid_graph(val):
    val = str(val).strip()
    valid_starts = ('L','Z','T','固','无','9999','/','R','0','1','2','z','l','附','工','通','C','见','A','N','7','8','9','3','4','5','6','八','J','F')
    return val.startswith(valid_starts)

# 判断结尾为“和”或“、”
def ends_with_he_or_dun(val):
    val = str(val).strip()
    return val.endswith("和") or val.endswith("、")

def merge_bs_name_columns(df):
    df_clean = df.copy()
    exclude_receivers = {"李树伟","薛爱迪","杨雪","徐雷","季嗣珉","朱强","孙波","戴鼎章"}

    for idx, row in df_clean.iterrows():
        row_values = row.tolist()
        bs_idx = 7
        graph_idx = 8
        receiver_idx = 12

        receiver_val = str(row_values[receiver_idx]).strip()
        # 如果接收人在排除名单里，跳过该行
        if receiver_val in exclude_receivers:
            continue

        # 逻辑1：布套名称列到有效图号列之间合并
        target_idx = None
        for i in range(graph_idx, len(row_values)):
            if is_valid_graph(row_values[i]):
                target_idx = i
                break

        if target_idx is not None and target_idx > graph_idx:
            merged_str = " ".join([str(row_values[j]).strip() for j in range(bs_idx, target_idx)])
            row_values[bs_idx] = merged_str

            new_idx = bs_idx + 1
            for j in range(target_idx, len(row_values)):
                row_values[new_idx] = row_values[j]
                new_idx += 1
            for j in range(new_idx, len(row_values)):
                row_values[j] = np.nan

        # 逻辑2：结尾为“和”或“、”列，合并左移
        current_idx = bs_idx
        while current_idx < len(row_values) - 1:
            val = str(row_values[current_idx]).strip()
            next_val = str(row_values[current_idx + 1]).strip()
            if val and next_val and ends_with_he_or_dun(val):
                row_values[current_idx] = val + str(next_val)
                for j in range(current_idx + 1, len(row_values) - 1):
                    row_values[j] = row_values[j + 1]
                row_values[-1] = np.nan
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
df_clean = df_clean[df_clean['接收人'].notna() & (df_clean['接收人'].astype(str).str.strip() != "")]

# =====================
# 保存到 Excel
# =====================
df_clean.to_excel("cleaned_data.xlsx", index=False)
print("清洗完成，结果已保存到 cleaned_data.xlsx")