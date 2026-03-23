import streamlit as st
import pandas as pd
import numpy as np

st.set_page_config(page_title="委托单数据看板", layout="wide")
st.title("📊 委托单数据看板")

# =====================
# 1. 数据读取（清洗后的 Excel 已有表头）
# =====================
file_path = "cleaned_data.xlsx"
df = pd.read_excel(file_path)

# =====================
# 2. 保证列顺序和名称
# =====================
cols = [
    "序号", "委托单编号", "制表人", "联系电话", "委托单位",
    "单位主管", "布套号", "布套名称", "图号", "工件名称",
    "版次", "使用设备", "接收人", "委托时间", "接收时间", "流程状态"
]
df = df[cols]  # 保证列顺序正确

# =====================
# 3. 数据预处理
# =====================
# 时间字段解析（支持 2026/3/22 10:31:23 格式）
df["委托时间"] = pd.to_datetime(df["委托时间"], errors='coerce')
df["接收时间"] = pd.to_datetime(df["接收时间"], errors='coerce')

# 去掉接收人为空的行（分析用）
df_valid = df.dropna(subset=["接收人"])

# =====================
# 4. 总览指标
# =====================
col1, col2, col3 = st.columns(3)
col1.metric("总委托单数", len(df))
col2.metric("已接收数量", df["接收时间"].notna().sum())
col3.metric("待接收数量", df["接收时间"].isna().sum())
st.divider()

# =====================
# 5. 接收人工作量分析（改进版）
# =====================
st.subheader("👤 接收人工作量分析")

# 固定 8 个接收人
receivers_list = ["李树伟","薛爱迪","杨雪","徐雷","季嗣珉","朱强","孙波","戴鼎章"]
df_receivers = df_valid[df_valid["接收人"].isin(receivers_list)]

# 选择时间范围
min_date = df_receivers["接收时间"].min().date()
max_date = df_receivers["接收时间"].max().date()
start_date = st.date_input("选择起始日期", min_value=min_date, max_value=max_date, value=min_date)
end_date = st.date_input("选择结束日期", min_value=min_date, max_value=max_date, value=max_date)

# 按选择日期筛选
df_filtered = df_receivers[
    (df_receivers["接收时间"].dt.date >= start_date) &
    (df_receivers["接收时间"].dt.date <= end_date)
]

# 选择统计单位：年或月
period_option = st.radio("选择统计周期", ("按年统计", "按月份统计"))

if period_option == "按年统计":
    df_filtered["统计周期"] = df_filtered["接收时间"].dt.year.astype(str)
else:
    df_filtered["统计周期"] = df_filtered["接收时间"].dt.to_period("M").astype(str)

# 计算工作量
workload = df_filtered.groupby(["接收人", "统计周期"]).size().unstack(fill_value=0)

st.bar_chart(workload)
st.dataframe(workload)


# =====================
# 6. 机床频次统计（按年份）
# =====================
st.subheader("⚙️ 机床使用频次统计")

# 用户选择年份
available_years = df_valid["接收时间"].dt.year.dropna().unique()
selected_year = st.selectbox("选择年份", sorted(available_years, reverse=True))

# 按年份筛选数据
df_year = df_valid[df_valid["接收时间"].dt.year == selected_year]

# 统计机床使用次数
machine_count = df_year["使用设备"].value_counts()

st.bar_chart(machine_count)
st.dataframe(machine_count.rename("使用次数"))

# =====================
# 7. 处理时长分析
# =====================
st.subheader("⏱ 处理时长分析")
df_valid["处理时长(小时)"] = (df_valid["接收时间"] - df_valid["委托时间"]).dt.total_seconds() / 3600
st.write("平均处理时长（小时）：", round(df_valid["处理时长(小时)"].mean(), 2))
st.line_chart(df_valid["处理时长(小时)"])

# =====================
# 8. 按接收人查看
# =====================
st.subheader("🔍 按接收人筛选")
selected_receiver = st.selectbox("选择接收人", df_valid["接收人"].dropna().unique())
df_person = df_valid[df_valid["接收人"] == selected_receiver]
st.write(f"{selected_receiver} 的任务列表")
st.dataframe(df_person)

# =====================
# 9. 原始数据展示
# =====================
st.subheader("📋 原始数据")
st.dataframe(df)