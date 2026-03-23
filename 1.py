import streamlit as st
import pandas as pd
import plotly.express as px

st.title("数控编程看板（无列名版）")

# ===== 读取数据（无表头）=====
df = pd.read_excel("data.xlsx", header=None)

# ===== 手动指定列 =====
df["申请人"] = df[2]
df["编程人"] = df[5]
df["分厂"] = df[4]
df["机床"] = df[11]
df["等级"] = df[10]
df["创建时间"] = pd.to_datetime(df[13], errors='coerce')
df["完成时间"] = pd.to_datetime(df[14], errors='coerce')
df["状态"] = df[15]

# ===== 衍生字段 =====
df["月份"] = df["创建时间"].dt.to_period("M").astype(str)
df["完成时长(小时)"] = (df["完成时间"] - df["创建时间"]).dt.total_seconds() / 3600

# ===== 侧边栏筛选 =====
st.sidebar.header("筛选")

person = st.sidebar.selectbox("选择编程人", ["全部"] + list(df["编程人"].dropna().unique()))
month = st.sidebar.selectbox("选择月份", ["全部"] + list(df["月份"].dropna().unique()))

# ===== 数据筛选 =====
df_filtered = df.copy()

if person != "全部":
    df_filtered = df_filtered[df_filtered["编程人"] == person]

if month != "全部":
    df_filtered = df_filtered[df_filtered["月份"] == month]

# ===== KPI =====
col1, col2, col3 = st.columns(3)

col1.metric("总任务数", len(df_filtered))
col2.metric("完成数量", (df_filtered["状态"] == "完成").sum())
col3.metric("待接收", (df_filtered["状态"] == "待接收").sum())

# ===== 人员工作量 =====
st.subheader("人员工作量")

person_count = df_filtered["编程人"].value_counts().reset_index()
person_count.columns = ["编程人", "任务数"]

fig1 = px.bar(person_count, x="编程人", y="任务数")
st.plotly_chart(fig1)

# ===== 机床使用 =====
st.subheader("机床使用频率")

machine_count = df_filtered["机床"].value_counts().reset_index()
machine_count.columns = ["机床", "次数"]

fig2 = px.bar(machine_count, x="机床", y="次数")
st.plotly_chart(fig2)

# ===== 等级分布 =====
st.subheader("任务等级分布")

level_count = df_filtered["等级"].value_counts().reset_index()
level_count.columns = ["等级", "数量"]

fig3 = px.pie(level_count, names="等级", values="数量")
st.plotly_chart(fig3)

# ===== 完成时长分析 =====
st.subheader("完成时长分析")

time_df = df_filtered.dropna(subset=["完成时长(小时)"])

fig4 = px.box(time_df, x="编程人", y="完成时长(小时)")
st.plotly_chart(fig4)