# -*- coding: utf-8 -*-
"""
隧道监测数据网页界面（Streamlit 版）

功能说明：
1. 支持 Excel 中包含多个 sheet，先选数据表，再选里程、隧道监测位置
2. 支持两类时间字段：
   - 数值时间，例如：时间/d
   - 日期时间，例如：时间
3. 点击“读取数据”后，不自动显示表格
4. 点击“绘制曲线”后，显示当前筛选条件下的曲线和点
5. 点击“显示数据”后，只显示当前筛选条件下的原始数据结果表格
6. 点击“隧道监测布置图”后，显示隧道监测布置图
7. 不显示转换后的长表，不显示其他多余表格
8. 日期时间横坐标刻度显示为：2026.4.16 0:00:00
9. 布置图按较小比例显示，避免过大铺满页面

默认文件路径：
D:\jupyter\pingtai\data.xlsx

布置图路径：
D:\jupyter\pingtai\tunnel_layout.png

运行命令：
streamlit run D:\jupyter\pingtai\app.py
"""

import os
import pandas as pd
import plotly.graph_objects as go
import streamlit as st

# =========================
# 1. 页面基础设置
# =========================
st.set_page_config(
    page_title="隧道监测变形数据交互界面",
    layout="wide"
)

st.title("隧道监测变形数据交互界面")

DEFAULT_FILE_PATH = "data.xlsx"
LAYOUT_IMAGE_PATH = "tunnel_layout.png"
LAYOUT_IMAGE_WIDTH = 900   # 布置图显示宽度，可按需要改成 700、800、1000


# =========================
# 2. 工具函数
# =========================
def resolve_time_column(columns):
    """
    自动识别时间列
    优先顺序：
    1. 时间
    2. 时间/d
    3. 包含“时间”的列
    """
    columns = [str(c).strip() for c in columns]

    for candidate in ["时间", "时间/d"]:
        if candidate in columns:
            return candidate

    for col in columns:
        if "时间" in col:
            return col

    raise ValueError("未找到时间列，请确认表中存在“时间”或“时间/d”列")


def infer_time_type(series, time_col_name):
    """
    判断时间列属于：
    - numeric：数值时间（例如时间/d）
    - datetime：日期时间
    """
    col_name = str(time_col_name).strip()

    if "/d" in col_name.lower():
        return "numeric"

    dt_try = pd.to_datetime(series, errors="coerce")
    valid_count = dt_try.notna().sum()

    if valid_count > 0 and valid_count >= max(1, len(series) // 2):
        return "datetime"

    return "numeric"


def parse_monitor_column(col_name):
    """
    解析监测列名，例如：
    K701+118.9-中标靶水平收敛
    K701+118.9-顶标靶相对下沉
    K701+112.5--底标靶水平收敛

    返回：
    mileage, location
    """
    col_name = str(col_name).strip()

    if "-" not in col_name:
        return None, None

    parts = col_name.split("-", 1)
    mileage = parts[0].strip()
    location = parts[1].strip().lstrip("-").strip()

    if not mileage or not location:
        return None, None

    return mileage, location


@st.cache_data
def get_workbook_sheets(file_path):
    """
    获取工作簿中的所有 sheet 名称
    """
    xls = pd.ExcelFile(file_path)
    return xls.sheet_names


@st.cache_data
def load_sheet_data(file_path, sheet_name):
    """
    读取指定 sheet，并转换成长表

    返回：
    raw_df: 原始宽表
    long_df: 转换后的长表
    column_map: {里程: {隧道监测位置: 原始列名}}
    mileages: 里程列表
    locations_by_mileage: 各里程下可用隧道监测位置列表
    time_col: 时间列名
    time_type: 时间类型（numeric / datetime）
    """
    raw_df = pd.read_excel(file_path, sheet_name=sheet_name)
    raw_df.columns = [str(c).strip() for c in raw_df.columns]

    time_col = resolve_time_column(raw_df.columns)
    time_type = infer_time_type(raw_df[time_col], time_col)

    if time_type == "datetime":
        raw_df[time_col] = pd.to_datetime(raw_df[time_col], errors="coerce")
    else:
        raw_df[time_col] = pd.to_numeric(raw_df[time_col], errors="coerce")

    raw_df = raw_df.dropna(subset=[time_col]).copy()

    long_rows = []
    column_map = {}

    for col in raw_df.columns:
        if col == time_col:
            continue

        mileage, location = parse_monitor_column(col)
        if mileage is None:
            continue

        temp = pd.DataFrame({
            "时间值": raw_df[time_col],
            "里程": mileage,
            "隧道监测位置": location,
            "测值": pd.to_numeric(raw_df[col], errors="coerce"),
            "原始列名": col
        })

        temp = temp.dropna(subset=["测值"]).copy()
        if not temp.empty:
            long_rows.append(temp)

        if mileage not in column_map:
            column_map[mileage] = {}
        column_map[mileage][location] = col

    if not long_rows:
        raise ValueError("没有识别到有效监测列，请检查列名是否类似：K701+118.9-顶标靶相对下沉")

    long_df = pd.concat(long_rows, ignore_index=True)
    long_df = long_df.sort_values(["里程", "隧道监测位置", "时间值"]).reset_index(drop=True)

    mileages = sorted(long_df["里程"].dropna().unique().tolist())

    locations_by_mileage = {}
    for mileage in mileages:
        mileage_locations = (
            long_df.loc[long_df["里程"] == mileage, "隧道监测位置"]
            .dropna()
            .unique()
            .tolist()
        )
        locations_by_mileage[mileage] = sorted(mileage_locations)

    return raw_df, long_df, column_map, mileages, locations_by_mileage, time_col, time_type


def filter_long_data(long_df, selected_mileage, selected_locations, time_range, time_type):
    """
    按当前筛选条件过滤长表（用于绘图）
    """
    if not selected_mileage or not selected_locations:
        return pd.DataFrame()

    t_start, t_end = time_range

    df_plot = long_df[
        (long_df["里程"] == selected_mileage) &
        (long_df["隧道监测位置"].isin(selected_locations))
    ].copy()

    if time_type == "datetime":
        df_plot = df_plot[
            (df_plot["时间值"] >= pd.Timestamp(t_start)) &
            (df_plot["时间值"] <= pd.Timestamp(t_end))
        ].copy()
    else:
        df_plot = df_plot[
            (df_plot["时间值"] >= float(t_start)) &
            (df_plot["时间值"] <= float(t_end))
        ].copy()

    return df_plot


def filter_raw_data(raw_df, time_col, column_map, selected_mileage, selected_locations, time_range, time_type):
    """
    按当前筛选条件过滤原始数据结果表格（宽表形式）
    只保留：
    - 时间列
    - 当前里程 + 当前所选隧道监测位置对应的原始列
    """
    if not selected_mileage or not selected_locations:
        return pd.DataFrame()

    target_columns = [time_col]

    if selected_mileage in column_map:
        for location in selected_locations:
            original_col = column_map[selected_mileage].get(location)
            if original_col and original_col in raw_df.columns:
                target_columns.append(original_col)

    if len(target_columns) == 1:
        return pd.DataFrame()

    filtered_raw = raw_df[target_columns].copy()
    t_start, t_end = time_range

    if time_type == "datetime":
        filtered_raw = filtered_raw[
            (filtered_raw[time_col] >= pd.Timestamp(t_start)) &
            (filtered_raw[time_col] <= pd.Timestamp(t_end))
        ].copy()
    else:
        filtered_raw = filtered_raw[
            (filtered_raw[time_col] >= float(t_start)) &
            (filtered_raw[time_col] <= float(t_end))
        ].copy()

    filtered_raw = filtered_raw.sort_values(by=time_col).reset_index(drop=True)
    return filtered_raw


def get_y_axis_title(selected_sheet, selected_locations):
    """
    纵坐标标题
    """
    if any("速率" in str(x) for x in selected_locations) or "Sheet2" in str(selected_sheet):
        return "变形速率（mm/d）"
    return "变形量（mm）"


def build_plot(df_plot, selected_sheet, selected_mileage, selected_locations, time_type, show_lines=True, show_markers=True):
    """
    绘制曲线图
    """
    if show_lines and show_markers:
        mode = "lines+markers"
    elif show_lines:
        mode = "lines"
    elif show_markers:
        mode = "markers"
    else:
        mode = "lines+markers"

    fig = go.Figure()

    for location in selected_locations:
        sub = df_plot[df_plot["隧道监测位置"] == location].sort_values("时间值")
        if sub.empty:
            continue

        if time_type == "datetime":
            hover_time = "时间：%{x|%Y.%-m.%-d %-H:%M:%S}<br>"
        else:
            hover_time = "时间：%{x:.3f} d<br>"

        fig.add_trace(
            go.Scatter(
                x=sub["时间值"],
                y=sub["测值"],
                mode=mode,
                name=location,
                hovertemplate=(
                    f"数据表：{selected_sheet}<br>"
                    f"里程：{selected_mileage}<br>"
                    f"隧道监测位置：{location}<br>"
                    + hover_time +
                    "测值：%{y:.4f}<extra></extra>"
                )
            )
        )

    x_axis_title = "时间" if time_type == "datetime" else "时间 / d"
    y_axis_title = get_y_axis_title(selected_sheet, selected_locations)

    fig.update_layout(
        title=f"{selected_sheet} | {selected_mileage} 监测曲线",
        xaxis_title=x_axis_title,
        yaxis_title=y_axis_title,
        template="plotly_white",
        hovermode="x unified",
        height=620,
        legend_title="隧道监测位置"
    )

    fig.update_xaxes(showgrid=True)
    fig.update_yaxes(showgrid=True, zeroline=True, zerolinewidth=1)

    if time_type == "datetime":
        fig.update_xaxes(
            tickformat="%Y.%-m.%-d %-H:%M:%S",
            hoverformat="%Y.%-m.%-d %-H:%M:%S",
            tickangle=0
        )

    return fig


# =========================
# 3. session_state 初始化
# =========================
if "file_loaded" not in st.session_state:
    st.session_state.file_loaded = False

if "loaded_file_path" not in st.session_state:
    st.session_state.loaded_file_path = DEFAULT_FILE_PATH

if "available_sheets" not in st.session_state:
    st.session_state.available_sheets = []

if "show_plot" not in st.session_state:
    st.session_state.show_plot = False

if "show_data" not in st.session_state:
    st.session_state.show_data = False

if "show_layout_image" not in st.session_state:
    st.session_state.show_layout_image = False


# =========================
# 4. 侧边栏：文件读取
# =========================
st.sidebar.header("参数设置")

file_path = st.sidebar.text_input(
    "Excel 文件路径",
    value=st.session_state.loaded_file_path
)

load_button = st.sidebar.button("读取数据")

if load_button:
    try:
        available_sheets = get_workbook_sheets(file_path)

        st.session_state.loaded_file_path = file_path
        st.session_state.available_sheets = available_sheets
        st.session_state.file_loaded = True

        st.session_state.show_plot = False
        st.session_state.show_data = False
        st.session_state.show_layout_image = False

        st.sidebar.success("数据读取成功")
    except Exception as e:
        st.session_state.file_loaded = False
        st.session_state.available_sheets = []
        st.session_state.show_plot = False
        st.session_state.show_data = False
        st.session_state.show_layout_image = False
        st.sidebar.error(f"读取失败：{e}")


# =========================
# 5. 读取成功后：选择数据表 + 筛选条件
# =========================
if st.session_state.file_loaded and st.session_state.available_sheets:
    selected_sheet = st.sidebar.selectbox(
        "选择数据表",
        st.session_state.available_sheets
    )

    try:
        raw_df, long_df, column_map, mileages, locations_by_mileage, time_col, time_type = load_sheet_data(
            st.session_state.loaded_file_path,
            selected_sheet
        )
    except Exception as e:
        st.error(f"当前数据表解析失败：{e}")
        st.stop()

    st.sidebar.markdown("---")
    selected_mileage = st.sidebar.selectbox("选择里程", mileages)

    mileage_locations = locations_by_mileage.get(selected_mileage, [])

    location_key = f"location_select_{selected_sheet}_{selected_mileage}"
    selected_locations = st.sidebar.multiselect(
        "选择隧道监测位置",
        options=mileage_locations,
        default=mileage_locations,
        key=location_key
    )

    min_time = raw_df[time_col].min()
    max_time = raw_df[time_col].max()

    if time_type == "datetime":
        time_range = st.sidebar.slider(
            "选择时间范围",
            min_value=min_time.to_pydatetime(),
            max_value=max_time.to_pydatetime(),
            value=(min_time.to_pydatetime(), max_time.to_pydatetime())
        )
    else:
        min_value = float(min_time)
        max_value = float(max_time)
        step_value = round((max_value - min_value) / 200, 6) if max_value > min_value else 0.01
        if step_value <= 0:
            step_value = 0.01

        time_range = st.sidebar.slider(
            "选择时间范围",
            min_value=min_value,
            max_value=max_value,
            value=(min_value, max_value),
            step=step_value
        )

    show_markers = st.sidebar.checkbox("显示散点", value=True)
    show_lines = st.sidebar.checkbox("显示连线", value=True)

    plot_button = st.sidebar.button("绘制曲线")
    show_data_button = st.sidebar.button("显示数据")
    show_layout_button = st.sidebar.button("隧道监测布置图")

    if plot_button:
        st.session_state.show_plot = True

    if show_data_button:
        st.session_state.show_data = True

    if show_layout_button:
        st.session_state.show_layout_image = True

    st.success(
        f"数据已加载完成。当前数据表：{selected_sheet}。"
        f"点击“绘制曲线”“显示数据”或“隧道监测布置图”后展示结果。"
    )

    if st.session_state.show_layout_image:
        st.subheader("隧道监测布置图")
        if os.path.exists(LAYOUT_IMAGE_PATH):
            left_col, center_col, right_col = st.columns([1, 3, 1])
            with center_col:
                st.image(LAYOUT_IMAGE_PATH, width=LAYOUT_IMAGE_WIDTH)
        else:
            st.warning(
                "未找到布置图文件，请把图片放到："
                + LAYOUT_IMAGE_PATH
            )

    df_plot = filter_long_data(
        long_df=long_df,
        selected_mileage=selected_mileage,
        selected_locations=selected_locations,
        time_range=time_range,
        time_type=time_type
    )

    filtered_raw = filter_raw_data(
        raw_df=raw_df,
        time_col=time_col,
        column_map=column_map,
        selected_mileage=selected_mileage,
        selected_locations=selected_locations,
        time_range=time_range,
        time_type=time_type
    )

    if st.session_state.show_plot:
        st.subheader("监测曲线图")

        if not selected_locations:
            st.warning("请至少选择一个隧道监测位置")
        elif df_plot.empty:
            st.warning("当前筛选条件下没有可绘制的数据")
        else:
            fig = build_plot(
                df_plot=df_plot,
                selected_sheet=selected_sheet,
                selected_mileage=selected_mileage,
                selected_locations=selected_locations,
                time_type=time_type,
                show_lines=show_lines,
                show_markers=show_markers
            )
            st.plotly_chart(fig, use_container_width=True)

    if st.session_state.show_data:
        st.subheader("当前筛选出的原始数据结果表格")

        if not selected_locations:
            st.warning("请至少选择一个隧道监测位置")
        elif filtered_raw.empty:
            st.warning("当前筛选条件下没有可显示的原始数据")
        else:
            st.dataframe(filtered_raw, use_container_width=True)

else:
    st.info("请先在左侧输入 Excel 文件路径，然后点击“读取数据”")
