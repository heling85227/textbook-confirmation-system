"""
共享 UI 组件模块

提供各页面通用的 Streamlit 组件：
- show_header: 统一页面头部
- styled_dataframe: 统一表格样式（对齐、隐藏ID、格式化）
- excel_export / excel_export_by_class: DataFrame → Excel
- apply_excel_borders: Excel 边框样式
"""
import io
import streamlit as st
import pandas as pd
from openpyxl.styles import Border, Side, Font


# ── 页面头部 ──

def show_header(title: str, subtitle: str = None) -> None:
    """渲染统一的页面头部（蓝色渐变标题栏）"""
    html = f'<div class="app-header"><h1>{title}</h1>'
    if subtitle:
        html += f'<p>{subtitle}</p>'
    html += '</div>'
    st.markdown(html, unsafe_allow_html=True)


# ── 统一表格组件 ──

# 金额列（¥格式）
_MONEY_COLS = {"定价", "单价", "结算单价", "结算价", "结算价(元)", "calc_price", "小计", "小计(元)", "subtotal",
               "price", "actual_price", "总价", "费用", "费用合计", "总计", "实洋(元)", "合计(元)",
               "定价(元)", "实洋价", "实洋总计"}

# 数值列（整数格式）
_NUM_COLS = {"数量", "quantity", "total_qty", "征订总量", "teacher_qty", "教师用书", "分配数量", "学生人数",
             "教材数", "序号", "记录条数", "student_count"}

# 折数列（百分比格式）
_PCT_COLS = {"折扣率", "discount_rate"}

# ID 列（隐藏）
_ID_COLS = {"id", "student_id_pk", "textbook_id", "semester_id"}

# 日期列（居中）
_DATE_COLS = {"日期", "发放日期", "distribute_date", "created_at", "学期", "semester_name",
              "领书时间", "出版日期"}


def styled_dataframe(df: pd.DataFrame, hide_ids: bool = True, money_cols: set = None,
                     num_cols: set = None, **kwargs) -> None:
    """
    统一渲染 DataFrame，自动格式化金额/数量/折扣率、隐藏 ID 列。
    所有数据列由 CSS 统一居中。

    Args:
        df: 要显示的 DataFrame
        hide_ids: 是否自动隐藏 ID 列
        money_cols: 额外金额列（除内置外）
        num_cols: 额外数值列（除内置外）
        **kwargs: 传给 st.dataframe 的额外参数
    """
    if df is None or df.empty:
        st.info("📭 暂无数据")
        return

    import streamlit as st
    from streamlit import column_config as cc

    display_df = df.copy()
    column_config = {}

    # 自动检测并隐藏 ID 列
    cols_to_hide = set()
    if hide_ids:
        for c in display_df.columns:
            low = str(c).lower()
            if low in {x.lower() for x in _ID_COLS} or low.endswith("_id"):
                cols_to_hide.add(c)

    # 为各列设置对齐和格式
    for c in display_df.columns:
        if c in cols_to_hide:
            column_config[c] = None  # 隐藏
            continue

        cname = str(c)
        # 金额列 → ¥格式
        is_money = cname in _MONEY_COLS or (money_cols and cname in money_cols)
        is_num = cname in _NUM_COLS or (num_cols and cname in num_cols)
        is_date = cname in _DATE_COLS

        if is_money:
            column_config[c] = cc.NumberColumn(cname, format="¥%.2f", help=None, alignment="center")
        elif cname in _PCT_COLS:
            column_config[c] = cc.NumberColumn(cname, format="%.0f%%", help=None, alignment="center")
        elif is_num:
            column_config[c] = cc.NumberColumn(cname, format="%d", help=None, alignment="center")
        elif is_date:
            column_config[c] = cc.TextColumn(cname, help=None, alignment="center")
        else:
            # 文本列
            column_config[c] = cc.TextColumn(cname, help=None, alignment="center")

    # 构造 kwargs
    df_kwargs = {"use_container_width": True, "hide_index": True,
                 "column_config": column_config}
    # 过滤掉已隐藏的列
    if cols_to_hide:
        display_df = display_df.drop(columns=list(cols_to_hide))

    # 用 pandas Styler 强制所有单元格 + 表头居中（CSS 无法穿透 Streamlit 的 iframe）
    display_df = display_df.style.set_properties(
        **{"text-align": "center"}
    ).set_table_styles([
        {"selector": "thead th",
         "props": [("text-align", "center"), ("white-space", "nowrap")]},
        {"selector": "thead tr th",
         "props": [("text-align", "center")]},
    ])

    df_kwargs.update(kwargs)
    st.dataframe(display_df, **df_kwargs)


# ── Excel 导出 ──

def excel_export(df: pd.DataFrame, sheet_name: str = "Sheet1") -> bytes:
    """将 DataFrame 导出为 Excel 字节流（单 sheet）"""
    output = io.BytesIO()
    with pd.ExcelWriter(output, engine="openpyxl") as writer:
        df.to_excel(writer, index=False, sheet_name=sheet_name)
    return output.getvalue()


def apply_excel_borders(ws, min_row: int, max_row: int, min_col: int, max_col: int) -> None:
    """给指定区域添加细边框"""
    thin = Side(style="thin")
    border = Border(left=thin, right=thin, top=thin, bottom=thin)
    for row in ws.iter_rows(min_row=min_row, max_row=max_row, min_col=min_col, max_col=max_col):
        for cell in row:
            cell.border = border


def excel_export_by_class(df: pd.DataFrame, class_col: str = "班级", file_prefix: str = "导出") -> bytes:
    """
    按班级分 sheet 导出 Excel。

    Args:
        df: 要导出的 DataFrame
        class_col: 用于分 sheet 的列名
        file_prefix: 文件名前缀

    Returns:
        Excel 字节流
    """
    output = io.BytesIO()
    with pd.ExcelWriter(output, engine="openpyxl") as writer:
        if class_col in df.columns:
            classes = sorted(df[class_col].dropna().unique())
            if len(classes) <= 1:
                df.to_excel(writer, index=False, sheet_name="全部班级")
            else:
                for cls in classes:
                    cls_df = df[df[class_col] == cls]
                    sheet_name = str(cls)[:31]
                    cls_df.to_excel(writer, index=False, sheet_name=sheet_name)
        else:
            df.to_excel(writer, index=False, sheet_name="数据")
    return output.getvalue()
