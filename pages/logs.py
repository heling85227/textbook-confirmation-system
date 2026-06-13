"""
系统日志页面
============

查看导入操作记录和错误详情。
"""

import streamlit as st
import pandas as pd
from datetime import datetime, date

from database import query_df, execute_sql
from components import show_header, excel_export
from utils import (
    get_filtered_list, get_filtered_colleges, get_filtered_majors,
    safe_int, safe_float, safe_str, safe_field, read_import_logs
)


def system_logs_page():
    show_header("📋 系统日志", "查看导入操作记录和错误详情")

    col1, col2 = st.columns([0.7, 0.3])
    with col2:
        if st.button("🔄 刷新日志", use_container_width=True):
            st.rerun()

    logs = read_import_logs(50)
    if not logs or logs == ["暂无日志记录"]:
        st.info("暂无日志记录")
        return

    for log_text in logs:
        st.code(log_text.strip(), language="text")

    st.caption("日志文件位置：logs/import.log")
