"""
费用统计页面
============

多维度汇总学生教材费用，支持按学期分列展示，可导出 Excel。
"""

import streamlit as st
import pandas as pd
from datetime import date

from database import query_df, execute_sql, PRICE_CALC, PRICE_JOIN
from components import show_header, excel_export, styled_dataframe
from utils import (
    get_filtered_list, get_filtered_class_names, get_filtered_colleges,
    get_filtered_majors, safe_int, safe_float, safe_str, safe_field,
    read_import_logs
)


def statistics_page():
    show_header("📊 费用统计", "多维度汇总学生教材费用，支持按学期分列展示，可导出 Excel")

    semesters = query_df("SELECT id, name FROM semesters ORDER BY id DESC")
    if semesters.empty:
        st.warning("⚠️ 请先添加学期")
        return

    class_names = get_filtered_class_names()

    # 筛选条件
    with st.expander("🔍 筛选条件", expanded=True):
        col1, col2, col3, col4 = st.columns(4)
        with col1:
            sel_semester_names = st.multiselect(
                "选择学期（可多选）",
                options=list(semesters["name"]),
                default=list(semesters["name"])[:1] if not semesters.empty else [],
                format_func=lambda x: x,
                placeholder="请选择学期"
            )
        with col2:
            stat_college = st.selectbox("学院", ["全部"] + get_filtered_colleges(), key="stat_college")
        with col3:
            if stat_college != "全部":
                stat_major_opts = ["全部"] + get_filtered_list("students", "major", "college = %s", (stat_college,))
            else:
                stat_major_opts = ["全部"] + get_filtered_majors()
            stat_major = st.selectbox("专业", stat_major_opts, key="stat_major")
        with col4:
            stat_class_where = "1=1"
            stat_class_params = []
            if stat_college != "全部":
                stat_class_where += " AND college = %s"
                stat_class_params.append(stat_college)
            if stat_major != "全部":
                stat_class_where += " AND major = %s"
                stat_class_params.append(stat_major)
            stat_class_opts = ["全部"] + (
                get_filtered_list("students", "class_name", stat_class_where, tuple(stat_class_params))
                if stat_class_params else class_names
            )
            sel_classes = st.multiselect(
                "选择班级（可多选）",
                options=stat_class_opts,
                default=[],
                format_func=lambda x: x,
                placeholder="请选择班级"
            )

    # ── 单学生查询 ──
    with st.expander("🔍 单学生查询", expanded=(st.session_state.get("stat_search_type", "不限（显示全部）") != "不限（显示全部）")):
        col_search1, col_search2 = st.columns(2)
        with col_search1:
            search_type = st.selectbox("学生查询方式", ["不限（显示全部）", "按姓名", "按学号", "按身份证"], key="stat_search_type")
        search_value = ""
        if search_type != "不限（显示全部）":
            with col_search2:
                search_value = st.text_input(
                    f"输入{'姓名' if '姓名' in search_type else '学号' if '学号' in search_type else '身份证号'}",
                    key="stat_search_value"
                ).strip()

    col_group = st.columns(1)[0]
    with col_group:
        group_by = st.selectbox("汇总维度", ["按学生", "按班级", "按专业", "按年级", "按学院"], key="stat_group")

    if not sel_semester_names:
        st.info("📭 请至少选择一个学期")
        return

    sem_id_map = {r["name"]: r["id"] for _, r in semesters.iterrows()}
    sel_sem_ids = [sem_id_map[n] for n in sel_semester_names]

    # 构建查询：获取明细数据
    ph = ",".join(["%s"] * len(sel_sem_ids))
    sql = f"""
            SELECT s.id as student_id_pk, s.student_id, s.name as student_name,
                   s.grade, s.college, s.major, s.class_name, s.id_card,
                   t.name as textbook_name,
                   {PRICE_CALC} as calc_price,
                   d.quantity,
                   {PRICE_CALC} * d.quantity as subtotal,
                   sem.name as semester_name, sem.id as semester_id,
                   d.distribute_date
            FROM distributions d
            JOIN students s ON d.student_id = s.id
            JOIN textbooks t ON d.textbook_id = t.id
            JOIN semesters sem ON t.semester_id = sem.id
            {PRICE_JOIN}
            LEFT JOIN student_exemptions e ON e.semester_id = sem.id AND e.student_id = s.id
            WHERE t.semester_id IN ({ph})

              AND (e.id IS NULL OR e.is_exempt = 0)
        """

    params = list(sel_sem_ids)

    # 单学生搜索条件
    if search_type != "不限（显示全部）" and search_value:
        if search_type == "按姓名":
            sql += " AND s.name LIKE %s"
            params.append(f"%{search_value}%")
        elif search_type == "按学号":
            sql += " AND s.student_id LIKE %s"
            params.append(f"%{search_value}%")
        elif search_type == "按身份证":
            sql += " AND s.id_card LIKE %s"
            params.append(f"%{search_value}%")

    if stat_college != "全部":
        sql += " AND s.college = %s"
        params.append(stat_college)
    if stat_major != "全部":
        sql += " AND s.major = %s"
        params.append(stat_major)
    if sel_classes and "全部" not in sel_classes:
        placeholders_c = ",".join(["%s"] * len(sel_classes))
        sql += f" AND s.class_name IN ({placeholders_c})"
        params.extend(sel_classes)

    sql += " ORDER BY s.class_name, s.name, sem.id, t.name"

    raw_df = query_df(sql, tuple(params))

    if raw_df.empty:
        st.info("📭 暂无发放数据可统计")
        return

    raw_df["subtotal"] = pd.to_numeric(raw_df["subtotal"], errors="coerce").fillna(0)
    raw_df["calc_price"] = pd.to_numeric(raw_df["calc_price"], errors="coerce").fillna(0)

    # ── 按学生维度：生成透视表（学生 × 学期）──
    if group_by == "按学生":
        pivot = raw_df.groupby(["student_id_pk", "student_id", "student_name", "class_name", "grade", "major", "college", "semester_name"])["subtotal"].sum().reset_index()

        pivot_table = pivot.pivot_table(
            index=["student_id_pk", "student_id", "student_name", "class_name", "grade", "major", "college"],
            columns="semester_name",
            values="subtotal",
            aggfunc="sum",
            fill_value=0
        ).reset_index()

        sem_cols = [c for c in pivot_table.columns if c not in ["student_id_pk", "student_id", "student_name", "class_name", "grade", "major", "college"]]
        pivot_table["总计"] = pivot_table[sem_cols].sum(axis=1)
        pivot_table["总计"] = pivot_table["总计"].round(2)

        display_cols = ["学号", "姓名", "班级", "年级", "专业", "学院"] + sem_cols + ["总计"]
        result_df = pivot_table.rename(columns={
            "student_id": "学号",
            "student_name": "姓名",
            "class_name": "班级",
            "grade": "年级",
            "major": "专业",
            "college": "学院"
        })
        result_df = result_df[display_cols]

        total_fee = result_df["总计"].sum()
        record_count = len(result_df)
        avg_fee = total_fee / record_count if record_count > 0 else 0

        m1, m2, m3 = st.columns(3)
        with m1:
            st.metric("💰 结算总费用", f"¥{total_fee:,.2f}")
        with m2:
            st.metric("📊 学生数", f"{record_count}")
        with m3:
            st.metric("📈 人均结算", f"¥{avg_fee:,.2f}")

        styled_dataframe(result_df, hide_ids=True)

        excel_data = excel_export(result_df, "费用统计_按学生")
        st.download_button(
            "📥 导出费用统计",
            data=excel_data,
            file_name=f"费用统计_按学生_{date.today()}.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            use_container_width=True,
            type="primary"
        )

        with st.expander("📋 查看费用明细列表"):
            detail_df = raw_df[["student_id", "student_name", "class_name", "textbook_name",
                                "calc_price", "quantity", "subtotal", "semester_name", "distribute_date"]]
            detail_df = detail_df.rename(columns={
                "student_id": "学号", "student_name": "姓名", "class_name": "班级",
                "textbook_name": "教材名称",
                "calc_price": "结算单价", "quantity": "数量",
                "subtotal": "小计", "semester_name": "学期", "distribute_date": "发放日期"
            })
            styled_dataframe(detail_df, hide_ids=True)

    else:
        group_map = {
            "按班级": ["class_name"],
            "按专业": ["major"],
            "按年级": ["grade"],
            "按学院": ["college"],
        }
        gcols = group_map[group_by]

        pivot = raw_df.groupby(gcols + ["semester_name"])["subtotal"].sum().reset_index()

        pivot_table = pivot.pivot_table(
            index=gcols,
            columns="semester_name",
            values="subtotal",
            aggfunc="sum",
            fill_value=0
        ).reset_index()

        sem_cols = [c for c in pivot_table.columns if c not in gcols]
        pivot_table["总计"] = pivot_table[sem_cols].sum(axis=1)
        pivot_table["总计"] = pivot_table["总计"].round(2)

        rename_map = {
            "class_name": "班级",
            "major": "专业",
            "grade": "年级",
            "college": "学院"
        }
        result_df = pivot_table.rename(columns=rename_map)

        total_fee = result_df["总计"].sum()
        record_count = len(result_df)

        m1, m2 = st.columns(2)
        with m1:
            st.metric("💰 结算总费用", f"¥{total_fee:,.2f}")
        with m2:
            st.metric(f"📊 {group_by}条目", f"{record_count}")

        styled_dataframe(result_df, hide_ids=True)

        excel_data = excel_export(result_df, "费用统计")
        st.download_button(
            "📥 导出费用统计",
            data=excel_data,
            file_name=f"费用统计_{group_by}_{date.today()}.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            use_container_width=True,
            type="primary"
        )

        with st.expander("📋 查看费用明细列表"):
            detail_df = raw_df[["student_id", "student_name", "class_name", "textbook_name",
                                "calc_price", "quantity", "subtotal", "semester_name", "distribute_date"]]
            detail_df = detail_df.rename(columns={
                "student_id": "学号", "student_name": "姓名", "class_name": "班级",
                "textbook_name": "教材名称",
                "calc_price": "结算单价", "quantity": "数量",
                "subtotal": "小计", "semester_name": "学期", "distribute_date": "发放日期"
            })
            styled_dataframe(detail_df, hide_ids=True)
