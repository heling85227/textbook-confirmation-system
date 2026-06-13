"""
学期管理页面
==============

学期管理功能，包括新增学期、生成测试数据、查看和删除学期。
"""

import streamlit as st
from datetime import date

from components import show_header
from utils import get_current_academic_info
from test_data import generate_test_data
from database import query_df, execute_sql


def semester_management():
    show_header("📅 学期管理", "管理学年和学期，系统会自动识别当前学期")

    ay, sem = get_current_academic_info()

    col1, col2, col3, col4 = st.columns([2, 1, 1, 1])
    with col1:
        st.info(f"🔍 系统自动识别当前：**{ay} 学年 {sem}**")
    with col2:
        if st.button("➕ 新增学期", use_container_width=True, type="primary"):
            st.session_state.show_semester_form = True
    with col3:
        if st.button("🎲 生成测试数据", use_container_width=True, type="secondary"):
            st.session_state.show_gen_confirm = True
    with col4:
        if st.button("🔄 刷新", use_container_width=True):
            st.rerun()

    # 生成测试数据确认
    if st.session_state.get("show_gen_confirm"):
        with st.expander("🎲 生成测试数据", expanded=True):
            st.warning("⚠️ 将生成完整的测试数据（学期、学生、教材、发放记录），已有同名数据会自动跳过。确定要生成吗？")
            c1, c2 = st.columns(2)
            with c1:
                if st.button("✅ 确认生成", use_container_width=True, type="primary"):
                    with st.spinner("正在生成测试数据..."):
                        n_stu, n_tb, n_sem = generate_test_data()
                    st.success(f"✅ 测试数据生成完成！学生 {n_stu} 条，教材 {n_tb} 条，涉及 {n_sem} 个学期")
                    st.session_state.show_gen_confirm = False
                    st.rerun()
            with c2:
                if st.button("取消", use_container_width=True):
                    st.session_state.show_gen_confirm = False
                    st.rerun()

    if st.session_state.get("show_semester_form"):
        with st.expander("✨ 新增学期", expanded=True):
            with st.form("semester_form_inner"):
                c1, c2, c3 = st.columns([2, 1, 1])
                with c1:
                    name = st.text_input("学期名称", placeholder=f"如：{ay} {sem}")
                with c2:
                    ac_year = st.text_input("学年", value=ay)
                with c3:
                    sem_name = st.selectbox("学期", ["第一学期", "第二学期"],
                                            index=1 if sem == "第二学期" else 0)

                col_a, col_b = st.columns(2)
                with col_a:
                    save_btn = st.form_submit_button("💾 保存", use_container_width=True)
                with col_b:
                    cancel_btn = st.form_submit_button("取消", use_container_width=True)

                if save_btn:
                    full_name = name or f"{ac_year} {sem_name}"
                    try:
                        execute_sql(
                            "INSERT INTO semesters (name, academic_year, semester_name) VALUES (%s, %s, %s)",
                            (full_name, ac_year, sem_name)
                        )
                        st.success(f"✅ 已添加学期：{full_name}")
                        st.session_state.show_semester_form = False
                        st.rerun()
                    except Exception as e:
                        if "UNIQUE" in str(e) or "unique" in str(e).lower():
                            st.error("❌ 该学期已存在")
                        else:
                            st.error(f"保存失败：{e}")
                if cancel_btn:
                    st.session_state.show_semester_form = False
                    st.rerun()

    df = query_df("SELECT * FROM semesters ORDER BY id DESC")
    if not df.empty:
        st.markdown("#### 📋 已有学期")
        for _, row in df.iterrows():
            with st.container():
                col1, col2, col3, col4 = st.columns([3, 1, 1, 0.8])
                with col1:
                    st.write(f"**{row['name']}**")
                with col2:
                    st.caption(f"📅 {row.get('academic_year', '-')}")
                with col3:
                    st.caption(f"📝 {row.get('semester_name', '-')}")
                with col4:
                    if st.button("🗑️", key=f"del_sem_{row['id']}", help="删除此学期"):
                        execute_sql("DELETE FROM semesters WHERE id = %s", (row['id'],))
                        st.warning(f"已删除学期「{row['name']}」")
                        st.rerun()
        st.divider()
    else:
        st.warning("⚠️ 暂无学期数据，请先点击「新增学期」添加")
