"""
管理端 — 学生反馈处理页面 (V3.0)
====================================

功能：
- 待处理反馈卡片式展示（接受/驳回）
- 全部记录按学生分组，可展开查看明细
- 按学期、状态、反馈类型筛选
- 批量处理操作
"""

import streamlit as st
import pandas as pd
from database import query_df, execute_sql, PRICE_CALC, PRICE_JOIN
from database_v2 import (
    get_confirmations, process_feedback, get_confirmation_stats,
    FEEDBACK_TYPES, CONFIRMATION_STATUS
)
from components import show_header, styled_dataframe
from utils import get_filtered_colleges, get_filtered_majors, get_filtered_class_names


def feedback_page():
    show_header("📢 学生反馈处理", "查看和处理学生的教材核对反馈")

    # ── 学期选择 ──
    semesters = query_df("SELECT id, name FROM semesters ORDER BY id DESC")
    if semesters.empty:
        st.warning("请先添加学期")
        return

    semester = st.selectbox(
        "选择学期",
        [(r["id"], r["name"]) for _, r in semesters.iterrows()],
        format_func=lambda x: x[1],
        key="fb_sem",
    )
    if not semester:
        return
    sem_id, sem_name = semester

    # ── 统计概览 ──
    stats = get_confirmation_stats(sem_id)
    if not stats.empty:
        s = stats.iloc[0]
        total = int(s.get("total", 0) or 0)
        pending = int(s.get("pending_count", 0) or 0)
        confirmed = int(s.get("confirmed_count", 0) or 0)
        disputed = int(s.get("disputed_count", 0) or 0)

        c1, c2, c3, c4 = st.columns(4)
        c1.metric("📋 总记录", total)
        c2.metric("⏳ 待确认", pending)
        c3.metric("✅ 已确认", confirmed)
        c4.metric("⚠️ 有异议", disputed)

    st.divider()

    # ── Tab 切换 ──
    tab_pending, tab_all = st.tabs(["⚠️ 待处理反馈", "📋 全部确认记录"])

    # ========== Tab 1: 待处理反馈 ==========
    with tab_pending:
        disputed_df = get_confirmations(semester_id=sem_id, status="disputed")

        if disputed_df.empty:
            st.success("✅ 当前没有待处理的反馈")
        else:
            st.caption(f"共 {len(disputed_df)} 条待处理反馈")

            # 反馈类型筛选
            fb_types = ["全部"] + FEEDBACK_TYPES
            fb_filter = st.selectbox("按反馈类型筛选", fb_types, key="fb_type_filter")

            if fb_filter != "全部":
                disputed_df = disputed_df[disputed_df["feedback_type"] == fb_filter]

            for idx, row in disputed_df.iterrows():
                _render_feedback_card(row, key_prefix="pending")

    # ========== Tab 2: 全部记录（按学生分组） ==========
    with tab_all:
        _render_grouped_records(sem_id, sem_name)


# ═════════════════════════════════════════════════════════
# 反馈卡片（待处理）
# ═════════════════════════════════════════════════════════

def _render_feedback_card(row, key_prefix="fb"):
    """渲染单条待处理反馈卡片"""
    student_name = row.get("student_name", "未知")
    sid = row.get("sid", "")
    class_name = row.get("class_name", "")
    textbook_name = row.get("textbook_name", "—")
    feedback_type = row.get("feedback_type", "")
    feedback_detail = row.get("feedback_detail", "")
    updated_at = row.get("updated_at", "")

    with st.container():
        st.markdown(f"""
        <div style="background:#fffbeb; border-left:4px solid #f59e0b; padding:12px 16px;
                    border-radius:8px; margin-bottom:12px;">
            <div style="display:flex; justify-content:space-between; align-items:center;">
                <div>
                    <strong>👤 {student_name}</strong>
                    <span style="color:#6b7280; margin-left:8px;">学号: {sid} | 班级: {class_name}</span>
                </div>
                <span style="background:#fef3c7; color:#92400e; padding:2px 10px;
                             border-radius:12px; font-size:13px;">{feedback_type}</span>
            </div>
            <div style="margin-top:8px; color:#374151;">
                📖 教材：<strong>{textbook_name}</strong>
            </div>
            <div style="margin-top:4px; color:#4b5563;">
                💬 {feedback_detail if feedback_detail else '（无详细说明）'}
            </div>
            <div style="margin-top:4px; color:#9ca3af; font-size:12px;">
                🕐 {updated_at}
            </div>
        </div>
        """, unsafe_allow_html=True)

        col1, col2, col3 = st.columns([1, 1, 3])
        with col1:
            if st.button("✅ 接受", key=f"{key_prefix}_accept_{row['id']}", type="primary",
                         use_container_width=True):
                ok, msg = process_feedback(row["id"], "accept", "已处理")
                if ok:
                    st.success(msg)
                    st.rerun()
                else:
                    st.error(msg)
        with col2:
            if st.button("❌ 驳回", key=f"{key_prefix}_{row['id']}",
                         use_container_width=True):
                st.session_state[f"reject_{row['id']}"] = True
                st.rerun()

        # 驳回原因输入
        if st.session_state.get(f"reject_{row['id']}"):
            with st.form(f"reject_form_{row['id']}"):
                reject_reason = st.text_input("驳回原因", placeholder="请填写驳回原因")
                col_a, col_b = st.columns(2)
                with col_a:
                    submitted = st.form_submit_button("确认驳回", type="primary", use_container_width=True)
                with col_b:
                    cancelled = st.form_submit_button("取消", use_container_width=True)
                if cancelled:
                    st.session_state.pop(f"reject_{row['id']}", None)
                    st.rerun()
                if submitted:
                    reason = reject_reason.strip() or "不符合反馈条件"
                    ok, msg = process_feedback(row["id"], "reject", reason)
                    if ok:
                        st.session_state.pop(f"reject_{row['id']}", None)
                        st.success(msg)
                        st.rerun()
                    else:
                        st.error(msg)


# ═════════════════════════════════════════════════════════
# 全部记录（按学生分组）
# ═════════════════════════════════════════════════════════

def _render_grouped_records(sem_id, sem_name):
    """按学生分组展示所有确认记录"""

    all_df = get_confirmations(semester_id=sem_id)

    if all_df.empty:
        st.info("📭 暂无确认记录")
        return

    # 筛选栏
    fcol1, fcol2, fcol3 = st.columns([1.5, 1.5, 2])
    with fcol1:
        status_filter = st.selectbox(
            "状态",
            ["全部", "pending", "confirmed", "disputed"],
            format_func=lambda x: {
                "全部": "全部状态", "pending": "⏳ 待确认",
                "confirmed": "✅ 已确认", "disputed": "⚠️ 有异议"
            }.get(x, x),
            key="grp_status",
        )
    with fcol2:
        type_filter = st.selectbox(
            "反馈类型",
            ["全部"] + FEEDBACK_TYPES,
            key="grp_type",
        )
    with fcol3:
        search_kw = st.text_input("🔍 搜索学生/学号/班级", placeholder="输入关键字...", key="grp_search")

    # 应用筛选
    filtered = all_df.copy()
    if status_filter != "全部":
        filtered = filtered[filtered["status"] == status_filter]
    if type_filter != "全部":
        filtered = filtered[filtered["feedback_type"] == type_filter]
    if search_kw.strip():
        kw = search_kw.strip().lower()
        mask = (
            filtered["student_name"].fillna("").str.lower().str.contains(kw, na=False) |
            filtered["sid"].fillna("").str.lower().str.contains(kw, na=False) |
            filtered["class_name"].fillna("").str.lower().str.contains(kw, na=False)
        )
        filtered = filtered[mask]

    st.caption(f"共 **{len(filtered)}** 条记录 / 涉及 **{filtered['student_name'].nunique() if not filtered.empty else 0}** 名学生")

    if filtered.empty:
        st.info("📭 无匹配记录")
        return

    # 按学生分组
    grouped = list(filtered.groupby("student_name", sort=False))

    # 展示每个学生
    for student_name, stu_df in grouped:
        first = stu_df.iloc[0]
        sid = first.get("sid", "")
        class_name = first.get("class_name", "")

        # 统计该学生的状态
        stu_total = len(stu_df)
        stu_confirmed = len(stu_df[stu_df["status"] == "confirmed"])
        stu_pending = len(stu_df[stu_df["status"] == "pending"])
        stu_disputed = len(stu_df[stu_df["status"] == "disputed"])

        # 根据状态决定颜色
        if stu_disputed > 0:
            card_bg = "#fef2f2"
            border_color = "#dc2626"
            status_badge = f"⚠️ {stu_disputed} 条异议"
            badge_bg = "#fee2e2"
            badge_color = "#991b1b"
        elif stu_pending > 0:
            card_bg = "#fffbeb"
            border_color = "#f59e0b"
            status_badge = f"⏳ {stu_pending} 条待确认"
            badge_bg = "#fef3c7"
            badge_color = "#92400e"
        else:
            card_bg = "#f0fdf4"
            border_color = "#059669"
            status_badge = f"✅ {stu_confirmed} 条已确认"
            badge_bg = "#dcfce7"
            badge_color = "#166534"

        with st.expander(f"**{student_name}**  |  {sid}  |  {class_name}  |  共 {stu_total} 本教材", expanded=(stu_disputed > 0)):
            # 教材明细表
            detail_data = []
            for _, r in stu_df.iterrows():
                status = r.get("status", "")
                fb_t = r.get("feedback_type", "")
                fb_d = r.get("feedback_detail", "")
                admin_resp = r.get("admin_response", "")
                admin_act = r.get("admin_action", "")
                updated = r.get("updated_at", "")

                # 状态标签
                status_tag = {
                    "pending": ("⏳ 待核对", "#fef3c7"),
                    "confirmed": ("✅ 已确认", "#dcfce7"),
                    "disputed": ("⚠️ 有异议", "#fee2e2"),
                }.get(status, (status, "#f3f4f6"))

                detail_data.append({
                    "📖 教材名称": r.get("textbook_name", "—"),
                    "📌 状态": status_tag[0],
                    "💬 反馈类型": fb_t if fb_t else "—",
                    "📝 反馈详情": fb_d if fb_d else "—",
                    "💡 管理员回复": admin_resp if admin_resp else "—",
                    "🔧 处理动作": admin_act if admin_act else "—",
                    "🕐 时间": str(updated)[:16] if updated else "—",
                })

            detail_df = pd.DataFrame(detail_data)
            st.dataframe(detail_df, use_container_width=True, hide_index=True)

            # 快捷操作区（仅对该学生的有异议项显示处理按钮）
            disputed_items = stu_df[stu_df["status"] == "disputed"]
            if not disputed_items.empty:
                st.markdown("---")
                cols = st.columns(min(3, len(disputed_items)))
                for i, (_, d_row) in enumerate(disputed_items.iterrows()):
                    with cols[i % 3]:
                        bk_name = d_row.get("textbook_name", "")[:12]
                        c_op1, c_op2 = st.columns(2)
                        with c_op1:
                            if st.button(f"✅ 接受", key=f"gaccept_{d_row['id']}", use_container_width=True, type="primary"):
                                ok, msg = process_feedback(d_row["id"], "accept", "已接受")
                                if ok:
                                    st.success(msg); st.rerun()
                                else:
                                    st.error(msg)
                        with c_op2:
                            if st.button(f"❌ 驳回", key=f"greject_{d_row['id']}", use_container_width=True):
                                st.session_state[f"greject_{d_row['id']}"] = True
                                st.rerun()

                        # 内联驳回
                        if st.session_state.get(f"greject_{d_row['id']}"):
                            rr = st.text_input("驳回原因", key=f"grr_{d_row['id']}")
                            gok, gcancel = st.columns(2)
                            with gok:
                                if st.button("确认", key=f"grr_ok_{d_row['id']}", use_container_width=True):
                                    reason = rr.strip() or "不符合条件"
                                    ok, msg = process_feedback(d_row["id"], "reject", reason)
                                    if ok:
                                        st.session_state.pop(f"greject_{d_row['id']}", None)
                                        st.success(msg); st.rerun()
                                    else:
                                        st.error(msg)
                            with gcancel:
                                if st.button("取消", key=f"grr_cancel_{d_row['id']}", use_container_width=True):
                                    st.session_state.pop(f"greject_{d_row['id']}", None)
                                    st.rerun()

    # 导出功能
    st.divider()
    from components import excel_export

    export_df = filtered.rename(columns={
        "student_name": "学生姓名",
        "sid": "学号",
        "class_name": "班级",
        "textbook_name": "教材名称",
        "status": "状态",
        "feedback_type": "反馈类型",
        "feedback_detail": "反馈详情",
        "admin_response": "管理员回复",
        "admin_action": "处理动作",
        "updated_at": "更新时间",
    })
    excel_data = excel_export(export_df, "确认记录")
    st.download_button(
        "📥 导出当前记录",
        data=excel_data,
        file_name=f"教材确认记录_{sem_name}.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    )
