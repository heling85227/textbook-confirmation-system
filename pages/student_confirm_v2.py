"""
学生端 — 核对确认与签名页面 (V2.0)
====================================

功能：
- Region 0: 通知中心（未读通知列表 + 红色徽章）
- Step 1: 核对区（查看本学期教材，逐条确认或提出异议）
- Step 2: 反馈区（选择反馈类型 + 填写详情）
- Step 3: 签名确认区（手写签名确认）

状态机：
    pending → confirmed（学生确认无误）
    pending → disputed（学生有异议）
    disputed → pending（管理员接受反馈，需重新确认）
    disputed → confirmed（管理员驳回反馈）
"""

import streamlit as st
import pandas as pd
import base64
import io
from datetime import datetime

from database import query_df, execute_sql, PRICE_CALC
from database_v2 import (
    upsert_confirmation, get_confirmations, get_confirmation_stats,
    add_notification, get_notifications, get_unread_count,
    mark_notification_read, mark_all_notifications_read,
    FEEDBACK_TYPES, CONFIRMATION_STATUS,
)
from components import show_header, styled_dataframe


def student_confirm_page():
    student = st.session_state.user
    show_header("✅ 教材核对确认", "核对您的教材信息并签名确认")

    student_id = int(student["id"])

    # ── 获取有教材的学期 ──
    sem_df = query_df("""
        SELECT DISTINCT sem.id, sem.name
        FROM textbook_orders o
        JOIN students s ON s.class_name = o.class_name
        JOIN semesters sem ON o.semester_id = sem.id
        WHERE s.id = %s
        ORDER BY sem.id DESC
    """, (student_id,))

    if sem_df.empty:
        st.info("📭 暂无教材数据，请联系管理员")
        return

    # ── 学期选择 ──
    semester = st.selectbox(
        "选择学期",
        [(r["id"], r["name"]) for _, r in sem_df.iterrows()],
        format_func=lambda x: x[1],
        key="sc_sem",
    )
    if not semester:
        return
    sem_id, sem_name = semester

    # ── 通知中心 ──
    _render_notification_center(student_id, sem_id)

    st.divider()

    # ── 确认状态概览 ──
    confs = get_confirmations(student_id=student_id, semester_id=sem_id)

    # 如果没有确认记录，自动创建
    if confs.empty:
        _auto_create_confirmations(student_id, sem_id)
        confs = get_confirmations(student_id=student_id, semester_id=sem_id)

    if confs.empty:
        st.info("📭 本学期暂无教材需要核对")
        return

    # 统计
    total = len(confs)
    confirmed = len(confs[confs["status"] == "confirmed"])
    disputed = len(confs[confs["status"] == "disputed"])
    pending = len(confs[confs["status"] == "pending"])

    c1, c2, c3, c4 = st.columns(4)
    c1.metric("📋 总教材数", total)
    c2.metric("⏳ 待核对", pending)
    c3.metric("✅ 已确认", confirmed)
    c4.metric("⚠️ 有异议", disputed)

    # 检查是否全部已确认（可签名）
    all_confirmed = (pending == 0 and disputed == 0 and total > 0)

    st.divider()

    # ── 核对确认主区域 ──
    st.markdown("### 📋 教材核对清单")
    st.caption("请逐条核对您的教材信息，确认无误请点击「✅ 确认」，如有异议请点击「⚠️ 有异议」")

    for _, row in confs.iterrows():
        _render_confirmation_row(student_id, sem_id, row)

    st.divider()

    # ── 签名确认区 ──
    _render_signature_section(student_id, sem_id, sem_name, all_confirmed, total, confirmed, disputed)


# ═════════════════════════════════════════════════════════
# 通知中心
# ═════════════════════════════════════════════════════════

def _render_notification_center(student_id, sem_id):
    """渲染通知中心"""
    unread = get_unread_count(student_id, sem_id)

    with st.expander(f"🔔 通知中心" + (f"（{unread} 条未读）" if unread > 0 else ""), expanded=(unread > 0)):
        if unread > 0:
            st.caption(f"您有 **{unread}** 条未读通知")

            notifications = get_notifications(student_id, sem_id, unread_only=True)

            if not notifications.empty:
                for _, n in notifications.iterrows():
                    ntype = n.get("type", "")
                    icon = {
                        "confirmation_required": "📋",
                        "feedback_processed": "✅",
                        "feedback_rejected": "❌",
                        "new_semester": "📅",
                        "system": "🔧",
                    }.get(ntype, "📌")

                    title = n.get("title", "")
                    content = n.get("content", "")
                    created = n.get("created_at", "")

                    st.markdown(f"""
                    <div style="background:#eff6ff; border-left:3px solid #3b82f6;
                                padding:8px 12px; border-radius:6px; margin-bottom:8px;">
                        <strong>{icon} {title}</strong>
                        <div style="color:#4b5563; font-size:13px;">{content}</div>
                        <div style="color:#9ca3af; font-size:11px; margin-top:4px;">{created}</div>
                    </div>
                    """, unsafe_allow_html=True)

                if st.button("✅ 全部标为已读", key="mark_all_read"):
                    mark_all_notifications_read(student_id, sem_id)
                    st.success("已全部标为已读")
                    st.rerun()
            else:
                st.info("📭 暂无未读通知")
        else:
            st.info("📭 暂无未读通知")


# ═════════════════════════════════════════════════════════
# 核对行
# ═════════════════════════════════════════════════════════

def _render_confirmation_row(student_id, sem_id, row):
    """渲染单条教材核对行"""
    textbook_name = row.get("textbook_name", "—")
    status = row.get("status", "pending")
    feedback_type = row.get("feedback_type", "")
    feedback_detail = row.get("feedback_detail", "")
    admin_response = row.get("admin_response", "")
    conf_id = int(row["id"])

    # 状态图标
    status_icon = {"pending": "⏳", "confirmed": "✅", "disputed": "⚠️"}.get(status, "❓")
    status_label = {"pending": "待核对", "confirmed": "已确认", "disputed": "有异议"}.get(status, status)

    # 背景色
    bg_color = {"pending": "#f9fafb", "confirmed": "#ecfdf5", "disputed": "#fef2f2"}.get(status, "#f9fafb")
    border_color = {"pending": "#d1d5db", "confirmed": "#059669", "disputed": "#dc2626"}.get(status, "#d1d5db")

    with st.container():
        st.markdown(f"""
        <div style="background:{bg_color}; border-left:4px solid {border_color};
                    padding:10px 16px; border-radius:8px; margin-bottom:8px;">
            <div style="display:flex; justify-content:space-between; align-items:center;">
                <div>
                    {status_icon} <strong>{textbook_name}</strong>
                    <span style="background:{'#dcfce7' if status=='confirmed' else '#fef3c7' if status=='disputed' else '#f3f4f6'};
                                 padding:2px 8px; border-radius:10px; font-size:12px; margin-left:8px;">
                        {status_label}
                    </span>
                </div>
            </div>
            {"<div style='color:#dc2626; font-size:13px; margin-top:4px;'>📝 反馈：" + feedback_type + ("：" + feedback_detail if feedback_detail else "") + "</div>" if status == "disputed" and feedback_type else ""}
            {"<div style='color:#059669; font-size:13px; margin-top:4px;'>💬 管理员回复：" + admin_response + "</div>" if admin_response else ""}
        </div>
        """, unsafe_allow_html=True)

        # 操作按钮（仅 pending 状态可操作）
        if status == "pending":
            col1, col2, col3 = st.columns([1, 1, 3])
            with col1:
                if st.button("✅ 确认", key=f"cfm_{conf_id}", type="primary", use_container_width=True):
                    upsert_confirmation(student_id, sem_id,
                                        textbook_id=row.get("textbook_id"),
                                        status="confirmed")
                    st.success("已确认")
                    st.rerun()
            with col2:
                if st.button("⚠️ 有异议", key=f"dsp_{conf_id}", use_container_width=True):
                    st.session_state[f"dispute_{conf_id}"] = True
                    st.rerun()

        # 异议表单
        if st.session_state.get(f"dispute_{conf_id}"):
            with st.form(f"dispute_form_{conf_id}"):
                fb_type = st.selectbox("反馈类型", FEEDBACK_TYPES, key=f"fbt_{conf_id}")
                fb_detail = st.text_area("详细说明", placeholder="请描述您的异议，如：少领1本、教材名称不对等",
                                          key=f"fbd_{conf_id}")
                col_a, col_b = st.columns(2)
                with col_a:
                    submitted = st.form_submit_button("提交异议", type="primary", use_container_width=True)
                with col_b:
                    cancelled = st.form_submit_button("取消", use_container_width=True)
                if cancelled:
                    st.session_state.pop(f"dispute_{conf_id}", None)
                    st.rerun()
                if submitted:
                    detail = fb_detail.strip() if fb_detail else fb_type
                    upsert_confirmation(student_id, sem_id,
                                        textbook_id=row.get("textbook_id"),
                                        status="disputed",
                                        feedback_type=fb_type,
                                        feedback_detail=detail)
                    # 通知管理员（通过系统日志体现）
                    st.session_state.pop(f"dispute_{conf_id}", None)
                    st.warning("异议已提交，等待管理员处理")
                    st.rerun()

        # disputed 状态 — 允许学生撤回异议（回到 pending）
        if status == "disputed" and not admin_response:
            col_r1, col_r2 = st.columns([1, 3])
            with col_r1:
                if st.button("↩️ 撤回异议", key=f"wd_{conf_id}", use_container_width=True):
                    upsert_confirmation(student_id, sem_id,
                                        textbook_id=row.get("textbook_id"),
                                        status="pending",
                                        feedback_type="",
                                        feedback_detail="")
                    st.info("已撤回异议，可重新核对")
                    st.rerun()

        # disputed 状态且管理员已接受（需重新确认）
        if status == "disputed" and admin_response:
            if st.button(f"🔄 重新确认", key=f"recfm_{conf_id}", use_container_width=True):
                upsert_confirmation(student_id, sem_id,
                                    textbook_id=row.get("textbook_id"),
                                    status="confirmed")
                st.success("已重新确认")
                st.rerun()


# ═════════════════════════════════════════════════════════
# 签名确认区
# ═════════════════════════════════════════════════════════

def _render_signature_section(student_id, sem_id, sem_name, all_confirmed, total, confirmed, disputed):
    """渲染签名确认区"""
    st.markdown("### ✍️ 签名确认")

    # 检查是否已有签名
    existing_sig = query_df(
        "SELECT id, signature_data, signed_at FROM signatures WHERE student_id = %s AND semester_id = %s",
        (student_id, sem_id)
    )

    if not existing_sig.empty:
        sig_data = existing_sig.iloc[0]["signature_data"]
        signed_at = existing_sig.iloc[0]["signed_at"]
        st.success(f"✅ 您已于 {signed_at} 完成签名确认")

        if sig_data:
            st.markdown("##### 已保存的签名")
            st.markdown(
                f'<img src="data:image/png;base64,{sig_data}" style="border:1px solid #ddd; border-radius:4px; max-width:400px; background:#fff;" />',
                unsafe_allow_html=True,
            )

        if st.button("🔄 重新签名", key="resign_v2"):
            execute_sql(
                "DELETE FROM signatures WHERE student_id = %s AND semester_id = %s",
                (student_id, sem_id)
            )
            st.rerun()
        return

    # 尚未签名
    if not all_confirmed:
        pending_count = total - confirmed - disputed
        st.warning(f"⚠️ 请先完成所有教材的核对确认（还有 {pending_count} 条待核对）后再签名")
        return

    # 全部已确认，可以签名
    st.caption("所有教材已确认无误，请在下方签名板上手写签名，完成最终确认")

    from streamlit_drawable_canvas import st_canvas

    canvas_result = st_canvas(
        stroke_width=2,
        stroke_color="#000",
        background_color="#fff",
        height=300,
        width=600,
        drawing_mode="freedraw",
        display_toolbar=False,
        key=f"sig_cvs_v2_{sem_id}",
    )

    if st.button("✅ 确认签名", type="primary", key=f"sv_v2_{sem_id}", use_container_width=True):
        if canvas_result is None or canvas_result.image_data is None:
            st.warning("画板加载中，请稍后")
        else:
            from PIL import Image
            img = Image.fromarray(canvas_result.image_data.astype("uint8"), mode="RGBA").convert("RGB")
            buf = io.BytesIO()
            img.save(buf, format="PNG")
            b64 = base64.b64encode(buf.getvalue()).decode("utf-8")
            rgb = canvas_result.image_data[:, :, :3]
            if (rgb == 255).all():
                st.warning("⚠️ 请先在画板上签名")
            else:
                now = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
                execute_sql(
                    "INSERT OR REPLACE INTO signatures (student_id, semester_id, signature_data, signed_at) VALUES (%s, %s, %s, %s)",
                    (student_id, sem_id, b64, now)
                )
                st.success("✅ 签名保存成功！")
                st.rerun()


# ═════════════════════════════════════════════════════════
# 辅助函数
# ═════════════════════════════════════════════════════════

def _auto_create_confirmations(student_id, sem_id):
    """自动为该学生创建本学期的待确认记录"""
    # 查找该学生本学期应领取的教材
    books = query_df("""
        SELECT DISTINCT o.textbook_id
        FROM textbook_orders o
        JOIN students s ON s.class_name = o.class_name
        WHERE o.semester_id = %s AND s.id = %s
    """, (sem_id, student_id))

    if books.empty:
        return

    for _, row in books.iterrows():
        tid = int(row["textbook_id"])
        # 检查是否已存在
        existing = query_df(
            """SELECT id FROM student_confirmations
               WHERE student_id = %s AND semester_id = %s AND textbook_id = %s""",
            (student_id, sem_id, tid)
        )
        if existing.empty:
            upsert_confirmation(student_id, sem_id, tid, status="pending")
