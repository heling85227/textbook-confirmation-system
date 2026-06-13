"""
学生费用查询 + 领书确认页面
=============================

学生查看个人各学期教材费用明细与汇总，查看领书情况并手写签名确认。
"""

import streamlit as st
import pandas as pd
import base64
from datetime import date, datetime

from database import query_df, execute_sql, PRICE_CALC, PRICE_JOIN
from components import show_header, excel_export, styled_dataframe
from utils import (
    get_filtered_list, get_filtered_colleges, get_filtered_majors,
    safe_int, safe_float, safe_str, safe_field, read_import_logs
)


def student_query_page():
    student = st.session_state.user
    show_header("🎓 我的教材信息", f"查看个人教材费用与领书情况")

    st.markdown(f"""
    <div class="student-card">
        <div style="display:flex; gap:24px; align-items:center; flex-wrap:wrap;">
            <div style="text-align:center;">
                <div style="font-size:48px;">👤</div>
            </div>
            <div>
                <h2 style="margin:0 0 4px 0; color:#1e40af;">{student['name']}</h2>
                <p style="margin:0; color:#6b7280;">学号：{student['student_id']} &nbsp;|&nbsp; 班级：{student.get('class_name', '-')} &nbsp;|&nbsp; 专业：{student.get('major', '-')}</p>
                <p style="margin:4px 0 0 0; color:#9ca3af; font-size:13px;">学院：{student.get('college', '-')} &nbsp;|&nbsp; 年级：{student.get('grade', '-')}</p>
            </div>
        </div>
    </div>
    """, unsafe_allow_html=True)

    # 获取该学生的学期列表（有教材的学期）
    sem_df = query_df("""
        SELECT DISTINCT sem.id, sem.name
        FROM textbook_orders o
        JOIN students s ON s.class_name = o.class_name
        JOIN semesters sem ON o.semester_id = sem.id
        WHERE s.id = %s
        ORDER BY sem.id DESC
    """, (student["id"],))

    if sem_df.empty:
        st.info("📭 暂无教材数据，请联系管理员")
        return

    tab1, tab2, tab3 = st.tabs(["📊 费用明细", "📖 领书情况", "✍️ 领取签名确认"])

    # ========== Tab 1: 费用明细（保留原有） ==========
    with tab1:
        sql = f"""
            SELECT t.name as textbook_name, t.publisher, t.editor,
                   {PRICE_CALC} as calc_price,
                   d.quantity,
                   {PRICE_CALC} * d.quantity as subtotal,
                   sem.name as semester_name, d.distribute_date
            FROM distributions d
            JOIN textbooks t ON d.textbook_id = t.id
            JOIN semesters sem ON t.semester_id = sem.id
            {PRICE_JOIN}
            LEFT JOIN student_exemptions e ON e.semester_id = sem.id AND e.student_id = d.student_id
            WHERE d.student_id = %s
              AND (e.id IS NULL OR e.is_exempt = 0)
            ORDER BY sem.id, t.name
        """
        df = query_df(sql, (student["id"],))

        if df.empty:
            st.info("📭 暂无教材发放记录")
        else:
            df["subtotal"] = pd.to_numeric(df["subtotal"], errors="coerce").fillna(0)
            df["calc_price"] = pd.to_numeric(df["calc_price"], errors="coerce").fillna(0)

            total_fee = df["subtotal"].sum()
            book_count = len(df["textbook_name"].unique())

            m1, m2, m3 = st.columns(3)
            with m1:
                st.metric("📚 总费用", f"¥{total_fee:,.2f}")
            with m2:
                st.metric("📖 教材种类", f"{book_count} 种")
            with m3:
                st.metric("📝 记录条数", f"{len(df)} 条")

            st.divider()

            st.markdown("### 📊 按学期汇总")
            sem_summary = df.groupby("semester_name")["subtotal"].sum().reset_index()
            sem_summary.columns = ["学期", "费用合计"]
            sem_summary["费用合计"] = sem_summary["费用合计"].round(2)

            cols_sem = st.columns(len(sem_summary) if len(sem_summary) > 0 else 1)
            for i, (_, row) in enumerate(sem_summary.iterrows()):
                with cols_sem[i % len(cols_sem)]:
                    st.metric(row["学期"], f"¥{row['费用合计']:,.2f}")

            st.divider()

            st.markdown("### 📋 费用明细")
            detail_df = df.rename(columns={
                "textbook_name": "教材名称", "publisher": "出版社", "editor": "主编",
                "calc_price": "结算价(元)", "quantity": "数量", "subtotal": "小计(元)",
                "semester_name": "学期", "distribute_date": "发放日期"
            })
            styled_dataframe(detail_df, hide_ids=True)

            excel_data = excel_export(detail_df, "费用明细")
            st.download_button("📥 导出费用明细", data=excel_data,
                               file_name=f"教材费用_{student['student_id']}_{student['name']}.xlsx",
                               mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")

    # ========== Tab 2: 领书情况 ==========
    with tab2:
        st.markdown("### 📖 我的领书情况")

        # 按学期展示
        for _, sem_row in sem_df.iterrows():
            sem_id = sem_row["id"]
            sem_name = sem_row["name"]

            # 查询该学期该班级征订的所有教材
            books = query_df("""
                SELECT tm.name, tm.isbn, tm.publisher, tm.price
                FROM textbook_orders o
                JOIN textbooks_master tm ON o.textbook_id = tm.id
                JOIN students s ON s.class_name = o.class_name
                WHERE o.semester_id = %s AND s.id = %s
                ORDER BY tm.name
            """, (sem_id, student["id"]))

            if books.empty:
                continue

            # 查询该学期该学生已领取的教材ID
            dist = query_df("""
                SELECT DISTINCT t.name as tb_name, d.distribute_date
                FROM distributions d
                JOIN textbooks t ON d.textbook_id = t.id
                WHERE d.student_id = %s AND t.semester_id = %s
            """, (student["id"], sem_id))
            dist_names = set(r["tb_name"] for _, r in dist.iterrows())
            dist_dates = {r["tb_name"]: r["distribute_date"] for _, r in dist.iterrows()}

            # 检查是否已签名
            sig_exists = query_df(
                "SELECT id, signed_at FROM signatures WHERE student_id = %s AND semester_id = %s",
                (student["id"], sem_id)
            )

            with st.expander(f"📚 {sem_name}" + (" ✅ 已签名确认" if not sig_exists.empty else ""),
                             expanded=True):
                rows = []
                picked = 0
                for _, bk in books.iterrows():
                    name = bk["name"]
                    has_book = name in dist_names
                    if has_book:
                        picked += 1
                    rows.append({
                        "教材名称": name,
                        "ISBN": safe_str(bk.get("isbn", "")),
                        "领取状态": "✅ 已领" if has_book else "❌ 未领",
                        "领书日期": str(dist_dates.get(name, "")) if has_book else "-",
                    })

                status_df = pd.DataFrame(rows)
                st.dataframe(
                    status_df,
                    use_container_width=True,
                    hide_index=True,
                    column_config={
                        "教材名称": st.column_config.TextColumn("教材名称", width="medium", alignment="center"),
                        "ISBN": st.column_config.TextColumn("ISBN", width="medium", alignment="center"),
                        "领取状态": st.column_config.TextColumn("领取状态", width="small", alignment="center"),
                        "领书日期": st.column_config.TextColumn("领书日期", width="small", alignment="center"),
                    },
                )

                total_books = len(books)
                c1, c2, c3 = st.columns(3)
                c1.metric("📖 该学期教材数", total_books)
                c2.metric("✅ 已领取", picked)
                c3.metric("❌ 未领取", total_books - picked)

    # ========== Tab 3: 签名确认 ==========
    with tab3:
        st.markdown("### ✍️ 手写签名确认领书")
        st.caption("请在下方签名板上手写签名，确认已领取本学期教材。每学期只需签名一次。")

        # 选择要签名的学期
        sig_sem_opts = [(r["id"], r["name"]) for _, r in sem_df.iterrows()]
        sig_sem = st.selectbox(
            "选择学期",
            sig_sem_opts,
            format_func=lambda x: x[1],
            key="sig_sem_select",
        )

        if sig_sem:
            sem_id, sem_name = sig_sem

            # 检查是否已签名
            existing_sig = query_df(
                "SELECT id, signature_data, signed_at FROM signatures WHERE student_id = %s AND semester_id = %s",
                (student["id"], sem_id)
            )

            if not existing_sig.empty:
                # 显示已有签名
                sig_data = existing_sig.iloc[0]["signature_data"]
                signed_at = existing_sig.iloc[0]["signed_at"]
                st.success(f"✅ 您已於 {signed_at} 完成签名确认")

                # 显示签名图片
                if sig_data:
                    st.markdown("##### 已保存的签名")
                    st.markdown(
                        f'<img src="data:image/png;base64,{sig_data}" style="border:1px solid #ddd; border-radius:4px; max-width:400px; background:#fff;" />',
                        unsafe_allow_html=True,
                    )

                if st.button("🔄 重新签名", key="resign_btn"):
                    execute_sql(
                        "DELETE FROM signatures WHERE student_id = %s AND semester_id = %s",
                        (student["id"], sem_id)
                    )
                    st.rerun()

            else:
                # 签名板 - streamlit-drawable-canvas
                from streamlit_drawable_canvas import st_canvas
                
                st.markdown("##### ✍️ 请在下方面板上签名")
                st.caption("用鼠标或手指在白色区域签名，点击「确认签名」保存")

                canvas_result = st_canvas(
                    stroke_width=2,
                    stroke_color="#000",
                    background_color="#fff",
                    height=300,
                    width=600,
                    drawing_mode="freedraw",
                    display_toolbar=False,
                    key=f"sig_cvs_{sem_id}",
                )

                if st.button("✅ 确认签名", type="primary", key=f"sv_{sem_id}", use_container_width=True):
                    if canvas_result is None or canvas_result.image_data is None:
                        st.warning("画板加载中，请稍后")
                    else:
                        import io
                        from PIL import Image
                        img = Image.fromarray(canvas_result.image_data.astype("uint8"), mode="RGBA").convert("RGB")
                        buf = io.BytesIO()
                        img.save(buf, format="PNG")
                        b64 = base64.b64encode(buf.getvalue()).decode("utf-8")
                        rgb = canvas_result.image_data[:,:,:3]
                        if (rgb==255).all():
                            st.warning("⚠️ 请先在画板上签名")
                        else:
                            execute_sql(
                                "INSERT OR REPLACE INTO signatures (student_id, semester_id, signature_data, signed_at) VALUES (%s, %s, %s, %s)",
                                (student["id"], sem_id, b64, datetime.now()))
                            st.success("✅ 签名保存成功！")
                            st.rerun()
