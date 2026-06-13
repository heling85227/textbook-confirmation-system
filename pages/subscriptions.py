"""
征订总表管理模块 (Subscriptions Management Module)

本模块实现教材征订总表的管理功能，包括：
- 征订总览查看、编辑和删除
- 手动新增征订记录
- 从 Excel 导入征订数据
- 一键下发到班级征订明细

此模块从原始 app.py 中提取并重构，保持所有原始逻辑不变。
"""

import streamlit as st
import pandas as pd
import math
import sqlite3
import re

from database import query_df, execute_sql, get_sqlite_path
from components import show_header, excel_export, styled_dataframe
from utils import (
    get_filtered_list, get_filtered_grades, get_filtered_majors,
    get_filtered_colleges, safe_int, safe_float, safe_str,
    normalize_grade, make_template_df, read_excel_upload,
    write_import_log, get_class_student_counts, parse_major_grade_from_scope
)


def subscription_management():
    """征订总表：导入教务处原始征订数据（专业年级粒度），一键下发拆分到班级"""
    show_header("📋 征订总表", "管理原始征订数据（专业/年级粒度），自动按实际班级人数分摊后下发")

    semesters = query_df("SELECT id, name FROM semesters ORDER BY id DESC")
    if semesters.empty:
        st.warning("⚠️ 请先在「学期管理」中添加学期")
        return

    semester_options = ["全部"] + [f"{r['id']}|{r['name']}" for _, r in semesters.iterrows()]
    master_books = query_df("SELECT id, name, isbn, publisher, editor, price, course_name FROM textbooks_master ORDER BY name")
    master_options = [(0, "➕ 新增教材...")] + [(r["id"], r["name"]) for _, r in master_books.iterrows()]

    tab1, tab2, tab3, tab4 = st.tabs(["📋 征订总览", "➕ 手动新增", "📥 导入 Excel", "🚀 一键下发"])

    # ── Tab1：征订总览 ──
    with tab1:
        st.markdown("#### 📋 征订总览")
        st.caption("这里列出教务处原始征订数据（专业/年级粒度），支持在线编辑和批量删除。勾选后可在 Tab4 一键下发。")
        with st.expander("🔍 筛选条件", expanded=True):
            col1, col2, col3 = st.columns(3)
            with col1:
                f_sub_sem = st.selectbox("学期", semester_options, key="sub_sem_filter")
            with col2:
                sub_college_opts = ["全部"] + get_filtered_colleges()
                f_sub_college = st.selectbox("学院", sub_college_opts, key="sub_college_filter")
            with col3:
                status_opts = ["全部", "待下发", "已下发"]
                f_sub_status = st.selectbox("状态", status_opts, key="sub_status_filter")

        sub_sql = """SELECT s.id, sm.name as semester_name, s.book_name, s.isbn,
                            s.college, s.major, s.grade, s.class_scope,
                            s.total_qty, s.teacher_qty, s.price,
                            s.status, s.remark, s.source, s.created_at, s.dispatched_at
                     FROM textbook_subscriptions s
                     JOIN semesters sm ON s.semester_id = sm.id
                     WHERE 1=1"""
        sub_params = []
        if f_sub_sem != "全部":
            sub_sql += " AND s.semester_id = %s"; sub_params.append(int(f_sub_sem.split("|")[0]))
        if f_sub_college != "全部":
            sub_sql += " AND s.college = %s"; sub_params.append(f_sub_college)
        if f_sub_status != "全部":
            status_map = {"待下发": "pending", "已下发": "dispatched"}
            sub_sql += " AND s.status = %s"; sub_params.append(status_map[f_sub_status])
        sub_sql += " ORDER BY s.semester_id DESC, s.college, s.major, s.grade, s.book_name"

        sub_df = query_df(sub_sql, tuple(sub_params) if sub_params else None)
        total = len(sub_df)

        if not sub_df.empty:
            # 统计
            pending_cnt = (sub_df["status"] == "pending").sum()
            dispatched_cnt = (sub_df["status"] == "dispatched").sum()
            mc1, mc2, mc3 = st.columns(3)
            mc1.metric("📄 合计记录", total)
            mc2.metric("⏳ 待下发", pending_cnt)
            mc3.metric("✅ 已下发", dispatched_cnt)

            # ── 分页（从 session_state 读取，控件移到表格下方）──
            sub_ps = st.session_state.get("sub_ps", 50)
            sub_page = st.session_state.get("sub_page", 1)
            sub_pages = max(1, math.ceil(total / sub_ps))
            # 页码保护
            if sub_page > sub_pages:
                sub_page = sub_pages
                st.session_state.sub_page = sub_pages

            st.caption(f"共 **{total}** 条 ｜ 第 {sub_page}/{sub_pages} 页")

            # 当前页
            s_start = (sub_page - 1) * sub_ps
            s_end = min(s_start + sub_ps, total)
            page_df = sub_df.iloc[s_start:s_end].copy()

            # 原始备份用于对比修改（先用 NaN 安全的字典）
            edit_cols = ["college", "major", "grade", "class_scope", "book_name", "isbn", "price", "total_qty", "teacher_qty", "remark"]
            orig_key = "_sub_orig"
            page_df[orig_key] = page_df[edit_cols].apply(
                lambda row: {col: ("" if pd.isna(row[col]) and col != "price" and col != "total_qty" and col != "teacher_qty" else row[col]) for col in edit_cols},
                axis=1
            ).tolist()

            # ── 构建展示用 DataFrame，确保类型正确 ──
            # 全选初始化
            sel_all_init = st.session_state.get(f"sub_selall_{sub_page}", False)

            # 状态转中文
            status_lbl = {"pending": "⏳ 待下发", "dispatched": "✅ 已下发"}
            status_display = page_df["status"].map(status_lbl).fillna(page_df["status"])

            # 班级范围（处理 None/空）
            class_scope_display = page_df["class_scope"].fillna("全部班级").replace("", "全部班级")

            # 下发时间
            dispatched_display = page_df["dispatched_at"].fillna("—")

            # 数值列——强制转 float 并填 0，避免 NaN 导致 NumberColumn 异常
            price_vals = pd.to_numeric(page_df["price"], errors="coerce").fillna(0.0)
            total_qty_vals = pd.to_numeric(page_df["total_qty"], errors="coerce").fillna(0)
            teacher_qty_vals = pd.to_numeric(page_df["teacher_qty"], errors="coerce").fillna(0)

            # 文本列——填空字符串
            text_fill = lambda s: s.fillna("").astype(str)

            display_df = pd.DataFrame({
                "选择": sel_all_init,
                "学期": page_df["semester_name"].fillna("").values,
                "学院": text_fill(page_df["college"]),
                "专业": text_fill(page_df["major"]),
                "年级": text_fill(page_df["grade"]),
                "班级范围": text_fill(class_scope_display),
                "教材名称": text_fill(page_df["book_name"]),
                "书号": text_fill(page_df["isbn"]),
                "单价": price_vals,
                "征订总量": total_qty_vals.astype(int),
                "教师用书": teacher_qty_vals.astype(int),
                "状态": status_display.values,
                "备注": text_fill(page_df["remark"]),
                "下发时间": dispatched_display.values,
            })

            edited = st.data_editor(
                display_df,
                use_container_width=True,
                hide_index=True,
                column_order=["选择", "学期", "学院", "专业", "年级", "班级范围",
                             "教材名称", "书号", "单价", "征订总量", "教师用书", "状态", "备注", "下发时间"],
                column_config={
                    "选择": st.column_config.CheckboxColumn("选择"),
                    "学期": st.column_config.TextColumn("学期", disabled=True, alignment="center"),
                    "学院": st.column_config.TextColumn("学院", alignment="center"),
                    "专业": st.column_config.TextColumn("专业", alignment="center"),
                    "年级": st.column_config.TextColumn("年级", alignment="center"),
                    "班级范围": st.column_config.TextColumn("班级范围", alignment="center"),
                    "教材名称": st.column_config.TextColumn("教材名称", alignment="center"),
                    "书号": st.column_config.TextColumn("书号", alignment="center"),
                    "单价": st.column_config.NumberColumn("单价", format="¥%.2f", alignment="center"),
                    "征订总量": st.column_config.NumberColumn("征订总量", min_value=0, alignment="center"),
                    "教师用书": st.column_config.NumberColumn("教师用书", min_value=0, alignment="center"),
                    "状态": st.column_config.TextColumn("状态", disabled=True, alignment="center"),
                    "备注": st.column_config.TextColumn("备注", alignment="center"),
                    "下发时间": st.column_config.TextColumn("下发时间", disabled=True, alignment="center"),
                },
                key=f"sub_editor_{sub_page}"
            )

            # ═══ 布局：上行(全选+分页) + 下行(红删蓝存) ═══
            # 检测修改（DB列名 → 展示列名映射）
            col_map = {
                "college": "学院", "major": "专业", "grade": "年级",
                "class_scope": "班级范围", "book_name": "教材名称", "isbn": "书号",
                "price": "单价", "total_qty": "征订总量", "teacher_qty": "教师用书", "remark": "备注"
            }
            changes_sub = []
            for i in range(len(edited)):
                orig = page_df.iloc[i][orig_key]
                for db_col, disp_col in col_map.items():
                    new_val = edited.iloc[i].get(disp_col)
                    old_val = orig.get(db_col)
                    if db_col in ("price", "total_qty", "teacher_qty"):
                        try:
                            nv = float(new_val) if not pd.isna(new_val) else 0.0
                        except (ValueError, TypeError):
                            nv = 0.0
                        try:
                            ov = float(old_val) if not pd.isna(old_val) else 0.0
                        except (ValueError, TypeError):
                            ov = 0.0
                        if abs(nv - ov) > 1e-9:
                            changes_sub.append((int(page_df.iloc[i]["id"]), db_col, nv))
                    else:
                        nv_str = str(new_val or "").strip()
                        ov_str = str(old_val or "").strip()
                        if nv_str != ov_str:
                            changes_sub.append((int(page_df.iloc[i]["id"]), db_col, nv_str if db_col != "class_scope" else (new_val if new_val and str(new_val) != "全部班级" else "")))

            # 处理按钮触发的操作（用 session_state 代替 query_params）
            pending = st.session_state.get("pending_action", "")
            if pending == "sub_save":
                if changes_sub:
                    for sid, col, val in changes_sub:
                        execute_sql(f"UPDATE textbook_subscriptions SET {col}=%s WHERE id=%s", (val, sid))
                    st.toast(f"✅ 已保存 {len(changes_sub)} 处修改", icon="✅")
                st.session_state.pending_action = ""
                st.rerun()
            elif pending == "sub_del":
                select_all = st.session_state.get(f"sub_selall_{sub_page}", False)
                if select_all:
                    del_ids_sub = tuple(page_df["id"].tolist())
                else:
                    selected = edited[edited["选择"] == True].index if "选择" in edited.columns else []
                    del_ids_sub = tuple(page_df.iloc[selected]["id"].tolist()) if len(selected) > 0 else ()
                if del_ids_sub:
                    if len(del_ids_sub) == 1:
                        execute_sql("DELETE FROM textbook_subscriptions WHERE id = %s", (del_ids_sub[0],))
                    else:
                        ph_sub = ",".join(["%s"] * len(del_ids_sub))
                        execute_sql(f"DELETE FROM textbook_subscriptions WHERE id IN ({ph_sub})", del_ids_sub)
                    st.toast(f"✅ 已删除 {len(del_ids_sub)} 条", icon="✅")
                st.session_state.pending_action = ""
                st.rerun()

            # ── 上行：全选 + 分页 ──
            r1_sel, r1_info, r1_ps, r1_prev, r1_num, r1_next = st.columns([1.5, 2, 1, 0.7, 0.7, 0.7])
            with r1_sel:
                select_all = st.checkbox("全选本页", key=f"sub_selall_{sub_page}",
                    help="勾选后删除操作将应用于本页全部记录")
            with r1_info:
                st.caption(f"共 **{total}** 条")
            with r1_ps:
                st.selectbox("每页", [50, 100, 200], key="sub_ps",
                             on_change=lambda: st.session_state.update({"sub_page": 1}),
                             label_visibility="collapsed")
            with r1_prev:
                if st.button("◀", key="sub_pp", disabled=(sub_page <= 1), use_container_width=True):
                    st.session_state.sub_page = sub_page - 1; st.rerun()
            with r1_num:
                st.markdown(f"<div style='text-align:center;padding-top:5px;font-weight:500'>{sub_page}/{sub_pages}</div>", unsafe_allow_html=True)
            with r1_next:
                if st.button("▶", key="sub_np", disabled=(sub_page >= sub_pages), use_container_width=True):
                    st.session_state.sub_page = sub_page + 1; st.rerun()

            # ── 下行：删除(红) + 保存(蓝) ──
            del_count = len(page_df) if select_all else (int(edited["选择"].sum()) if "选择" in edited.columns else 0)
            del_disabled = "disabled" if del_count == 0 else ""
            save_disabled = "disabled" if not changes_sub else ""
            del_btn_label = f"🗑️ 删除（{del_count}条）"
            save_btn_label = f"💾 保存修改（{len(changes_sub)}处）" if changes_sub else "💾 保存修改"

            r2_del, r2_save = st.columns([1, 1])
            with r2_del:
                st.button(del_btn_label, key=f"sub_del_btn_{sub_page}", type="secondary",
                          disabled=(del_count == 0), use_container_width=True,
                          on_click=lambda: st.session_state.update({"pending_action": "sub_del"}))
            with r2_save:
                st.button(save_btn_label, key=f"sub_save_btn_{sub_page}", type="primary",
                          disabled=(not changes_sub), use_container_width=True,
                          on_click=lambda: st.session_state.update({"pending_action": "sub_save"}))
        else:
            st.info("暂无征订数据，请通过「手动新增」或「导入 Excel」添加")

    # ── Tab2：手动新增 ──
    with tab2:
        st.markdown("#### ➕ 手动新增征订记录")
        st.caption("按「专业+年级」粒度登记。征订数量按各班学生人数自动计算（每人1本），下发时每班分配其学生人数。")

        sem_id2 = st.selectbox("学期*", [(r["id"], r["name"]) for _, r in semesters.iterrows()],
            format_func=lambda x: x[1], key="sub_sem2")

        # 教材选择（支持多选）
        actual_books = [(r["id"], r["name"]) for _, r in master_books.iterrows()]
        if not actual_books:
            st.warning("⚠️ 教材库为空，请先在「教材表管理」中添加教材")
            return
        picked_books = st.multiselect(
            "选择已有教材（可多选）",
            options=actual_books,
            format_func=lambda x: x[1],
            key="sub_book2",
            placeholder="点击选择一本或多本教材..."
        )

        if picked_books:
            st.caption(f"📚 已选 **{len(picked_books)}** 本教材：")
            # 展示所选教材摘要
            summary_rows = []
            for pid, pname in picked_books:
                r = master_books[master_books["id"] == pid]
                if not r.empty:
                    summary_rows.append({
                        "教材名称": pname,
                        "书号": str(r.iloc[0].get("isbn", "") or ""),
                        "单价(元)": float(r.iloc[0].get("price", 0)),
                    })
            styled_dataframe(pd.DataFrame(summary_rows), hide_ids=True)

        with st.expander("🎯 征订范围（级联选择）", expanded=True):
            st.divider()
            st.caption("🎯 征订范围（按顺序选择，级联过滤）")

            # ── ① 年级 ──
            all_grades2 = get_filtered_grades()
            sub_grade = st.selectbox("① 年级", ["（不限）"] + all_grades2, key="sub_grade2",
                format_func=lambda x: f"📘 {x}" if x != "（不限）" else x)

            # ── ② 学院（级联年级）──
            if sub_grade != "（不限）":
                all_colleges2 = get_filtered_list("students", "college", "grade=%s OR grade=%s",
                    (sub_grade.rstrip("级"), sub_grade.rstrip("级") + "级"))
            else:
                all_colleges2 = get_filtered_colleges()
            sub_college = st.selectbox("② 学院", ["（不限）"] + all_colleges2, key="sub_college2",
                format_func=lambda x: f"🏫 {x}" if x != "（不限）" else x)

            # ── ③ 专业（级联年级+学院）──
            major_where_parts = []
            major_params_parts = []
            if sub_grade != "（不限）":
                major_where_parts.append("(grade=%s OR grade=%s)")
                major_params_parts.extend([sub_grade.rstrip("级"), sub_grade.rstrip("级") + "级"])
            if sub_college != "（不限）":
                major_where_parts.append("college=%s")
                major_params_parts.append(sub_college)
            if major_where_parts:
                all_majors2 = get_filtered_list("students", "major",
                    " AND ".join(major_where_parts), tuple(major_params_parts))
            else:
                all_majors2 = get_filtered_majors()
            sub_major = st.selectbox("③ 专业", ["（不限）"] + all_majors2, key="sub_major2",
                format_func=lambda x: f"📚 {x}" if x != "（不限）" else x)

            # ── ④ 班级范围（自动计算，无需手动填写）──
            _preview_grade   = normalize_grade(sub_grade) if sub_grade != "（不限）" else None
            _preview_college = None if sub_college == "（不限）" else sub_college
            _preview_major   = None if sub_major  == "（不限）" else sub_major

            auto_scope = ""
            sub_class_names = []
            if _preview_grade or _preview_college or _preview_major:
                preview_df = get_class_student_counts(_preview_grade, _preview_college, _preview_major)
                if not preview_df.empty:
                    matched_classes = preview_df["class_name"].tolist()
                    total_stu = int(preview_df["student_count"].sum())
                    # 自动生成班级范围说明
                    parts = []
                    if _preview_grade: parts.append(_preview_grade)
                    if _preview_college: parts.append(_preview_college)
                    if _preview_major: parts.append(_preview_major)
                    auto_scope = " ".join(parts) + " 全部班级"
                    st.success(f"📋 ④ 班级范围（自动匹配）：**{auto_scope}** — {len(matched_classes)}个班级，{total_stu}名学生")
                    # 班级详情
                    with st.expander(f"查看匹配的 {len(matched_classes)} 个班级"):
                        class_preview_rows = []
                        for _, pr in preview_df.iterrows():
                            class_preview_rows.append({
                                "班级": pr["class_name"],
                                "学生人数": int(pr["student_count"])
                            })
                        styled_dataframe(pd.DataFrame(class_preview_rows), hide_ids=True)

                    # 可选：进一步缩小班级范围（多选）
                    st.caption("如需精确到特定班级，可从下方多选（留空=匹配全部）")
                    sub_class_names = st.multiselect(
                        "精确班级（可多选）",
                        options=matched_classes,
                        key="sub_class_names2",
                        placeholder="留空 = 匹配上述全部班级"
                    )
                    if sub_class_names:
                        auto_scope = "、".join(sub_class_names)
                        st.info(f"✂️ 仅下发到：**{auto_scope}**")
                else:
                    st.warning("⚠️ ④ 当前条件下未找到任何班级，请检查学生数据中是否有对应年级的数据")
                    preview_df = pd.DataFrame()  # 确保变量存在
            else:
                st.info("💡 未选择任何条件 = 匹配学生表中所有班级（请谨慎使用）")
                auto_scope = "全部班级"
                preview_df = pd.DataFrame()  # 确保变量存在

        # ── 按班级人数自动计算数量（每人1本）──
        with st.expander("📦 征订数量", expanded=True):
            st.caption("📦 征订数量（按班级人数自动计算，每人1本）")
            if not preview_df.empty:
                sub_total_qty = int(preview_df["student_count"].sum())
                q_c1, q_c2 = st.columns(2)
                with q_c1:
                    st.metric("征订总量（学生用书）", f"{sub_total_qty} 本",
                             delta=f"{len(preview_df)}个班级" if not sub_class_names else f"{len(sub_class_names)}个精确班级")
                with q_c2:
                    sub_teacher_qty = st.number_input("教师用书数量", min_value=0, value=0, step=1, key="sub_teacher_qty2")
                # 每班按人数分配预览
                if sub_class_names:
                    filtered_class = preview_df[preview_df["class_name"].isin(sub_class_names)]
                else:
                    filtered_class = preview_df
                prows = [{"班级": pr["class_name"], "学生人数": int(pr["student_count"]), "分配数量": int(pr["student_count"])}
                         for _, pr in filtered_class.iterrows()]
                if prows:
                    st.caption("📦 每班分配预览（每人1本）：")
                    styled_dataframe(pd.DataFrame(prows), hide_ids=True)
            else:
                sub_total_qty = 0
                sub_teacher_qty = 0
                st.warning("⚠️ 请先选择年级/学院/专业，系统将自动按班级人数计算征订数量")

        sub_remark = st.text_input("备注", key="sub_remark2")

        if st.button("💾 保存到征订总表", use_container_width=True, type="primary", key="sub_save2"):
            if not picked_books:
                st.error("❌ 请至少选择一本教材")
                st.stop()

            saved_count = 0
            for pid, pname in picked_books:
                r = master_books[master_books["id"] == pid]
                if r.empty:
                    continue
                b = r.iloc[0]
                execute_sql("""INSERT INTO textbook_subscriptions
                    (semester_id, textbook_id, book_name, isbn, publisher, editor, price, course_name,
                     college, major, grade, class_scope, class_names, total_qty, teacher_qty, status, remark, source)
                    VALUES (%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,'pending',%s,'manual')""",
                    (sem_id2[0], pid, b["name"],
                     str(b.get("isbn", "") or ""), str(b.get("publisher", "") or ""),
                     str(b.get("editor", "") or ""), float(b.get("price", 0)),
                     str(b.get("course_name", "") or ""),
                     _preview_college or "", _preview_major or "", _preview_grade or "",
                     auto_scope, ",".join(sub_class_names) if sub_class_names else "",
                     sub_total_qty, sub_teacher_qty, sub_remark))
                saved_count += 1

            st.success(f"✅ 已保存 **{saved_count}** 条征订记录到总表")
            st.rerun()

    # ── Tab3：导入 Excel ──
    with tab3:
        st.markdown("#### 📥 从 Excel 导入原始征订数据")
        st.caption("支持教务处格式：专业年级（如「电商2023」）、征订数量、教师用书数量等列。")

        # 模板下载
        with st.expander("📄 模板下载与说明", expanded=False):
            tmpl_cols = ["学期", "学院", "专业", "年级", "班级范围说明", "教材名称", "书号(ISBN)", "出版社", "主编", "单价(元)", "征订总量", "教师用书数量", "备注"]
            tmpl_df = make_template_df(tmpl_cols)
            st.download_button("📄 下载征订总表导入模板", data=excel_export(tmpl_df, "征订总表导入模板"),
                file_name="征订总表导入模板.xlsx", mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                use_container_width=True, type="secondary")
            st.caption("Excel 列名建议包含：学期、学院、专业、年级、班级范围说明、教材名称、书号(ISBN)、出版社、主编、单价(元)、征订总量、教师用书数量")

        uploaded3 = st.file_uploader("选择 Excel 文件", type=["xlsx", "xls"], key="sub_upload3")
        if uploaded3:
            try:
                raw3 = read_excel_upload(uploaded3)
                st.info(f"📄 检测到 {len(raw3)} 行数据")

                # 列映射
                cmap3 = {}
                for col in raw3.columns:
                    cl = col.strip().lower()
                    if "教材" in cl and "名" in cl: cmap3["book_name"] = col
                    elif "学期" in cl: cmap3["semester_name"] = col
                    elif "专业" in cl and "年级" in cl: cmap3["major_grade"] = col  # 合并列
                    elif "年级" in cl: cmap3["grade"] = col
                    elif "学院" in cl: cmap3["college"] = col
                    elif "专业" in cl: cmap3["major"] = col
                    elif "班级" in cl and "范围" in cl: cmap3["class_scope"] = col
                    elif "班级" in cl: cmap3["class_scope"] = col
                    elif "教师" in cl and ("用书" in cl or "数量" in cl): cmap3["teacher_qty"] = col
                    elif "单价" in cl: cmap3["price"] = col
                    elif ("数量" in cl or "合计" in cl or "总量" in cl) and "教师" not in cl: cmap3["total_qty"] = col
                    elif "书号" in cl or "isbn" in cl: cmap3["isbn"] = col
                    elif "出版社" in cl: cmap3["publisher"] = col
                    elif "主编" in cl or "作者" in cl: cmap3["editor"] = col
                    elif "备注" in cl or "说明" in cl: cmap3["remark"] = col

                if "book_name" not in cmap3 or "semester_name" not in cmap3:
                    st.error("❌ 缺少必要列：教材名称、学期")
                else:
                    # 预览
                    st.markdown("**列映射预览：**")
                    mapped_preview = {v: k for k, v in cmap3.items()}
                    preview_rows = [{"Excel列名": k, "识别为": v} for k, v in mapped_preview.items()]
                    styled_dataframe(pd.DataFrame(preview_rows), hide_ids=True)

                    has_major_grade = "major_grade" in cmap3
                    if has_major_grade:
                        st.info("📌 检测到「专业年级」合并列，将自动拆解专业名和年级")

                    target_sem = st.selectbox("确认导入到学期", [(r["id"], r["name"]) for _, r in semesters.iterrows()],
                        format_func=lambda x: x[1], key="sub_import_sem3")

                    if st.button("✅ 确认导入到征订总表", use_container_width=True, type="primary", key="sub_import3_btn"):
                        sem_map3 = {r["name"]: r["id"] for _, r in semesters.iterrows()}
                        success3, errors3, total3 = 0, [], len(raw3)
                        progress3 = st.progress(0, text="正在导入...")

                        # ── 预处理：前向填充学期和学院（兼容合并单元格模板）──
                        last_sem, last_college = "", ""
                        for i in range(len(raw3)):
                            sv = safe_str(raw3.iloc[i].get(cmap3.get("semester_name", ""), ""))
                            cv = safe_str(raw3.iloc[i].get(cmap3.get("college", ""), ""))
                            if sv:
                                last_sem = sv
                            elif last_sem:
                                raw3.at[raw3.index[i], cmap3["semester_name"]] = last_sem
                            if cv:
                                last_college = cv
                            elif last_college and "college" in cmap3:
                                raw3.at[raw3.index[i], cmap3["college"]] = last_college

                        # 统计智能解析命中数
                        auto_parse_count = 0

                        for i, (_, row3) in enumerate(raw3.iterrows()):
                            try:
                                book_name3 = safe_str(row3.get(cmap3.get("book_name", ""), ""))
                                if not book_name3: continue
                                # 学期
                                sem_name3 = safe_str(row3.get(cmap3.get("semester_name", ""), ""))
                                sid3 = sem_map3.get(sem_name3, target_sem[0])
                                # 专业年级拆解
                                mg_val3 = safe_str(row3.get(cmap3.get("major_grade", ""), "")) if has_major_grade else ""
                                grade3  = safe_str(row3.get(cmap3.get("grade", ""), ""))
                                major3  = safe_str(row3.get(cmap3.get("major", ""), ""))
                                if has_major_grade and mg_val3 and not (grade3 and major3):
                                    mg_parts3 = re.split(r'(\d{4})', mg_val3, maxsplit=1)
                                    if len(mg_parts3) >= 2:
                                        major3 = major3 or mg_parts3[0].strip()
                                        grade3 = grade3 or mg_parts3[1].strip()
                                    else:
                                        major3 = major3 or mg_val3
                                college3     = safe_str(row3.get(cmap3.get("college", ""), ""))

                                # ── 智能解析：从班级范围说明中提取专业/年级 ──
                                class_scope3 = safe_str(row3.get(cmap3.get("class_scope", ""), ""))
                                parsed_grades = []  # 从 scope 解析出的年级列表
                                if (not major3 or not grade3) and class_scope3:
                                    parsed_major, parsed_grades = parse_major_grade_from_scope(class_scope3, college3)
                                    if parsed_major and not major3:
                                        major3 = parsed_major
                                        auto_parse_count += 1
                                    if parsed_grades and not grade3:
                                        grade3 = parsed_grades[0]  # 先用第一个年级做后续判断
                                # 确定最终要写入的年级列表（优先用模板显式年级，否则用解析结果）
                                final_grades = [grade3] if grade3 else []
                                if parsed_grades and not grade3:
                                    final_grades = parsed_grades
                                # 如果没有年级，给个空列表（兜底）
                                if not final_grades:
                                    final_grades = [grade3]
                                total_qty3   = safe_int(row3.get(cmap3.get("total_qty", ""), 0))
                                teacher_qty3 = safe_int(row3.get(cmap3.get("teacher_qty", ""), 0))
                                price3       = safe_float(row3.get(cmap3.get("price", ""), 0))
                                isbn3        = safe_str(row3.get(cmap3.get("isbn", ""), ""))
                                pub3         = safe_str(row3.get(cmap3.get("publisher", ""), ""))
                                ed3          = safe_str(row3.get(cmap3.get("editor", ""), ""))
                                remark3      = safe_str(row3.get(cmap3.get("remark", ""), ""))
                                # 教材主表
                                old3 = query_df("SELECT id FROM textbooks_master WHERE name=%s", (book_name3,))
                                if old3.empty:
                                    execute_sql("INSERT INTO textbooks_master (name,isbn,publisher,editor,price) VALUES (%s,%s,%s,%s,%s)",
                                        (book_name3, isbn3, pub3, ed3, price3))
                                    _c3 = sqlite3.connect(get_sqlite_path())
                                    _cc3 = _c3.cursor(); _cc3.execute("SELECT MAX(id) FROM textbooks_master"); tid3 = _cc3.fetchone()[0]; _c3.close()
                                else:
                                    tid3 = old3.iloc[0]["id"]
                                # 写入征订总表（多年级则拆分为多条记录，每条保留原始总量）
                                for g3 in final_grades:
                                    execute_sql("""INSERT INTO textbook_subscriptions
                                        (semester_id,textbook_id,book_name,isbn,publisher,editor,price,
                                         college,major,grade,class_scope,total_qty,teacher_qty,status,remark,source)
                                        VALUES (%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,'pending',%s,'import')""",
                                        (sid3, tid3, book_name3, isbn3, pub3, ed3, price3,
                                         college3, major3, g3, class_scope3, total_qty3, teacher_qty3, remark3))
                                    success3 += 1
                            except Exception as e3:
                                errors3.append(f"{safe_str(row3.get(cmap3.get('book_name',''),''))}: {str(e3)[:100]}")
                            progress3.progress((i + 1) / total3, text=f"已处理 {i+1}/{total3}")
                        # 统一展示结果（对齐学生管理）
                        mc_ok, mc_err = st.columns(2)
                        mc_ok.metric("✅ 导入成功", success3)
                        mc_err.metric("⚠️ 导入失败", len(errors3))
                        if auto_parse_count > 0:
                            st.info(f"🔍 智能解析：从「班级范围说明」自动提取了 **{auto_parse_count}** 条的专业/年级信息")
                        if errors3:
                            with st.expander(f"查看 {len(errors3)} 条失败详情"):
                                for e3 in errors3:
                                    st.caption(f"• {e3}")
                        write_import_log("征订总表导入", uploaded3.name, total3, success3, errors3)
                        st.rerun()
            except Exception as e:
                st.error(f"❌ 读取失败：{e}")

    # ── Tab4：一键下发 ──
    with tab4:
        st.markdown("#### 🚀 一键下发：将征订总表数据拆分到班级征订明细")
        st.caption("""
        **下发逻辑：**  
        - 系统根据「学院 + 专业 + 年级」条件，在 students 表中匹配实际班级  
        - 按各班级实际学生人数下发（每人1本），班级学生人数即为该班征订数量  
        - 写入「教材征订表（按班级明细）」，供后续发放使用  
        - 下发后状态变为「已下发」，可重复下发（会覆盖同条件的已有征订）
        """)

        # 选择要下发的记录
        pending_df = query_df("""
            SELECT ts.id, sm.name as semester_name, ts.book_name, ts.college, ts.major, ts.grade,
                   ts.class_scope, ts.total_qty, ts.teacher_qty, ts.status
            FROM textbook_subscriptions ts
            JOIN semesters sm ON ts.semester_id = sm.id
            WHERE ts.status = 'pending'
            ORDER BY ts.semester_id DESC, ts.college, ts.major, ts.grade
        """)

        if pending_df.empty:
            st.info("✅ 暂无待下发的征订记录")
        else:
            st.markdown(f"共有 **{len(pending_df)}** 条待下发记录：")
            status_show = pending_df.copy()
            status_show["班级范围"] = status_show["class_scope"].fillna("全部班级").replace("", "全部班级")
            styled_dataframe(status_show[["id","semester_name","college","major","grade","班级范围","book_name","total_qty","teacher_qty"]].rename(columns={
                "semester_name":"学期","college":"学院","major":"专业","grade":"年级",
                "book_name":"教材名称","total_qty":"征订总量","teacher_qty":"教师用书"
            }), hide_ids=True)

            # 下发选项 — checkbox 多选 + 全选/取消全选
            st.divider()
            st.caption("🎯 选择要下发的记录（勾选后点击底部按钮下发）")

            # 全选 / 取消全选 开关
            select_all_key = "dispatch_select_all"
            if select_all_key not in st.session_state:
                st.session_state[select_all_key] = False

            cb_col1, cb_col2 = st.columns([1, 5])
            with cb_col1:
                if st.button("☑️ 全选" if not st.session_state[select_all_key] else "☐ 取消全选",
                             use_container_width=True, key="dispatch_toggle_all"):
                    st.session_state[select_all_key] = not st.session_state[select_all_key]
                    st.rerun()
            with cb_col2:
                select_count = len(pending_df) if st.session_state[select_all_key] else 0
                st.caption(f"已选 **{select_count}** / {len(pending_df)} 条" if st.session_state[select_all_key] else "点击「全选」或逐条勾选")

            # 逐条 checkbox
            selected_ids = []
            for _, prow in pending_df.iterrows():
                pid = int(prow["id"])
                cb_key = f"dispatch_cb_{pid}"
                # 根据全选状态决定默认值
                default_val = st.session_state[select_all_key]
                if st.checkbox(
                    f"#{pid} 【{prow['semester_name']}】{prow['college']} {prow['major']} {prow['grade']}级 — {prow['book_name']}（总量{prow['total_qty']}，教师{prow['teacher_qty']}）",
                    value=st.session_state.get(cb_key, default_val),
                    key=cb_key
                ):
                    selected_ids.append(pid)

            if selected_ids:
                st.caption(f"✅ 已选中 **{len(selected_ids)}** 条记录")
            else:
                st.warning("⚠️ 请至少勾选一条记录")

            conflict_mode = st.radio("遇到同条件已有征订时", ["跳过（保留原有）", "覆盖（删除原有后重新写入）"], horizontal=True, key="conflict_mode")

            dispatch_disabled = len(selected_ids) == 0
            if st.button("🚀 确认下发所选记录", use_container_width=True, type="primary",
                         key="dispatch_btn", disabled=dispatch_disabled):
                to_dispatch = pending_df[pending_df["id"].isin(selected_ids)]

                # 获取完整信息
                dispatch_detail = query_df("""
                    SELECT ts.*, sm.id as sem_id_val
                    FROM textbook_subscriptions ts
                    JOIN semesters sm ON ts.semester_id = sm.id
                """ + " WHERE ts.id IN (" + ",".join(["%s"]*len(to_dispatch)) + ")",
                    tuple(to_dispatch["id"].tolist()))

                total_created = 0
                total_skipped = 0
                total_errors  = []
                total_records = len(dispatch_detail)
                progress = st.progress(0, text="正在下发...")
                st.info(f"🔄 正在处理 **{total_records}** 条征订记录...")

                for idx, (_, sub_row) in enumerate(dispatch_detail.iterrows()):
                    try:
                        sub_id    = sub_row["id"]
                        sem_id_v  = sub_row["semester_id"]
                        textbook_id = sub_row["textbook_id"]
                        book_name = sub_row["book_name"]

                        # 如果 textbook_id 为空，尝试通过书名查找教材主表
                        if not textbook_id:
                            bk_lookup = query_df("SELECT id FROM textbooks_master WHERE name=%s", (book_name,))
                            if not bk_lookup.empty:
                                textbook_id = int(bk_lookup.iloc[0]["id"])
                                # 回填 textbook_id
                                execute_sql("UPDATE textbook_subscriptions SET textbook_id=%s WHERE id=%s", (textbook_id, sub_row["id"]))
                            else:
                                total_errors.append(f"「{book_name}」：教材库中不存在，请先在教材征订表中添加该教材")
                                progress.progress((idx + 1) / total_records, text=f"处理中 {idx+1}/{total_records}")
                                continue
                        textbook_id = int(textbook_id)
                        college_v = sub_row["college"] or None
                        major_v   = sub_row["major"] or None
                        grade_v   = normalize_grade(sub_row["grade"]) if sub_row.get("grade") else None
                        total_qty_v   = int(sub_row["total_qty"] or 0)
                        teacher_qty_v = int(sub_row["teacher_qty"] or 0)
                        remark_v  = sub_row.get("remark", "") or ""

                        # 读取 class_names（Tab2 中手动精确选择的班级）
                        class_names_str = sub_row.get("class_names", "") or ""
                        class_names_list = [c.strip() for c in class_names_str.split(",") if c.strip()] if class_names_str else None

                        # 按班级人数下发（每人1本），每班分配 = 该班学生人数
                        class_df = get_class_student_counts(grade_v, college_v, major_v, class_names=class_names_list)
                        splits_v = [(row["class_name"], int(row["student_count"]), int(row["student_count"]))
                                     for _, row in class_df.iterrows()] if not class_df.empty else [(None, 0, 0)]

                        if len(splits_v) == 1 and splits_v[0][0] is None:
                            # 没找到匹配的班级
                            total_errors.append(f"「{book_name}」({college_v} {major_v} {grade_v}级)：未找到匹配班级，跳过")
                            progress.progress((idx + 1) / total_records, text=f"处理中 {idx+1}/{total_records}")
                            continue

                        for cn, sc, alloc in splits_v:
                            if alloc <= 0:
                                continue
                            # 检查是否已有同条件的征订记录
                            existing = query_df("""
                                SELECT id FROM textbook_orders
                                WHERE semester_id=%s AND textbook_id=%s AND class_name=%s
                            """, (sem_id_v, textbook_id, cn))

                            if not existing.empty:
                                if conflict_mode == "跳过（保留原有）":
                                    total_skipped += 1
                                    continue
                                else:
                                    # 覆盖：删除原有
                                    ph_del = ",".join(["%s"]*len(existing))
                                    execute_sql(f"DELETE FROM textbook_orders WHERE id IN ({ph_del})",
                                        tuple(existing["id"].tolist()))

                            execute_sql("""INSERT INTO textbook_orders
                                (semester_id, textbook_id, grade, college, major, class_name, quantity, remark)
                                VALUES (%s,%s,%s,%s,%s,%s,%s,%s)""",
                                (sem_id_v, textbook_id, grade_v or "", college_v or "", major_v or "",
                                 cn, alloc, (remark_v or "")))
                            total_created += 1

                        # 更新状态为已下发
                        execute_sql("UPDATE textbook_subscriptions SET status='dispatched', dispatched_at=CURRENT_TIMESTAMP WHERE id=%s", (sub_id,))

                    except Exception as e_d:
                        total_errors.append(f"ID={sub_row['id']} {sub_row['book_name']}: {str(e_d)[:120]}")
                    progress.progress((idx + 1) / total_records, text=f"处理中 {idx+1}/{total_records}")

                progress.empty()
                st.divider()
                st.markdown("### 📊 下发结果")
                mc1, mc2, mc3 = st.columns(3)
                mc1.metric("✅ 新增班级征订", total_created)
                mc2.metric("⏭️ 跳过（已有）", total_skipped)
                mc3.metric("⚠️ 失败/异常", len(total_errors))
                if total_errors:
                    with st.expander(f"查看 {len(total_errors)} 条失败详情"):
                        for e_d in total_errors:
                            st.caption(f"• {e_d}")
                if total_created > 0:
                    st.success(f"🎉 下发完成！新增 **{total_created}** 条班级征订记录到教材征订表")
                else:
                    st.warning("⚠️ 没有新增任何班级征订记录，请检查征订条件和班级数据")
