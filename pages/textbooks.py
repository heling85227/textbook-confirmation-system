"""
教材征订表管理页面 (textbooks.py)
功能：按班级管理征订明细（教材请前往「教材表管理」维护）
包含：征订列表、新增/编辑征订、导入Excel、教材批量分配四个标签页
"""

import streamlit as st
import pandas as pd
import math
import sqlite3
import re

from database import query_df, execute_sql, get_sqlite_path
from components import show_header, excel_export, styled_dataframe
from utils import (get_filtered_list, get_filtered_grades, get_filtered_majors,
                   get_filtered_colleges, get_filtered_class_names,
                   safe_int, safe_float, safe_str,
                   make_template_df, read_excel_upload,
                   write_import_log, get_class_student_counts)


def textbook_management():
    """教材征订表管理（双表架构：textbooks_master + textbook_orders）"""
    show_header("📖 教材征订表", "按班级管理征订明细（教材请前往「教材表管理」维护）")

    semesters = query_df("SELECT id, name FROM semesters ORDER BY id DESC")
    if semesters.empty:
        st.warning("⚠️ 请先在「学期管理」中添加学期")
        return

    semester_options = ["全部"] + [f"{r['id']}|{r['name']}" for _, r in semesters.iterrows()]
    master_books = query_df("SELECT id, name, isbn, publisher, editor, price, course_name FROM textbooks_master ORDER BY name")
    master_options = [(0, "➕ 新增教材...")] + [(r["id"], r["name"]) for _, r in master_books.iterrows()]

    tab1, tab2, tab3, tab4 = st.tabs(["📋 征订列表", "➕ 新增/编辑征订", "📥 导入 Excel", "📚 教材批量分配"])

    with tab1:
        with st.expander("🔍 筛选条件", expanded=True):
            col1, col2, col3, col4 = st.columns(4)
            with col1: f_sem = st.selectbox("学期", semester_options, key="t_sem")
            with col2: f_tcollege = st.selectbox("学院", ["全部"] + get_filtered_colleges(), key="t_college")
            with col3:
                # 专业：基于所选学院级联过滤
                if f_tcollege != "全部":
                    t_major_opts = ["全部"] + get_filtered_list("students", "major", "college = %s", (f_tcollege,))
                else:
                    t_major_opts = ["全部"] + get_filtered_majors()
                f_tmajor = st.selectbox("专业", t_major_opts, key="t_major")
            with col4:
                # 班级：基于所选学院+专业级联过滤
                t_class_where = "1=1"; t_class_params = []
                if f_tcollege != "全部":
                    t_class_where += " AND college = %s"; t_class_params.append(f_tcollege)
                if f_tmajor != "全部":
                    t_class_where += " AND major = %s"; t_class_params.append(f_tmajor)
                t_class_opts = ["全部"] + (get_filtered_list("students", "class_name", t_class_where, tuple(t_class_params)) if t_class_params else get_filtered_class_names())
                f_tclass = st.selectbox("班级", t_class_opts, key="t_class")
        f_tsearch = st.text_input("🔍 搜索（教材名）", key="t_search", placeholder="输入教材名称...")

        sql = """SELECT o.*, tm.name, tm.isbn, tm.publisher, tm.editor, tm.price, tm.course_name,
                        s.name as semester_name
                 FROM textbook_orders o
                 JOIN textbooks_master tm ON o.textbook_id = tm.id
                 JOIN semesters s ON o.semester_id = s.id WHERE 1=1"""
        params = []
        if f_sem != "全部": sql += " AND o.semester_id = %s"; params.append(int(f_sem.split("|")[0]))
        if f_tcollege != "全部": sql += " AND o.college = %s"; params.append(f_tcollege)
        if f_tmajor != "全部": sql += " AND o.major = %s"; params.append(f_tmajor)
        if f_tclass != "全部": sql += " AND o.class_name = %s"; params.append(f_tclass)
        if f_tsearch: sql += " AND tm.name LIKE %s"; params.append(f"%{f_tsearch}%")
        sql += " ORDER BY o.semester_id DESC, o.college, o.major, o.class_name, tm.name"

        df = query_df(sql, tuple(params) if params else None)
        if not df.empty:
            df["小计"] = pd.to_numeric(df["price"], errors="coerce").fillna(0) * pd.to_numeric(df["quantity"], errors="coerce").fillna(0)
            total_amount = df["小计"].sum()

            # 分页（从 session_state 读取，控件移到表格下方）
            t_ps = st.session_state.get("t_ps", 30)
            t_pg = st.session_state.get("t_pg", 1)
            tp = max(1, math.ceil(len(df) / t_ps))
            # 页码保护：删除后当前页可能越界，自动跳到最后一页
            if t_pg > tp:
                t_pg = tp
                st.session_state.t_pg = tp

            start = (t_pg - 1) * t_ps
            end = min(start + t_ps, len(df))
            page_df = df.iloc[start:end].copy()

            st.caption(f"共 **{len(df)}** 条征订记录 ｜ 第 {t_pg}/{tp} 页 ｜ 💰 总计 ¥{total_amount:,.2f}")

            cols = ["id","semester_name","college","major","class_name","grade",
                    "name","course_name","isbn","publisher","editor","price","quantity","remark"]
            rename_map = {"semester_name":"学期","college":"学院","major":"专业","class_name":"班级","grade":"年级",
                "name":"教材名称","course_name":"课程","isbn":"书号","publisher":"出版社","editor":"主编",
                "price":"单价(元)","quantity":"征订数量","remark":"备注"}

            display_df = page_df[cols].rename(columns=rename_map).copy()
            t_selall = st.session_state.get(f"t_sel_{t_pg}", False)
            display_df["选择"] = t_selall
            ed_k = f"tb_ed_{t_pg}"
            edited = st.data_editor(
                display_df,
                use_container_width=True, hide_index=True,
                column_order=["选择","学期","学院","专业","班级","年级","教材名称","课程","书号","出版社","主编","单价(元)","征订数量","备注"],
                column_config={
                    "选择": st.column_config.CheckboxColumn("选择"),
                    "学期": st.column_config.TextColumn("学期", disabled=True, alignment="center"),
                    "学院": st.column_config.TextColumn("学院", disabled=True, alignment="center"),
                    "专业": st.column_config.TextColumn("专业", disabled=True, alignment="center"),
                    "班级": st.column_config.TextColumn("班级", disabled=True, alignment="center"),
                    "年级": st.column_config.TextColumn("年级", disabled=True, alignment="center"),
                    "教材名称": st.column_config.TextColumn("教材名称", disabled=True, alignment="center"),
                    "课程": st.column_config.TextColumn("课程", disabled=True, alignment="center"),
                    "书号": st.column_config.TextColumn("书号", disabled=True, alignment="center"),
                    "出版社": st.column_config.TextColumn("出版社", disabled=True, alignment="center"),
                    "主编": st.column_config.TextColumn("主编", disabled=True, alignment="center"),
                    "单价(元)": st.column_config.NumberColumn("单价(元)", format="¥%.2f", disabled=True, alignment="center"),
                    "征订数量": st.column_config.NumberColumn("征订数量", min_value=0, step=1, alignment="center"),
                    "备注": st.column_config.TextColumn("备注", alignment="center"),
                },
                disabled=["学期","学院","专业","班级","年级","教材名称","课程","书号","出版社","主编"],
                key=ed_k
            )

            # ═══ 布局：上行(全选+分页) + 下行(红删蓝存) ═══
            # 检测数量+备注修改
            t_changes = []
            for i in range(len(edited)):
                rid = page_df.iloc[i]["id"]
                old_q = int(page_df.iloc[i].get("quantity", 0))
                new_q = int(edited.iloc[i].get("征订数量", 0))
                old_r = str(page_df.iloc[i].get("remark", "") or "")
                new_r = str(edited.iloc[i].get("备注", "") or "")
                if old_q != new_q:
                    t_changes.append((rid, "quantity", new_q))
                if old_r != new_r:
                    t_changes.append((rid, "remark", new_r))

            # 处理按钮触发的操作（用 session_state 代替 query_params）
            pending = st.session_state.get("pending_action", "")
            if pending == "t_save":
                if t_changes:
                    for rid, col, val in t_changes:
                        execute_sql(f"UPDATE textbook_orders SET {col}=%s WHERE id=%s", (val, rid))
                    st.toast(f"✅ 已保存 {len(t_changes)} 处修改", icon="✅")
                st.session_state.pending_action = ""
                st.rerun()
            elif pending == "t_del":
                t_sel_all = st.session_state.get(f"t_sel_{t_pg}", False)
                if t_sel_all:
                    del_ids = tuple(page_df["id"].tolist())
                else:
                    selected = edited[edited["选择"] == True].index if "选择" in edited.columns else []
                    del_ids = tuple(page_df.iloc[selected]["id"].tolist()) if len(selected) > 0 else ()
                if del_ids:
                    if len(del_ids) == 1:
                        execute_sql("DELETE FROM textbook_orders WHERE id=%s", (del_ids[0],))
                    else:
                        ph = ",".join(["%s"] * len(del_ids))
                        execute_sql(f"DELETE FROM textbook_orders WHERE id IN ({ph})", del_ids)
                    st.toast(f"✅ 已删除 {len(del_ids)} 条", icon="✅")
                st.session_state.pending_action = ""
                st.rerun()

            # ── 上行：全选 + 分页 ──
            r1_sel, r1_info, r1_ps, r1_prev, r1_num, r1_next = st.columns([1.5, 2, 1, 0.7, 0.7, 0.7])
            with r1_sel:
                t_sel_all = st.checkbox("全选本页", key=f"t_sel_{t_pg}",
                    help="勾选后删除操作将应用于本页全部记录")
            with r1_info:
                st.caption(f"共 **{len(df)}** 条 ｜ ¥{total_amount:,.2f}")
            with r1_ps:
                st.selectbox("每页", [30, 50, 100], key="t_ps",
                             on_change=lambda: st.session_state.update({"t_pg": 1}),
                             label_visibility="collapsed")
            with r1_prev:
                if st.button("◀", key="t_pp", disabled=(t_pg <= 1), use_container_width=True):
                    st.session_state.t_pg = t_pg - 1; st.rerun()
            with r1_num:
                st.markdown(f"<div style='text-align:center;padding-top:5px;font-weight:500'>{t_pg}/{tp}</div>", unsafe_allow_html=True)
            with r1_next:
                if st.button("▶", key="t_np", disabled=(t_pg >= tp), use_container_width=True):
                    st.session_state.t_pg = t_pg + 1; st.rerun()

            # ── 下行：删除(红) + 保存(蓝) ──
            del_count = len(page_df) if t_sel_all else (int(edited["选择"].sum()) if "选择" in edited.columns else 0)
            del_disabled = "disabled" if del_count == 0 else ""
            save_disabled = "disabled" if not t_changes else ""
            del_btn_label = f"🗑️ 删除（{del_count}条）"
            save_btn_label = f"💾 保存修改（{len(t_changes)}处）" if t_changes else "💾 保存修改"

            r2_del, r2_save = st.columns([1, 1])
            with r2_del:
                st.button(del_btn_label, key=f"t_del_btn_{t_pg}", type="secondary",
                          disabled=(del_count == 0), use_container_width=True,
                          on_click=lambda: st.session_state.update({"pending_action": "t_del"}))
            with r2_save:
                st.button(save_btn_label, key=f"t_save_btn_{t_pg}", type="primary",
                          disabled=(not t_changes), use_container_width=True,
                          on_click=lambda: st.session_state.update({"pending_action": "t_save"}))

        else:
            st.info("暂无征订数据")

    with tab2:
        st.markdown("#### ➕ 新增征订 / 管理教材库")
        edit_oid = st.number_input("编辑征订ID（留空为新增）", min_value=0, value=0, step=1, key="edit_oid")
        defaults = {}
        if edit_oid > 0:
            row = query_df("SELECT o.*, tm.name as book_name FROM textbook_orders o JOIN textbooks_master tm ON o.textbook_id=tm.id WHERE o.id=%s", (edit_oid,))
            if not row.empty: defaults = row.iloc[0].to_dict()

        # 教材选择放在表单外面，触发页面重绘
        st.caption("🔹 第一步：选择教材（从教材库选择，或选「新增教材」手动输入）")
        default_master_idx = 0 if edit_oid == 0 else next((i for i, o in enumerate(master_options) if o[0] == defaults.get("textbook_id", 0)), 0)
        picked = st.selectbox("选择已有教材", master_options,
            format_func=lambda x: x[1], index=default_master_idx, key="sel_master_out")
        is_new_book = (picked[0] == 0)

        # ── 联动：选中教材时自动填充信息（仅在选择变化时更新）──
        prev_book_id2 = st.session_state.get("_bk_prev_book_id", -1)
        if picked[0] != prev_book_id2:
            st.session_state["_bk_prev_book_id"] = picked[0]
            if not is_new_book:
                bk_info = master_books[master_books["id"] == picked[0]]
                if not bk_info.empty:
                    r = bk_info.iloc[0]
                    st.session_state["bk_name"] = r["name"]
                    st.session_state["bk_isbn"] = str(r.get("isbn", "") or "")
                    st.session_state["bk_pub"] = str(r.get("publisher", "") or "")
                    st.session_state["bk_ed"] = str(r.get("editor", "") or "")
                    st.session_state["bk_price"] = float(r.get("price", 0))
                    st.session_state["bk_course"] = str(r.get("course_name", "") or "")
            else:
                for k in ["bk_name", "bk_isbn", "bk_pub", "bk_ed", "bk_course"]:
                    st.session_state[k] = ""
                st.session_state["bk_price"] = 0.0

        st.caption("第二步：教材信息（从教材库自动带出，也可修改）")
        with st.expander("教材信息", expanded=is_new_book):
            cols_book = st.columns(3)
            with cols_book[0]:
                bk_name = st.text_input("教材名称*", placeholder="必填", key="bk_name")
                bk_isbn = st.text_input("书号(ISBN)", key="bk_isbn")
            with cols_book[1]:
                bk_publisher = st.text_input("出版社", key="bk_pub")
                bk_editor = st.text_input("主编", key="bk_ed")
            with cols_book[2]:
                bk_price = st.number_input("单价(元)*", min_value=0.0, value=0.0, step=0.01, format="%.2f", key="bk_price")
                bk_course = st.text_input("课程", key="bk_course")

        st.caption("🔹 第三步：选择征订范围")
        sem_id = st.selectbox("学期*", [(r["id"], r["name"]) for _, r in semesters.iterrows()],
            format_func=lambda x: x[1], index=next((i for i, s in enumerate([(r["id"], r["name"]) for _, r in semesters.iterrows()]) if s[0] == defaults.get("semester_id")), 0))
        sel_grades = st.multiselect("年级（可多选）", options=get_filtered_grades(), default=[defaults.get("grade")] if defaults.get("grade") else [], placeholder="请选择年级")
        # 学院：基于所选年级级联过滤
        if sel_grades:
            grade_where = " OR ".join(["grade = %s"] * len(sel_grades))
            college_opts = get_filtered_list("students", "college", grade_where, tuple(sel_grades))
        else:
            college_opts = get_filtered_colleges()
        sel_colleges = st.multiselect("学院（可多选）", options=college_opts, default=[defaults.get("college")] if defaults.get("college") else [], placeholder="请选择学院")

        # 专业：基于所选年级+学院级联过滤
        major_where = "1=1"; major_params = []
        if sel_grades:
            major_where += " AND (" + " OR ".join(["grade = %s"] * len(sel_grades)) + ")"
            major_params.extend(sel_grades)
        if sel_colleges:
            major_where += " AND (" + " OR ".join(["college = %s"] * len(sel_colleges)) + ")"
            major_params.extend(sel_colleges)
        major_opts = get_filtered_list("students", "major", major_where, tuple(major_params)) if major_params else get_filtered_majors()
        sel_majors = st.multiselect("专业（可多选）", options=major_opts, default=[defaults.get("major")] if defaults.get("major") else [], placeholder="请选择专业")

        # 班级：基于所选年级+学院+专业级联过滤
        class_where = "1=1"; class_params = []
        if sel_grades:
            class_where += " AND (" + " OR ".join(["grade = %s"] * len(sel_grades)) + ")"
            class_params.extend(sel_grades)
        if sel_colleges:
            class_where += " AND (" + " OR ".join(["college = %s"] * len(sel_colleges)) + ")"
            class_params.extend(sel_colleges)
        if sel_majors:
            class_where += " AND (" + " OR ".join(["major = %s"] * len(sel_majors)) + ")"
            class_params.extend(sel_majors)
        class_opts = get_filtered_list("students", "class_name", class_where, tuple(class_params)) if class_params else get_filtered_class_names()

        sel_classes = st.multiselect("班级（可多选）", options=class_opts, default=[defaults.get("class_name")] if defaults.get("class_name") else [], placeholder="请选择班级")

        # 显示各班级实际人数，支持按人数自动计算征订数量
        use_auto_qty = st.checkbox("📊 按各班级实际人数自动计算征订数量", value=False, key="auto_qty")
        if use_auto_qty and sel_classes:
            # 获取所选班级的实际人数分布
            auto_where = "1=1"; auto_params = []
            if sel_grades:
                auto_where += " AND (" + " OR ".join(["grade = %s"] * len(sel_grades)) + ")"
                auto_params.extend(sel_grades)
            if sel_colleges:
                auto_where += " AND (" + " OR ".join(["college = %s"] * len(sel_colleges)) + ")"
                auto_params.extend(sel_colleges)
            if sel_majors:
                auto_where += " AND (" + " OR ".join(["major = %s"] * len(sel_majors)) + ")"
                auto_params.extend(sel_majors)
            cls = list(set(sel_classes))
            ph = ",".join(["%s"] * len(cls))
            auto_where += f" AND class_name IN ({ph})"; auto_params.extend(cls)
            class_dist = get_filtered_list("students", "class_name", auto_where, tuple(auto_params)) if auto_params else sel_classes
            # 获取各班的实际学生数
            cnt_df = get_class_student_counts(class_names=sel_classes)
            if not cnt_df.empty:
                total_actual_students = int(cnt_df["student_count"].sum())
                auto_teacher = st.number_input("👨‍🏫 教师用书数量", min_value=0, value=0, step=1, key="auto_teacher")
                total_for_all = total_actual_students + auto_teacher
                # 按班级人数分配（每人1本）
                splits = [(row["class_name"], int(row["student_count"]), int(row["student_count"]))
                          for _, row in cnt_df.iterrows()]
                st.info(f"👥 所选 **{len(splits)}** 个班共 **{total_actual_students}** 名学生，合计征订 **{total_for_all}** 本")
                # 显示各班明细
                summary_rows = []
                for cn, sc, alloc in splits:
                    summary_rows.append({"班级": cn, "学生人数": sc, "分配数量": alloc})
                summary_df = pd.DataFrame(summary_rows)
                styled_dataframe(summary_df, hide_ids=True)
                o_qty = st.number_input("征订数量（每班）", min_value=0, value=total_for_all, step=1,
                                         help="系统自动按各班人数计算（每人1本）")
            else:
                o_qty = st.number_input("征订数量", min_value=0, value=int(defaults.get("quantity", 0)), step=1)
        else:
            o_qty = st.number_input("征订数量", min_value=0, value=int(defaults.get("quantity", 0)), step=1)
        o_remark = st.text_input("备注", value=defaults.get("remark", ""))

        combos = max(len(sel_grades) or 1, len(sel_colleges) or 1, len(sel_majors) or 1, len(sel_classes) or 1)
        if edit_oid == 0 and combos > 1:
            st.info(f"📌 将新增 **{combos}** 条征订记录")

        col_btn1, col_btn2 = st.columns(2)
        with col_btn1: submitted = st.button("💾 保存", use_container_width=True, type="primary")
        with col_btn2: delete_btn = st.button("🗑️ 删除", use_container_width=True) if edit_oid > 0 else False

        if submitted:
            if is_new_book and not bk_name:
                st.error("❌ 新增教材时，教材名称为必填")
                st.stop()
            # 1) 处理教材主表
            if is_new_book:
                execute_sql("INSERT INTO textbooks_master (name,isbn,publisher,editor,price,course_name) VALUES (%s,%s,%s,%s,%s,%s)",
                    (bk_name, bk_isbn, bk_publisher, bk_editor, bk_price, bk_course))
                conn = sqlite3.connect(get_sqlite_path())
                cur = conn.cursor()
                cur.execute("SELECT MAX(id) as mid FROM textbooks_master")
                mid = cur.fetchone()[0]
                conn.close()
            else:
                mid = picked[0]

            if edit_oid > 0:
                execute_sql("UPDATE textbook_orders SET semester_id=%s,grade=%s,college=%s,major=%s,class_name=%s,quantity=%s,remark=%s WHERE id=%s",
                    (sem_id[0], (sel_grades or [None])[0], (sel_colleges or [None])[0], (sel_majors or [None])[0], (sel_classes or [None])[0], o_qty, o_remark, edit_oid))
                st.success("✅ 征订已更新")
            else:
                # 如果启用了自动计算，构建班级→数量的映射表（使用表单内的变量）
                auto_qty_map = {}
                if use_auto_qty and sel_classes and 'cnt_df' in dir() and not cnt_df.empty:
                    splits_local = [(row["class_name"], int(row["student_count"]), int(row["student_count"]))
                                     for _, row in cnt_df.iterrows()]
                    for cn, sc, alloc in splits_local:
                        auto_qty_map[cn] = alloc

                grades, colleges, majors, classes = sel_grades or [None], sel_colleges or [None], sel_majors or [None], sel_classes or [None]
                cnt = 0
                for g in grades:
                    for c in colleges:
                        for m in majors:
                            for cl in classes:
                                # 如果有按班级分配的数量，使用之；否则用统一的 o_qty
                                final_qty = auto_qty_map.get(cl, o_qty)
                                execute_sql("INSERT INTO textbook_orders (semester_id,textbook_id,grade,college,major,class_name,quantity,remark) VALUES (%s,%s,%s,%s,%s,%s,%s,%s)",
                                    (sem_id[0], mid, g, c, m, cl, final_qty, o_remark))
                                # 同步到旧textbooks表保持兼容
                                execute_sql("INSERT INTO textbooks (semester_id,grade,college,major,class_name,name,isbn,publisher,editor,price,course_name,quantity,remark) VALUES (%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s)",
                                    (sem_id[0], g, c, m, cl, picked[1] if not is_new_book else bk_name,
                                     bk_isbn if is_new_book else next((r.get("isbn","") for _,r in master_books.iterrows() if r["id"]==picked[0]), ""),
                                     bk_publisher if is_new_book else next((r.get("publisher","") for _,r in master_books.iterrows() if r["id"]==picked[0]), ""),
                                     bk_editor if is_new_book else next((r.get("editor","") for _,r in master_books.iterrows() if r["id"]==picked[0]), ""),
                                     bk_price if is_new_book else next((r["price"] for _,r in master_books.iterrows() if r["id"]==picked[0]), 0),
                                     bk_course if is_new_book else next((r.get("course_name","") for _,r in master_books.iterrows() if r["id"]==picked[0]), ""),
                                     final_qty, o_remark))
                                cnt += 1
                st.success(f"✅ 已添加 {cnt} 条征订记录")
            st.rerun()
        if delete_btn:
            execute_sql("DELETE FROM textbook_orders WHERE id=%s", (edit_oid,))
            st.success("✅ 已删除")
            st.rerun()

    with tab3:
        st.markdown("#### 📥 从 Excel 导入教材征订")
        with st.expander("📄 模板下载与说明", expanded=False):
            template_cols = ["学期", "年级", "学院", "专业", "班级", "教材名称", "书号(ISBN)", "出版社", "主编", "单价(元)", "征订数量", "备注"]
            template_df = make_template_df(template_cols)
            st.download_button("📄 下载导入模板", data=excel_export(template_df, "教材征订导入模板"),
                file_name="教材征订导入模板.xlsx", mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                use_container_width=True, type="secondary")
        uploaded = st.file_uploader("选择 Excel 文件", type=["xlsx", "xls"], key="tb_upload3")
        if uploaded:
            try:
                raw_df = read_excel_upload(uploaded)
                st.info(f"📄 检测到 {len(raw_df)} 行")
                col_map = {}
                for col in raw_df.columns:
                    cl = col.strip().lower()
                    if "教材" in cl: col_map["name"] = col
                    elif "学期" in cl: col_map["semester_name"] = col
                    elif "年级" in cl: col_map["grade"] = col
                    elif "学院" in cl: col_map["college"] = col
                    elif "专业" in cl and "年级" in cl: col_map["major_grade"] = col  # 专业年级（合并列）
                    elif "专业" in cl: col_map["major"] = col
                    elif "班级" in cl: col_map["class_name"] = col
                    elif "学生人数" in cl or "人数" in cl: col_map["student_count"] = col  # 学生人数
                    elif "教师" in cl: col_map["teacher_books"] = col  # 教师用书
                    elif "单价" in cl: col_map["price"] = col
                    elif "数量" in cl or "合计" in cl: col_map["quantity"] = col
                    elif "书号" in cl: col_map["isbn"] = col
                    elif "出版社" in cl: col_map["publisher"] = col
                    elif "主编" in cl or "作者" in cl: col_map["editor"] = col
                    elif "备注" in cl or "说明" in cl: col_map["remark"] = col
                if "name" not in col_map or "semester_name" not in col_map:
                    st.error("❌ 缺少必要列：教材名称、学期")
                else:
                    # 检测是否有"专业年级"合并列，有则提示自动拆班
                    has_major_grade = "major_grade" in col_map
                    if has_major_grade:
                        st.info("📌 检测到「专业年级」列，导入时将自动按各班级实际人数分摊征订数量")

                    if st.button("✅ 确认导入", use_container_width=True, type="primary", key="tb_import3"):
                        mapped = raw_df.rename(columns={v: k for k, v in col_map.items()})
                        # 统一列顺序
                        col_order = ["semester_name","grade","college","major","class_name","major_grade","name","isbn","publisher","editor","price","quantity","student_count","teacher_books","remark"]
                        mapped = mapped[[c for c in col_order if c in mapped.columns]]
                        sem_map = {r["name"]: r["id"] for _, r in semesters.iterrows()}
                        success, errors, total_rows = 0, [], len(mapped)
                        progress_bar = st.progress(0, text="正在导入...")

                        for i, (_, row) in enumerate(mapped.iterrows()):
                            sid = sem_map.get(str(row["semester_name"]))
                            if sid is None: errors.append(f"学期「{row['semester_name']}」不存在"); continue
                            try:
                                nm, p = safe_str(row["name"]), safe_float(row.get("price", 0))
                                # 查找或创建教材主记录
                                old = query_df("SELECT id FROM textbooks_master WHERE name=%s AND COALESCE(isbn,'')=%s", (nm, str(row.get("isbn","") or "")))
                                if old.empty:
                                    execute_sql("INSERT INTO textbooks_master (name,isbn,publisher,editor,price) VALUES (%s,%s,%s,%s,%s)",
                                        (nm, str(row.get("isbn","") or ""), str(row.get("publisher","") or ""), str(row.get("editor","") or ""), p))
                                    conn2 = sqlite3.connect(get_sqlite_path())
                                    c2 = conn2.cursor(); c2.execute("SELECT MAX(id) FROM textbooks_master"); tid = c2.fetchone()[0]; conn2.close()
                                else:
                                    tid = old.iloc[0]["id"]

                                college_val = str(row.get("college","") or "")
                                major_val = str(row.get("major","") or "")
                                class_val = str(row.get("class_name","") or "")
                                grade_val = str(row.get("grade","") or "")
                                mg_val = str(row.get("major_grade","") or "")
                                total_qty = safe_int(row.get("quantity", 0))

                                # 如果启用了自动拆班且有专业年级列
                                if has_major_grade and mg_val and total_qty > 0:
                                    # 尝试在 students 表中查找匹配的班级
                                    # 用 college + 从专业年级提取的专业名来过滤
                                    mg_major = ""
                                    # 尝试从专业年级提取专业名（如"电商1241-1244"→"电商"）
                                    mg_parts = re.split(r'(\d)', mg_val, maxsplit=1)
                                    if mg_parts:
                                        mg_major = mg_parts[0].strip()
                                    # 用学院+提取的专业名查找各班人数
                                    class_df = get_class_student_counts(college=college_val, major=major_val or mg_major)
                                    if not class_df.empty:
                                        total_students = int(class_df["student_count"].sum())
                                        teacher_cnt = safe_int(row.get("teacher_books", 0))
                                        # 总征订量 = 学生人数 + 教师用书（如果未单独指定教师用书列，就用合计数量）
                                        if "teacher_books" in row and pd.notna(row.get("teacher_books")):
                                            total_for_split = total_qty
                                        else:
                                            total_for_split = total_students + teacher_cnt if teacher_cnt > 0 else total_qty
                                        splits = [(row["class_name"], int(row["student_count"]), int(row["student_count"]))
                                                   for _, row in class_df.iterrows()]
                                        for cn, sc, alloc in splits:
                                            execute_sql("INSERT INTO textbook_orders (semester_id,textbook_id,grade,college,major,class_name,quantity,remark) VALUES (%s,%s,%s,%s,%s,%s,%s,%s)",
                                                (sid, tid, grade_val or mg_major, college_val, major_val or mg_major, cn, alloc, str(row.get("remark","") or "")))
                                            execute_sql("INSERT INTO textbooks (semester_id,grade,college,major,class_name,name,isbn,publisher,editor,price,course_name,quantity,remark) VALUES (%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s)",
                                                (sid, grade_val or mg_major, college_val, major_val or mg_major, cn,
                                                 nm, str(row.get("isbn","") or ""), str(row.get("publisher","") or ""), str(row.get("editor","") or ""),
                                                 p, str(row.get("course_name","") or ""), alloc, str(row.get("remark","") or "")))
                                        success += 1
                                    else:
                                        # 数据库中没有匹配的班级，按原始数据导入
                                        execute_sql("INSERT INTO textbook_orders (semester_id,textbook_id,grade,college,major,class_name,quantity,remark) VALUES (%s,%s,%s,%s,%s,%s,%s,%s)",
                                            (sid, tid, grade_val, college_val, major_val, mg_val, total_qty, str(row.get("remark","") or "")))
                                        execute_sql("INSERT INTO textbooks (semester_id,grade,college,major,class_name,name,isbn,publisher,editor,price,course_name,quantity,remark) VALUES (%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s)",
                                            (sid, grade_val, college_val, major_val, mg_val,
                                             nm, str(row.get("isbn","") or ""), str(row.get("publisher","") or ""), str(row.get("editor","") or ""),
                                             p, str(row.get("course_name","") or ""), total_qty, str(row.get("remark","") or "")))
                                        success += 1
                                else:
                                    # 标准导入逻辑（已有班级列）
                                    execute_sql("INSERT INTO textbook_orders (semester_id,textbook_id,grade,college,major,class_name,quantity,remark) VALUES (%s,%s,%s,%s,%s,%s,%s,%s)",
                                        (sid, tid, grade_val, college_val, major_val, class_val, total_qty, str(row.get("remark","") or "")))
                                    execute_sql("INSERT INTO textbooks (semester_id,grade,college,major,class_name,name,isbn,publisher,editor,price,course_name,quantity,remark) VALUES (%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s)",
                                        (sid, grade_val, college_val, major_val, class_val,
                                         nm, str(row.get("isbn","") or ""), str(row.get("publisher","") or ""), str(row.get("editor","") or ""),
                                         p, str(row.get("course_name","") or ""), total_qty, str(row.get("remark","") or "")))
                                    success += 1
                            except Exception as e: errors.append(f"{nm}: {str(e)[:100]}")
                            progress_bar.progress((i + 1) / total_rows, text=f"已处理 {i+1}/{total_rows}")
                        # 统一展示结果
                        mc_ok2, mc_err2 = st.columns(2)
                        mc_ok2.metric("✅ 导入成功", success)
                        mc_err2.metric("⚠️ 导入失败", len(errors))
                        if errors:
                            with st.expander(f"查看 {len(errors)} 条失败详情"):
                                for e in errors:
                                    st.caption(f"• {e}")
                        write_import_log("教材征订导入", uploaded.name, total_rows, success, errors)
            except Exception as e:
                st.error(f"❌ 读取失败：{e}")

    with tab4:
        st.markdown("#### 📚 教材批量分配")
        st.caption("选择一本教材，批量分配给多个年级/学院/专业/班级")

        be_sem = st.selectbox("① 选择学期", [(r["id"], r["name"]) for _, r in semesters.iterrows()],
            format_func=lambda x: x[1], key="be_sem")
        be_book = st.selectbox("② 选择教材", master_options, format_func=lambda x: x[1], key="be_book")
        st.caption("③~⑥ 至少选一项，不选则视为全部")

        with st.container():
            col_a, col_b = st.columns(2)
            with col_a:
                be_colleges = st.multiselect("③ 学院（可多选）", options=get_filtered_colleges(), key="be_cl", placeholder="请选择学院")
                be_grades = st.multiselect("④ 年级（可多选）", options=get_filtered_grades(), key="be_gr", placeholder="请选择年级")
            with col_b:
                # 专业：根据所选学院级联过滤
                if be_colleges:
                    be_major_where = " OR ".join(["college = %s"] * len(be_colleges))
                    be_major_opts = get_filtered_list("students", "major", be_major_where, tuple(be_colleges))
                else:
                    be_major_opts = get_filtered_majors()
                be_majors = st.multiselect("⑤ 专业（可多选）", options=be_major_opts, key="be_mj", placeholder="请选择专业")
                # 班级：根据所选学院+专业级联过滤
                be_class_where = "1=1"; be_class_params = []
                if be_colleges:
                    be_class_where += " AND (" + " OR ".join(["college = %s"] * len(be_colleges)) + ")"
                    be_class_params.extend(be_colleges)
                if be_majors:
                    be_class_where += " AND (" + " OR ".join(["major = %s"] * len(be_majors)) + ")"
                    be_class_params.extend(be_majors)
                be_class_opts = get_filtered_list("students", "class_name", be_class_where, tuple(be_class_params)) if be_class_params else get_filtered_class_names()
                be_classes = st.multiselect("⑥ 班级（可多选）", options=be_class_opts, key="be_cls", placeholder="请选择班级")

        be_qty = st.number_input("⑦ 征订数量", min_value=0, value=1, step=1, key="be_qty")

        if be_book[0] > 0:
            grades = be_grades or [None]
            colleges = be_colleges or [None]
            majors = be_majors or [None]
            classes = be_classes or [None]
            total_gen = len(grades) * len(colleges) * len(majors) * len(classes)

            if total_gen > 0:
                st.info(f"📌 将生成 **{total_gen}** 条征订记录")

                if st.button("✅ 确认批量分配", use_container_width=True, type="primary", key="be_confirm"):
                    cnt = 0
                    for g in grades:
                        for c in colleges:
                            for m in majors:
                                for cl in classes:
                                    execute_sql(
                                        "INSERT INTO textbook_orders (semester_id,textbook_id,grade,college,major,class_name,quantity) VALUES (%s,%s,%s,%s,%s,%s,%s)",
                                        (be_sem[0], be_book[0], g, c, m, cl, be_qty))
                                    cnt += 1
                    st.success(f"✅ 已生成 {cnt} 条征订记录")
                    write_import_log("教材批量分配", be_book[1], total_gen, cnt, [])
                    st.rerun()
        else:
            st.warning("请先选择教材（不能选择「新增教材」）")

    # 底部快捷入口：教材表管理
    st.divider()
    st.caption("💡 如需添加/编辑教材信息，请前往侧边栏「📖 教材表管理」")
