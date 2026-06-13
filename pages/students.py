"""
学生管理页面
==============

学生管理功能，包括学生列表查看、新增/编辑学生、Excel 批量导入导出。
"""

import streamlit as st
import pandas as pd
import math
from datetime import date

from components import show_header, excel_export, excel_export_by_class, styled_dataframe
from utils import (
    get_filtered_list, get_filtered_grades, get_filtered_majors,
    get_filtered_colleges, safe_str, safe_field, make_template_df,
    read_excel_upload, write_import_log
)
from database import query_df, execute_sql


def student_management():
    show_header("👨‍🎓 学生管理", "管理学生基础信息，支持 Excel 批量导入导出")

    tab1, tab2, tab3 = st.tabs(["📋 学生列表", "➕ 新增/编辑", "📥 导入 Excel"])

    # ── Tab 1: 列表 ──
    with tab1:
        with st.expander("🔍 筛选条件", expanded=True):
            col1, col2, col3, col4, col5 = st.columns(5)
            with col1:
                f_grade = st.selectbox("年级", ["全部"] + get_filtered_grades(), key="s_grade")
            with col2:
                # 学院：基于所选年级级联过滤
                if f_grade != "全部":
                    college_filter = get_filtered_list("students", "college", "grade = %s", (f_grade,))
                else:
                    college_filter = get_filtered_colleges()
                f_college = st.selectbox("学院", ["全部"] + college_filter, key="s_college")
            with col3:
                # 专业：基于所选年级+学院级联过滤
                s_major_where = "1=1"; s_major_params = []
                if f_grade != "全部":
                    s_major_where += " AND grade = %s"; s_major_params.append(f_grade)
                if f_college != "全部":
                    s_major_where += " AND college = %s"; s_major_params.append(f_college)
                major_options = get_filtered_list("students", "major", s_major_where, tuple(s_major_params)) if s_major_params else get_filtered_majors()
                f_major = st.selectbox("专业", ["全部"] + major_options, key="s_major")
            with col4:
                # 班级：基于所选年级+学院+专业级联过滤
                f_where = "1=1"
                f_params = []
                if f_grade != "全部":
                    f_where += " AND grade = %s"; f_params.append(f_grade)
                if f_college != "全部":
                    f_where += " AND college = %s"; f_params.append(f_college)
                if f_major != "全部":
                    f_where += " AND major = %s"; f_params.append(f_major)
                class_options = get_filtered_list("students", "class_name", f_where, tuple(f_params))
                f_class = st.selectbox("班级", ["全部"] + class_options, key="s_class")
            with col5:
                search = st.text_input("🔍 搜索", key="s_search", placeholder="姓名/学号/身份证...")

        sql = "SELECT id, id_card, student_id, name, grade, college, major, class_name FROM students WHERE 1=1"
        params = []
        for val, col in [(f_grade, "grade"), (f_college, "college"), (f_major, "major"), (f_class, "class_name")]:
            if val != "全部":
                sql += f" AND {col} = %s"; params.append(val)
        if search:
            sql += " AND (name LIKE %s OR student_id LIKE %s OR id_card LIKE %s)"
            like = f"%{search}%"; params.extend([like, like, like])
        sql += " ORDER BY grade, class_name, name"

        df = query_df(sql, tuple(params) if params else None)
        total = len(df)

        # ── 分页（从 session_state 读取，控件移到表格下方）──
        page_size = st.session_state.get("s_ps", 50)
        page = st.session_state.get("s_pg", 1)
        total_pages = max(1, math.ceil(total / page_size))
        # 页码保护：删除后当前页可能越界，自动跳到最后页
        if page > total_pages:
            page = total_pages
            st.session_state.s_pg = total_pages

        # 导出按钮保持在顶部右侧
        _, exp_col = st.columns([4, 1])
        with exp_col:
            excel_data = excel_export_by_class(
                df.rename(columns={
                    "id_card": "身份证号", "student_id": "学号", "name": "姓名",
                    "grade": "年级", "college": "学院", "major": "专业", "class_name": "班级"
                }),
                class_col="班级", file_prefix="学生名单"
            ) if not df.empty else b""
            st.download_button("📥 导出（按班级分页）", data=excel_data,
                               file_name=f"学生名单_{date.today()}.xlsx",
                               mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                               use_container_width=True, disabled=df.empty)

        st.caption(f"共 **{total}** 名学生 ｜ 第 {page}/{total_pages} 页")

        if not df.empty:
            # 当前页数据
            start_idx = (page - 1) * page_size
            end_idx = min(start_idx + page_size, total)
            page_df = df.iloc[start_idx:end_idx].copy()

            # 保存原始数据用于对比
            edit_cols = ["name", "id_card", "grade", "college", "major", "class_name"]
            orig_key = "_orig"
            page_df[orig_key] = page_df[edit_cols].to_dict("records")

            # 全选状态：如果之前勾选了全选，默认全勾上
            select_all_init = st.session_state.get(f"s_selall_{page}", False)
            page_df["选择"] = select_all_init
            display_df = page_df.rename(columns={
                "id_card": "身份证号", "student_id": "学号", "name": "姓名",
                "grade": "年级", "college": "学院", "major": "专业", "class_name": "班级"
            })

            edited = st.data_editor(
                display_df,
                use_container_width=True,
                hide_index=True,
                column_order=["选择", "学号", "姓名", "身份证号", "年级", "学院", "专业", "班级"],
                column_config={
                    "选择": st.column_config.CheckboxColumn("选择"),
                    "学号": st.column_config.TextColumn("学号", disabled=True, alignment="center"),
                    "姓名": st.column_config.TextColumn("姓名", alignment="center"),
                    "身份证号": st.column_config.TextColumn("身份证号", alignment="center"),
                    "年级": st.column_config.TextColumn("年级", alignment="center"),
                    "学院": st.column_config.TextColumn("学院", alignment="center"),
                    "专业": st.column_config.TextColumn("专业", alignment="center"),
                    "班级": st.column_config.TextColumn("班级", alignment="center"),
                },
                disabled=["学号"],
                key=f"stu_editor_{page}"
            )

            # ═══ 布局：上行(全选+分页) + 下行(红删蓝存) ═══
            # 检测修改
            changes = []
            for i in range(len(edited)):
                orig = page_df.iloc[i][orig_key]
                for col in edit_cols:
                    cn = {"name":"姓名","id_card":"身份证号","grade":"年级","college":"学院","major":"专业","class_name":"班级"}[col]
                    new_val = edited.iloc[i].get(cn)
                    old_val = orig.get(col)
                    if str(new_val or "") != str(old_val or ""):
                        changes.append((page_df.iloc[i]["id"], col, new_val))

            # 处理按钮触发的操作（用 session_state 代替 query_params，避免 form 提交导致会话丢失）
            pending = st.session_state.get("pending_action", "")
            if pending == "s_save":
                if changes:
                    for sid, col, val in changes:
                        execute_sql(f"UPDATE students SET {col}=%s WHERE id=%s", (val, sid))
                    st.toast(f"✅ 已保存 {len(changes)} 处修改", icon="✅")
                st.session_state.pending_action = ""
                st.rerun()
            elif pending == "s_del":
                select_all = st.session_state.get(f"s_selall_{page}", False)
                if select_all:
                    del_ids = tuple(page_df["id"].tolist())
                else:
                    selected = edited[edited["选择"] == True].index if "选择" in edited.columns else []
                    del_ids = tuple(page_df.iloc[selected]["id"].tolist()) if len(selected) > 0 else ()
                if del_ids:
                    if len(del_ids) == 1:
                        execute_sql("DELETE FROM students WHERE id = %s", (del_ids[0],))
                    else:
                        placeholders = ",".join(["%s"] * len(del_ids))
                        execute_sql(f"DELETE FROM students WHERE id IN ({placeholders})", del_ids)
                    st.toast(f"✅ 已删除 {len(del_ids)} 名学生", icon="✅")
                st.session_state.pending_action = ""
                st.rerun()

            # ── 上行：全选 + 分页 ──
            r1_sel, r1_info, r1_ps, r1_prev, r1_num, r1_next = st.columns([1.5, 2, 1, 0.7, 0.7, 0.7])
            with r1_sel:
                select_all = st.checkbox("全选本页", key=f"s_selall_{page}",
                    help="勾选后删除操作将应用于本页全部记录")
            with r1_info:
                st.caption(f"共 **{total}** 名学生")
            with r1_ps:
                st.selectbox("每页", [50, 100, 200, 500], key="s_ps",
                             on_change=lambda: st.session_state.update({"s_pg": 1}),
                             label_visibility="collapsed")
            with r1_prev:
                if st.button("◀", key="s_pp", disabled=(page <= 1), use_container_width=True):
                    st.session_state.s_pg = page - 1; st.rerun()
            with r1_num:
                st.markdown(f"<div style='text-align:center;padding-top:5px;font-weight:500'>{page}/{total_pages}</div>", unsafe_allow_html=True)
            with r1_next:
                if st.button("▶", key="s_np", disabled=(page >= total_pages), use_container_width=True):
                    st.session_state.s_pg = page + 1; st.rerun()

            # ── 下行：删除(红) + 保存(蓝) ──
            del_count = len(page_df) if select_all else (int(edited["选择"].sum()) if "选择" in edited.columns else 0)
            del_disabled = "disabled" if del_count == 0 else ""
            save_disabled = "disabled" if not changes else ""
            del_btn_label = f"🗑️ 删除（{del_count}人）"
            save_btn_label = f"💾 保存修改（{len(changes)}处）" if changes else "💾 保存修改"

            r2_del, r2_save = st.columns([1, 1])
            with r2_del:
                st.button(del_btn_label, key=f"s_del_btn_{page}", type="secondary",
                          disabled=(del_count == 0), use_container_width=True,
                          on_click=lambda: st.session_state.update({"pending_action": "s_del"}))
            with r2_save:
                st.button(save_btn_label, key=f"s_save_btn_{page}", type="primary",
                          disabled=(not changes), use_container_width=True,
                          on_click=lambda: st.session_state.update({"pending_action": "s_save"}))

    # ── Tab 2: 新增/编辑 ──
    with tab2:
        st.markdown("#### ➕ 新增 / 编辑学生")

        edit_id = st.number_input("编辑学生 ID（留空为新增）", min_value=0, value=0, step=1, key="edit_sid")
        defaults = {}
        if edit_id > 0:
            existing = query_df("SELECT * FROM students WHERE id = %s", (edit_id,))
            if not existing.empty:
                defaults = existing.iloc[0].to_dict()
                st.info(f"正在编辑：**{defaults.get('name', '')}** ({defaults.get('student_id', '')})")

        with st.form("student_form"):
            c1, c2 = st.columns(2)
            with c1:
                name = st.text_input("姓名 *", value=defaults.get("name", ""), placeholder="必填")
                id_card = st.text_input("身份证号 *", value=defaults.get("id_card", ""),
                                        max_chars=18, placeholder="必填，18位")
                student_id = st.text_input("学号 *", value=defaults.get("student_id", ""), placeholder="必填")
            with c2:
                grade = st.text_input("年级", value=defaults.get("grade", ""), placeholder="如：2024级")
                college = st.text_input("学院", value=defaults.get("college", ""))
                major = st.text_input("专业", value=defaults.get("major", ""))
            class_name = st.text_input("班级", value=defaults.get("class_name", ""), placeholder="如：软件工程1班")

            col_btn1, col_btn2 = st.columns(2)
            with col_btn1:
                submitted = st.form_submit_button("💾 保存", use_container_width=True)
            with col_btn2:
                delete_btn = st.form_submit_button("🗑️ 删除", use_container_width=True) if edit_id > 0 else False

            if submitted:
                if not all([name, id_card, student_id]):
                    st.error("❌ 姓名、身份证号、学号为必填项")
                else:
                    try:
                        if edit_id > 0:
                            execute_sql(
                                """UPDATE students SET name=%s,id_card=%s,student_id=%s,
                                   grade=%s,college=%s,major=%s,class_name=%s WHERE id=%s""",
                                (name, id_card, student_id, grade, college, major, class_name, edit_id))
                            st.success("✅ 已更新")
                        else:
                            execute_sql(
                                """INSERT INTO students (name,id_card,student_id,grade,college,major,class_name)
                                   VALUES (%s,%s,%s,%s,%s,%s,%s)""",
                                (name, id_card, student_id, grade, college, major, class_name))
                            st.success("✅ 已添加")
                        st.rerun()
                    except Exception as e:
                        if "UNIQUE" in str(e) or "unique" in str(e).lower():
                            st.error("❌ 身份证号或学号已存在，不能重复")
                        else:
                            st.error(f"保存失败：{e}")

            if delete_btn:
                execute_sql("DELETE FROM students WHERE id = %s", (edit_id,))
                st.success("✅ 已删除")

    # ── Tab 3: 导入 Excel ──
    with tab3:
        st.markdown("#### 📥 从 Excel 导入学生")

        # 下载模板
        with st.expander("📄 模板下载与说明", expanded=False):
            template_cols = ["身份证号", "学号", "姓名", "年级", "学院", "专业", "班级"]
            template_df = make_template_df(template_cols)
            template_bytes = excel_export(template_df, "学生导入模板")
            st.download_button(
                "📄 下载导入模板",
                data=template_bytes,
                file_name="学生导入模板.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                use_container_width=True,
                type="secondary"
            )
            st.caption("Excel 表头建议包含：身份证号、学号、姓名、年级、学院、专业、班级")

        uploaded = st.file_uploader("选择 Excel 文件", type=["xlsx", "xls"], key="stu_upload")
        if uploaded:
            try:
                raw_df = read_excel_upload(uploaded)
                st.info(f"📄 检测到 {len(raw_df)} 行，列名：{', '.join(raw_df.columns[:8])}")

                col_map = {}
                for col in raw_df.columns:
                    cl = col.strip().lower()
                    if "身份证" in cl or "id_card" in cl: col_map["id_card"] = col
                    elif "学号" in cl or "student_id" in cl or "编号" in cl: col_map["student_id"] = col
                    elif "姓名" in cl or "name" in cl: col_map["name"] = col
                    elif "年级" in cl or "grade" in cl: col_map["grade"] = col
                    elif "学院" in cl or "college" in cl: col_map["college"] = col
                    elif "专业" in cl or "major" in cl: col_map["major"] = col
                    elif "班级" in cl or "class" in cl: col_map["class_name"] = col

                if not all(k in col_map for k in ["id_card", "student_id", "name"]):
                    st.error("❌ 缺少必要列：身份证号、学号、姓名")
                    st.json(col_map)
                else:
                    preview = raw_df.head(5)
                    styled_dataframe(preview, hide_ids=True)

                    if st.button("✅ 确认导入", use_container_width=True, type="primary"):
                        mapped = raw_df.rename(columns={v: k for k, v in col_map.items()})
                        wanted = [c for c in ["id_card","student_id","name","grade","college","major","class_name"] if c in mapped.columns]
                        mapped = mapped[wanted].where(pd.notnull(mapped), None)

                        # 处理数字型字段（2023 → "2023"，避免变成 "2023.0"）
                        for text_col in ["grade", "college", "major", "class_name"]:
                            if text_col in mapped.columns:
                                mapped[text_col] = mapped[text_col].apply(
                                    lambda x: safe_field(x) if pd.notna(x) else None
                                )

                        success_count = 0
                        update_count = 0
                        errors = []
                        progress = st.progress(0)
                        total_rows = len(mapped)
                        for i, (_, row) in enumerate(mapped.iterrows()):
                            try:
                                # 使用 INSERT OR REPLACE 处理重复：学生已存在则更新全部字段
                                execute_sql(
                                    """INSERT OR REPLACE INTO students
                                       (id, id_card, student_id, name, grade, college, major, class_name)
                                       VALUES (
                                           (SELECT id FROM students WHERE student_id=%s),
                                           %s, %s, %s, %s, %s, %s, %s
                                       )""",
                                    (safe_str(row["student_id"]),
                                     safe_str(row["id_card"]), safe_str(row["student_id"]), safe_str(row["name"]),
                                     row.get("grade"), row.get("college"), row.get("major"), row.get("class_name")))
                                success_count += 1
                            except Exception as e:
                                errors.append(f"{row.get('name','?')}: {str(e)[:80]}")
                            progress.progress((i + 1) / total_rows)

                        # 写日志
                        write_import_log(
                            module="学生管理",
                            filename=uploaded.name,
                            total=total_rows,
                            success=success_count,
                            errors=errors
                        )

                        # 展示结果
                        sm_ok, sm_err = st.columns(2)
                        sm_ok.metric("✅ 导入成功", success_count)
                        sm_err.metric("⚠️ 导入失败", len(errors))
                        if errors:
                            with st.expander(f"查看 {len(errors)} 条失败详情"):
                                for err in errors:
                                    st.caption(f"• {err}")
            except Exception as e:
                st.error(f"❌ 读取失败：{e}")
