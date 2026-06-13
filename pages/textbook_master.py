"""
教材表管理页面 (textbook_master.py)
功能：维护教材主表（textbooks_master），供征订时选用
包含：教材列表、新增教材、导入Excel三个标签页
"""

import streamlit as st
import pandas as pd
import math
import sqlite3
import re
from datetime import datetime, timedelta

from database import query_df, execute_sql, get_sqlite_path
from components import show_header, excel_export, styled_dataframe
from utils import (safe_int, safe_float, safe_str, make_template_df,
                   read_excel_upload, write_import_log)


def textbook_master_management():
    """教材表管理 — 维护教材主表（textbooks_master），供征订时选用"""
    show_header("📖 教材表管理", "维护教材库，征订时从库中选择")

    tab1, tab2, tab3 = st.tabs(["📋 教材列表", "➕ 新增教材", "📥 导入 Excel"])

    # ── 引用次数统计 ──
    usage_sql = "SELECT textbook_id, COUNT(*) as cnt FROM textbook_orders GROUP BY textbook_id UNION ALL SELECT textbook_id, COUNT(*) as cnt FROM textbook_subscriptions WHERE textbook_id IS NOT NULL GROUP BY textbook_id"
    usage_df = query_df(usage_sql)
    usage_map = {}
    if not usage_df.empty:
        for _, ur in usage_df.iterrows():
            tid = safe_int(ur["textbook_id"])
            if tid <= 0:
                continue
            usage_map[tid] = usage_map.get(tid, 0) + safe_int(ur["cnt"])

    # ═══════════════ Tab1: 教材列表 ═══════════════
    with tab1:
        # ── 默认折扣率 & 批量操作 ──
        with st.expander("🔧 设置与批量操作", expanded=False):
            default_dr = st.number_input("🔢 导入默认折扣率（如 0.76 = 76折，1.0 = 原价）",
                min_value=0.0, max_value=1.0, value=st.session_state.get("master_default_dr", 0.76),
                step=0.01, format="%.2f", key="master_default_dr",
                help="导入 Excel 时所有教材将自动应用此折扣率，之后可在列表中逐个修改")

            b_col1, b_col2 = st.columns([1, 1])
            with b_col1:
                st.number_input("批量设置折扣率", min_value=0.0, max_value=1.0,
                                value=st.session_state.get("master_default_dr", 0.76),
                                step=0.01, format="%.2f", key="batch_dr_widget",
                                help="勾选教材后，点击右侧按钮批量修改折扣率")
            with b_col2:
                st.markdown("<div style='padding-top:28px;'></div>", unsafe_allow_html=True)
                st.button(f"🔄 批量修改折扣率", key="batch_dr_btn",
                          use_container_width=True, type="secondary",
                          on_click=lambda: st.session_state.update({"pending_action": "batch_dr"}))

        search_key = st.text_input("🔍 搜索教材（名称/ISBN/出版社）", key="master_search", placeholder="输入关键词搜索...")

        # ── 高级筛选 ──
        with st.expander("🔎 高级筛选", expanded=False):
            af_col1, af_col2, af_col3, af_col4 = st.columns([1, 1, 1, 1])
            all_masters = query_df(
                "SELECT DISTINCT publisher, editor, discount_rate, textbook_type FROM textbooks_master"
            )
            pub_opts = sorted(all_masters["publisher"].dropna().astype(str).unique().tolist()) if not all_masters.empty else []
            editor_opts = sorted(all_masters["editor"].dropna().astype(str).unique().tolist()) if not all_masters.empty else []
            type_opts = sorted([t for t in all_masters["textbook_type"].dropna().astype(str).unique().tolist() if t]) if not all_masters.empty else []
            with af_col1:
                filter_pub = st.multiselect("出版社", pub_opts, key="filter_pub", placeholder="全部出版社")
            with af_col2:
                filter_editor = st.multiselect("主编", editor_opts, key="filter_editor", placeholder="全部主编")
            with af_col3:
                filter_dr_mode = st.selectbox("折扣率", ["全部", "有折扣(<1.0)", "无折扣(=1.0)"], key="filter_dr_mode")
                if filter_dr_mode == "全部":
                    filter_dr_range = st.slider("折扣率范围", 0.0, 1.0, (0.0, 1.0), key="filter_dr_slider")
            with af_col4:
                filter_type = st.multiselect("教材类型", type_opts, key="filter_type", placeholder="全部类型")

        # ── 构建查询 ──
        base_sql = "FROM textbooks_master"
        where_parts = []
        params = []

        if search_key.strip():
            where_parts.append("(name LIKE %s OR COALESCE(isbn,'') LIKE %s OR COALESCE(publisher,'') LIKE %s)")
            like_val = f"%{search_key.strip()}%"
            params.extend([like_val, like_val, like_val])

        if filter_pub:
            ph = ",".join(["%s"] * len(filter_pub))
            where_parts.append(f"publisher IN ({ph})")
            params.extend(filter_pub)

        if filter_editor:
            ph = ",".join(["%s"] * len(filter_editor))
            where_parts.append(f"editor IN ({ph})")
            params.extend(filter_editor)

        if filter_dr_mode == "有折扣(<1.0)":
            where_parts.append("COALESCE(discount_rate, 1.0) < 1.0")
        elif filter_dr_mode == "无折扣(=1.0)":
            where_parts.append("COALESCE(discount_rate, 1.0) = 1.0")

        if "filter_dr_slider" in st.session_state and filter_dr_mode == "全部":
            dr_lo, dr_hi = st.session_state.filter_dr_slider
            where_parts.append("COALESCE(discount_rate, 1.0) BETWEEN %s AND %s")
            params.extend([dr_lo, dr_hi])

        if filter_type:
            ph = ",".join(["%s"] * len(filter_type))
            where_parts.append(f"COALESCE(textbook_type,'') IN ({ph})")
            params.extend(filter_type)

        where_sql = (" WHERE " + " AND ".join(where_parts)) if where_parts else ""
        count_sql = f"SELECT COUNT(*) as cnt {base_sql} {where_sql}"
        master_df = query_df(count_sql, tuple(params) if params else None)
        total_count = int(master_df.iloc[0]["cnt"]) if not master_df.empty else 0

        master_sql = f"SELECT id, name, isbn, publisher, editor, price, publication_date, course_name, discount_rate, actual_price, textbook_type {base_sql} {where_sql} ORDER BY name"
        master_df = query_df(master_sql, tuple(params) if params else None)

        if not master_df.empty:
            st.caption(f"共 **{len(master_df)}** 本教材，引用 **{sum(usage_map.values())}** 次")

            page_size = st.session_state.get("master_ps", 50)
            total_pages = max(1, math.ceil(len(master_df) / page_size))
            mp = st.session_state.get("master_page", 1)
            if mp > total_pages:
                mp = total_pages
                st.session_state.master_page = total_pages

            st.caption(f"共 **{len(master_df)}** 本教材 ｜ 第 {mp}/{total_pages} 页")

            start = (mp - 1) * page_size
            end = min(start + page_size, len(master_df))
            page_df = master_df.iloc[start:end].copy()

            master_selall = st.session_state.get(f"master_selall_{mp}", False)
            rows = []
            real_ids = []  # 索引→真实数据库ID 的映射
            for _, mr in page_df.iterrows():
                mid = int(mr["id"])
                real_ids.append(mid)
                usg = usage_map.get(mid, 0)
                dr = mr.get("discount_rate")
                if dr is None:
                    dr_val = 1.0
                else:
                    dr_val = float(dr)
                ap = mr.get("actual_price")
                # NaN ≠ None，必须用 pd.notna 判断，否则 NaN 不走 fallback 计算
                ap_val = float(ap) if ap is not None and not (isinstance(ap, float) and math.isnan(ap)) else None
                tb_type = str(mr.get("textbook_type", "") or "")
                rows.append({
                    "选择": master_selall,
                    "课程名称": str(mr.get("course_name", "") or ""),
                    "教材名称": mr["name"],
                    "书号": str(mr.get("isbn", "") or ""),
                    "主编": str(mr.get("editor", "") or ""),
                    "出版社": str(mr.get("publisher", "") or ""),
                    "出版日期": safe_str(mr.get("publication_date", "")),
                    "单价(元)": float(mr.get("price", 0) or 0),
                    "折扣率": dr_val,
                    "实洋(元)": ap_val if ap_val is not None else float(mr.get("price", 0) or 0) * dr_val,
                    "教材类型": tb_type,
                    "引用次数": usg,
                })

            display_df = pd.DataFrame(rows)
            edited = st.data_editor(
                display_df,
                use_container_width=True,
                hide_index=True,
                column_config={
                    "选择": st.column_config.CheckboxColumn("选择"),
                    "课程名称": st.column_config.TextColumn("课程名称", alignment="center"),
                    "教材名称": st.column_config.TextColumn("教材名称", alignment="center"),
                    "书号": st.column_config.TextColumn("书号", alignment="center"),
                    "主编": st.column_config.TextColumn("主编", alignment="center"),
                    "出版社": st.column_config.TextColumn("出版社", alignment="center"),
                    "出版日期": st.column_config.TextColumn("出版日期", alignment="center"),
                    "单价(元)": st.column_config.NumberColumn("单价(元)", format="¥%.2f", alignment="center"),
                    "折扣率": st.column_config.NumberColumn("折扣率", format="%.2f", min_value=0.0, max_value=1.0, help="如 0.76=76折，1.0=原价", alignment="center"),
                    "实洋(元)": st.column_config.NumberColumn("实洋(元)", format="¥%.2f", help="学生结算价格，留空则=单价×折扣率", alignment="center"),
                    "教材类型": st.column_config.TextColumn("教材类型", help="如：国规教材、公共基础课、专业课、实验实训等", alignment="center"),
                    "引用次数": st.column_config.NumberColumn("引用次数", disabled=True, alignment="center"),
                },
                disabled=["引用次数"],
                key=f"master_editor_{mp}"
            )

            # ═══ 布局：上行(全选+分页) + 下行(红删蓝存) ═══
            # 预检测是否有修改
            has_changes = False
            for i, row in edited.iterrows():
                eid = real_ids[i]
                o = display_df.iloc[i]
                if (str(row["教材名称"]) != str(o["教材名称"]) or
                    str(row["书号"]) != str(o["书号"]) or
                    str(row["出版社"]) != str(o["出版社"]) or
                    str(row["主编"]) != str(o["主编"]) or
                    str(row["出版日期"]) != str(o["出版日期"]) or
                    float(row["单价(元)"]) != float(o["单价(元)"]) or
                    str(row["课程名称"]) != str(o["课程名称"]) or
                    abs(float(row["折扣率"]) - float(o["折扣率"])) > 1e-9 or
                    abs((float(row.get("实洋(元)", 0) or 0)) - (float(o.get("实洋(元)", 0) or 0))) > 1e-9 or
                    str(row.get("教材类型", "") or "") != str(o.get("教材类型", "") or "")):
                    has_changes = True
                    break

            # 处理按钮触发的操作（统一处理所有 pending_action）
            pending = st.session_state.get("pending_action", "")
            if pending == "master_save":
                if has_changes:
                    changes = 0
                    for i, row in edited.iterrows():
                        eid = real_ids[i]
                        o = display_df.iloc[i]
                        new_name = str(row["教材名称"]).strip()
                        if not new_name: continue
                        # 实洋价格变化检测（NaN 视为空，不参与比较）
                        ap_new = row.get("实洋(元)")
                        ap_old = o.get("实洋(元)")
                        ap_diff = False
                        ap_new_na = pd.isna(ap_new) if ap_new is not None else True
                        ap_old_na = pd.isna(ap_old) if ap_old is not None else True
                        if ap_new_na != ap_old_na:
                            ap_diff = True
                        elif not ap_new_na:
                            try:
                                if abs(float(ap_new) - float(ap_old)) > 1e-9:
                                    ap_diff = True
                            except (ValueError, TypeError):
                                ap_diff = True

                        if (str(row["教材名称"]) != str(o["教材名称"]) or
                            str(row["书号"]) != str(o["书号"]) or
                            str(row["出版社"]) != str(o["出版社"]) or
                            str(row["主编"]) != str(o["主编"]) or
                            str(row["出版日期"]) != str(o["出版日期"]) or
                            abs(float(row["单价(元)"]) - float(o["单价(元)"])) > 1e-9 or
                            str(row["课程名称"]) != str(o["课程名称"]) or
                            abs(float(row["折扣率"]) - float(o["折扣率"])) > 1e-9 or
                            ap_diff or
                            str(row.get("教材类型", "") or "") != str(o.get("教材类型", "") or "")):
                            # actual_price: 只有用户手动改实洋列时才写入，否则 NULL 让折扣率生效
                            ap_val = row.get("实洋(元)")
                            if ap_diff:
                                ap_sql = None if (ap_val is None or (isinstance(ap_val, float) and math.isnan(ap_val))) else float(ap_val)
                            else:
                                ap_sql = None  # 用户只改了折扣率，不覆写实洋
                            execute_sql(
                                "UPDATE textbooks_master SET name=%s, isbn=%s, publisher=%s, editor=%s, publication_date=%s, price=%s, course_name=%s, discount_rate=%s, actual_price=%s, textbook_type=%s WHERE id=%s",
                                (new_name, str(row["书号"]).strip(), str(row["出版社"]).strip(),
                                 str(row["主编"]).strip(), str(row["出版日期"]).strip(),
                                 float(row["单价(元)"]), str(row["课程名称"]).strip(),
                                 float(row["折扣率"]),
                                 ap_sql,
                                 str(row.get("教材类型", "") or "").strip() or None,
                                 eid))
                            changes += 1
                    if changes > 0:
                        st.toast(f"✅ 已保存 {changes} 条修改", icon="✅")
                st.session_state.pending_action = ""
                st.rerun()
            elif pending == "master_del":
                sl_checked = (edited["选择"].sum() if "选择" in edited.columns else 0)
                select_all = st.session_state.get(f"master_selall_{mp}", False)
                if select_all:
                    del_targets = page_df
                elif sl_checked > 0:
                    del_targets = page_df.iloc[edited[edited["选择"] == True].index]
                else:
                    del_targets = pd.DataFrame()
                del_ids = []
                del_names = []
                blocked = []
                for _, drow in del_targets.iterrows():
                    eid = int(drow["id"])
                    if usage_map.get(eid, 0) == 0:
                        del_ids.append(eid)
                        del_names.append(str(drow["name"]))
                    else:
                        blocked.append(f"「{drow['name']}」(被{usage_map.get(eid,0)}条引用)")
                if blocked:
                    for b in blocked:
                        st.warning(f"{b}，无法删除")
                if del_ids:
                    placeholders = ",".join(["%s"] * len(del_ids))
                    execute_sql(f"DELETE FROM textbooks_master WHERE id IN ({placeholders})", tuple(del_ids))
                    st.toast(f"✅ 已删除 {len(del_ids)} 本", icon="✅")
                st.session_state.pending_action = ""
                st.rerun()
            elif pending == "batch_dr":
                sel_ids = [real_ids[i] for i in range(len(edited)) if edited.iloc[i]["选择"]]
                if sel_ids:
                    dr_val = st.session_state.get("batch_dr_widget", 0.76)
                    placeholders = ",".join(["%s"] * len(sel_ids))
                    execute_sql(f"UPDATE textbooks_master SET discount_rate=%s WHERE id IN ({placeholders})",
                                (dr_val,) + tuple(sel_ids))
                    st.toast(f"✅ 已批量修改 {len(sel_ids)} 本教材折扣率为 {dr_val:.2f}", icon="✅")
                st.session_state.pending_action = ""
                st.rerun()

            # ── 上行：全选 + 分页 ──
            r1_sel, r1_info, r1_ps, r1_prev, r1_num, r1_next = st.columns([1.5, 2, 1, 0.7, 0.7, 0.7])
            with r1_sel:
                select_all = st.checkbox("全选本页", key=f"master_selall_{mp}",
                    help="勾选后删除操作将应用于本页全部记录")
            with r1_info:
                st.caption(f"共 **{len(master_df)}** 本")
            with r1_ps:
                st.selectbox("每页", [50, 100, 200], key="master_ps",
                             on_change=lambda: st.session_state.update({"master_page": 1}),
                             label_visibility="collapsed")
            with r1_prev:
                if st.button("◀", key="master_prev", disabled=(mp <= 1), use_container_width=True):
                    st.session_state["master_page"] = max(1, mp - 1); st.rerun()
            with r1_num:
                st.markdown(f"<div style='text-align:center;padding-top:5px;font-weight:500'>{mp}/{total_pages}</div>", unsafe_allow_html=True)
            with r1_next:
                if st.button("▶", key="master_next", disabled=(mp >= total_pages), use_container_width=True):
                    st.session_state["master_page"] = min(total_pages, mp + 1); st.rerun()

            # ── 下行：删除(红) + 保存(蓝) ──
            sl_checked = (edited["选择"].sum() if not select_all and "选择" in edited.columns else 0)
            del_count = len(page_df) if select_all else int(sl_checked)
            del_disabled = "disabled" if del_count == 0 else ""
            save_disabled = "disabled" if not has_changes else ""
            del_btn_label = f"🗑️ 删除（{del_count}本）"
            # 统计确切的修改数量
            chg_count = 0
            if has_changes:
                for i, row in edited.iterrows():
                    eid = real_ids[i]
                    o = display_df.iloc[i]
                    if (str(row["教材名称"]) != str(o["教材名称"]) or
                        str(row["书号"]) != str(o["书号"]) or
                        str(row["出版社"]) != str(o["出版社"]) or
                        str(row["主编"]) != str(o["主编"]) or
                        str(row["出版日期"]) != str(o["出版日期"]) or
                        float(row["单价(元)"]) != float(o["单价(元)"]) or
                        str(row["课程名称"]) != str(o["课程名称"]) or
                        abs(float(row["折扣率"]) - float(o["折扣率"])) > 1e-9 or
                        abs((float(row.get("实洋(元)", 0) or 0)) - (float(o.get("实洋(元)", 0) or 0))) > 1e-9 or
                        str(row.get("教材类型", "") or "") != str(o.get("教材类型", "") or "")):
                        chg_count += 1
            save_btn_label = f"💾 保存修改（{chg_count}处）" if has_changes else "💾 保存修改"

            r2_del, r2_save = st.columns([1, 1])
            with r2_del:
                st.button(del_btn_label, key=f"master_del_btn_{mp}", type="secondary",
                          disabled=(del_count == 0), use_container_width=True,
                          on_click=lambda: st.session_state.update({"pending_action": "master_del"}))
            with r2_save:
                st.button(save_btn_label, key=f"master_save_btn_{mp}", type="primary",
                          disabled=(not has_changes), use_container_width=True,
                          on_click=lambda: st.session_state.update({"pending_action": "master_save"}))
        else:
            st.info("📭 教材库为空，请通过「新增教材」或「导入 Excel」添加")

    # ═══════════════ Tab2: 新增教材 ═══════════════
    with tab2:
        st.markdown("#### ➕ 手动新增教材")

        with st.form("master_add_form", clear_on_submit=True):
            ac1, ac2 = st.columns(2)
            with ac1:
                new_name = st.text_input("教材名称 *", placeholder="必填，如：高等数学（第七版）", key="new_master_name")
                new_isbn = st.text_input("书号(ISBN)", placeholder="978-7-04-059109-9", key="new_master_isbn")
                new_publisher = st.text_input("出版社", placeholder="高等教育出版社", key="new_master_publisher")
                new_editor = st.text_input("主编", placeholder="同济大学数学科学学院", key="new_master_editor")
                new_course = st.text_input("课程名称", placeholder="如：高等数学A", key="new_master_course")
            with ac2:
                new_price = st.number_input("单价(元)", min_value=0.0, value=0.0, step=0.01, format="%.2f", key="new_master_price")
                new_dr = st.number_input("折扣率", min_value=0.0, max_value=1.0,
                    value=st.session_state.get("master_default_dr", 0.76),
                    step=0.01, format="%.2f", key="new_master_dr",
                    help="如 0.76=76折，1.0=原价")
                new_ap = st.number_input("实洋(元)", min_value=0.0, value=None, step=0.01, format="%.2f",
                    key="new_master_ap", placeholder="留空=自动计算",
                    help="学生结算价格，留空则按 单价×折扣率 自动计算")
                new_pub_date = st.text_input("出版日期", placeholder="如：2023年08月", key="new_master_pubdate")
                new_type = st.text_input("教材类型", placeholder="如：国规教材、专业课、公共基础课", key="new_master_type")

            submitted = st.form_submit_button("✅ 确认添加", type="primary", use_container_width=True)

        if submitted:
            if not new_name.strip():
                st.error("❌ 教材名称为必填项")
            else:
                # 查重
                exist_check = query_df("SELECT id FROM textbooks_master WHERE name=%s", (new_name.strip(),))
                if not exist_check.empty:
                    st.error(f"❌ 教材「{new_name.strip()}」已存在（ID:{int(exist_check.iloc[0]['id'])}），请使用其他名称或前往列表编辑")
                else:
                    execute_sql(
                        "INSERT INTO textbooks_master (name, isbn, publisher, editor, price, publication_date, course_name, discount_rate, actual_price, textbook_type) "
                        "VALUES (%s,%s,%s,%s,%s,%s,%s,%s,%s,%s)",
                        (new_name.strip(),
                         new_isbn.strip() or None,
                         new_publisher.strip() or None,
                         new_editor.strip() or None,
                         new_price,
                         new_pub_date.strip() or None,
                         new_course.strip() or None,
                         new_dr,
                         new_ap if new_ap is not None and new_ap > 0 else None,
                         new_type.strip() or None))
                    st.toast(f"✅ 教材「{new_name.strip()}」已添加", icon="✅")
                    st.rerun()

    # ═══════════════ Tab3: 导入 Excel ═══════════════
    with tab3:
        st.markdown("#### 📥 从 Excel 导入教材")

        # 下载模板
        with st.expander("📄 模板下载与说明", expanded=False):
            template_cols = ["课程名称", "教材名称", "书号(ISBN)", "主编", "出版社", "出版日期", "单价(元)", "折扣率", "实洋(元)", "教材类型"]
            template_df = make_template_df(template_cols)
            template_bytes = excel_export(template_df, "教材表导入模板")
            st.download_button(
                "📄 下载导入模板",
                data=template_bytes,
                file_name="教材表导入模板.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                use_container_width=True,
                type="secondary"
            )
            st.caption("Excel 表头建议包含：课程名称、教材名称、书号(ISBN)、主编、出版社、出版日期、单价(元)、折扣率、实洋(元)、教材类型")

        uploaded = st.file_uploader("选择 Excel 文件", type=["xlsx", "xls"], key="master_import_upload")
        if uploaded:
            try:
                raw_df = read_excel_upload(uploaded)
                st.info(f"📄 检测到 {len(raw_df)} 行，列名：{', '.join(raw_df.columns[:10])}")

                # 列映射
                col_map = {}
                for col in raw_df.columns:
                    cl = col.strip()
                    cl_lower = cl.lower().replace(" ", "").replace(" ", "")
                    # 先匹配课程名称（必须放在"名称"之前，否则"课程名称"会被误映射到name）
                    if "课程名称" in cl or "课程" in cl or "科目" in cl or "course" in cl_lower:
                        col_map["course_name"] = cl
                    elif "教材名称" in cl or "书名" in cl or "name" in cl_lower or ("名称" in cl and "课程" not in cl and "折扣" not in cl):
                        col_map["name"] = cl
                    elif "isbn" in cl_lower or "书号" in cl:
                        col_map["isbn"] = cl
                    elif "出版社" in cl or "press" in cl_lower or "publisher" in cl_lower:
                        col_map["publisher"] = cl
                    elif "主编" in cl or "作者" in cl or "编者" in cl or "editor" in cl_lower or "author" in cl_lower:
                        col_map["editor"] = cl
                    elif "单价" in cl or "价格" in cl or "定价" in cl or "price" in cl_lower:
                        col_map["price"] = cl
                    elif "出版日期" in cl or "出版时间" in cl or "出版年" in cl or "pub_date" in cl_lower or "publication_date" in cl_lower or "publish_date" in cl_lower:
                        col_map["publication_date"] = cl
                    elif "实洋" in cl or "实洋" in cl or "actual_price" in cl_lower:
                        col_map["actual_price"] = cl
                    elif "教材类型" in cl or "类型" in cl or "type" in cl_lower:
                        col_map["textbook_type"] = cl
                    elif "国规" in cl or "规划教材" in cl or "national" in cl_lower:
                        col_map["textbook_type"] = cl  # 旧字段也映射到 textbook_type
                    elif "折扣" in cl or "discount" in cl_lower:
                        col_map["discount_rate"] = cl

                if "name" not in col_map:
                    st.error("❌ 缺少必要列：教材名称")
                    st.json(col_map)
                else:
                    # ── 列映射（做完再预览，出版日期也会格式化） ──
                    mapped = raw_df.rename(columns={v: k for k, v in col_map.items()})
                    wanted = [c for c in ["name", "isbn", "publisher", "editor", "price", "publication_date", "course_name", "discount_rate", "actual_price", "textbook_type"] if c in mapped.columns]
                    mapped = mapped[wanted].where(pd.notnull(mapped), None)

                    # ── 出版日期统一转 "YYYY年MM月"（Excel 序列号 / datetime 均支持） ──
                    if "publication_date" in mapped.columns:
                        def _fmt_pubdate(v):
                            if v is None or (isinstance(v, float) and pd.isna(v)):
                                return ""
                            if isinstance(v, (int, float)) and v > 1000:
                                try:
                                    d = datetime(1899, 12, 30) + timedelta(days=int(v))
                                    return f"{d.year}年{d.month:02d}月"
                                except:
                                    pass
                            if hasattr(v, "strftime"):
                                return f"{v.year}年{v.month:02d}月"
                            s = str(v).strip()
                            # 已经是 "年/月" 或 "年-月" 格式，统一为 "年 月"
                            m = re.match(r"(\d{4})[年/\-](\d{1,2})", s)
                            if m:
                                return f"{m.group(1)}年{int(m.group(2)):02d}月"
                            return s
                        mapped["publication_date"] = mapped["publication_date"].apply(_fmt_pubdate)

                    # ── 折扣率：Excel 有则用，无则用页面默认值 ──
                    default_dr = st.session_state.get("master_default_dr", 0.76)
                    if "discount_rate" not in mapped.columns:
                        mapped["discount_rate"] = default_dr
                    else:
                        mapped["discount_rate"] = mapped["discount_rate"].apply(
                            lambda v: float(v) if v is not None and not (isinstance(v, float) and pd.isna(v)) else default_dr)

                    total_rows = len(mapped)

                    # 预览（日期已格式化，不再是序列号）
                    preview = mapped.head(5)
                    styled_dataframe(preview, hide_ids=True)

                    # 收集已存在的教材名称（一次性查，不在循环里逐条查）
                    book_names = [safe_str(r.get("name", "")) for _, r in mapped.iterrows() if safe_str(r.get("name", ""))]
                    unique_names = list(set(book_names))
                    if unique_names:
                        placeholders = ",".join(["%s"] * len(unique_names))
                        exist_df = query_df(f"SELECT name, id FROM textbooks_master WHERE name IN ({placeholders})", tuple(unique_names))
                        exist_set = set(exist_df["name"].tolist()) if not exist_df.empty else set()
                        exist_detail = {}
                        for _, er in exist_df.iterrows():
                            exist_detail[er["name"]] = int(er["id"])
                    else:
                        exist_set = set()
                        exist_detail = {}

                    # 预扫描统计（也处理 Excel 内部重复）
                    seen_names = set()
                    pre_new = 0
                    pre_update = 0
                    pre_skip = 0
                    pre_dup = 0  # Excel 内部重名
                    pre_update_names = []
                    pre_dup_names = []
                    for _, row in mapped.iterrows():
                        bn = safe_str(row.get("name", ""))
                        if not bn:
                            pre_skip += 1
                            continue
                        if bn in seen_names:
                            pre_dup += 1
                            pre_dup_names.append(bn)
                            continue
                        seen_names.add(bn)
                        if bn in exist_set:
                            pre_update += 1
                            pre_update_names.append(bn)
                        else:
                            pre_new += 1

                    # 预览面板
                    st.markdown("---")
                    st.markdown("#### 📊 导入预览")
                    c_a, c_b, c_c, c_d = st.columns(4)
                    c_a.metric("✅ 将新增", pre_new)
                    c_b.metric("🔄 将更新", pre_update)
                    c_c.metric("⏭️ 跳过（空名）", pre_skip)
                    c_d.metric("⚠️ 重名（去重）", pre_dup)
                    if pre_update > 0 and pre_update_names:
                        with st.expander(f"📋 查看将被更新的 {pre_update} 本教材（数据库中已存在）"):
                            for nm in sorted(pre_update_names):
                                st.caption(f"• {nm}")
                    if pre_dup > 0 and pre_dup_names:
                        with st.expander(f"⚠️ Excel 内部重复 {pre_dup} 行（仅保留首次出现）"):
                            for nm in sorted(set(pre_dup_names)):
                                st.caption(f"• {nm}")

                    if st.button("✅ 确认导入", use_container_width=True, type="primary", key="master_import_btn"):
                        # 去重：名称相同则覆盖（UPDATE），否则 INSERT
                        # Excel 内部也去重（仅保留首次出现的书名）
                        success_count = 0
                        update_count = 0
                        skip_count = 0
                        dup_skip = 0
                        errors = []
                        imported_names = set()
                        progress = st.progress(0, text="正在导入...")
                        for i, (_, row) in enumerate(mapped.iterrows()):
                            try:
                                book_name = safe_str(row.get("name", ""))
                                if not book_name:
                                    skip_count += 1
                                    continue
                                # Excel 内部去重：同名只保留首次出现
                                if book_name in imported_names:
                                    dup_skip += 1
                                    continue
                                imported_names.add(book_name)
                                # 查重（用预扫描结果加速）
                                if book_name in exist_detail:
                                    eid = exist_detail[book_name]
                                    execute_sql(
                                        "UPDATE textbooks_master SET isbn=%s, publisher=%s, editor=%s, price=%s, publication_date=%s, course_name=%s, discount_rate=%s, actual_price=%s, textbook_type=%s WHERE id=%s",
                                        (safe_str(row.get("isbn")), safe_str(row.get("publisher")),
                                         safe_str(row.get("editor")), safe_float(row.get("price"), 0),
                                         safe_str(row.get("publication_date")),
                                         safe_str(row.get("course_name")),
                                         float(row["discount_rate"]) if "discount_rate" in row.index else default_dr,
                                         safe_float(row.get("actual_price")) if row.get("actual_price") is not None and not (isinstance(row.get("actual_price"), float) and pd.isna(row.get("actual_price"))) else None,
                                         safe_str(row.get("textbook_type")) or None,
                                         eid))
                                    update_count += 1
                                else:
                                    execute_sql(
                                        "INSERT INTO textbooks_master (name, isbn, publisher, editor, price, publication_date, course_name, discount_rate, actual_price, textbook_type) VALUES (%s,%s,%s,%s,%s,%s,%s,%s,%s,%s)",
                                        (book_name, safe_str(row.get("isbn")), safe_str(row.get("publisher")),
                                         safe_str(row.get("editor")), safe_float(row.get("price"), 0),
                                         safe_str(row.get("publication_date")),
                                         safe_str(row.get("course_name")),
                                         float(row["discount_rate"]) if "discount_rate" in row.index else default_dr,
                                         safe_float(row.get("actual_price")) if row.get("actual_price") is not None and not (isinstance(row.get("actual_price"), float) and pd.isna(row.get("actual_price"))) else None,
                                         safe_str(row.get("textbook_type")) or None))
                                    success_count += 1
                            except Exception as e:
                                errors.append(f"{row.get('name','?')}: {str(e)[:80]}")
                            progress.progress((i + 1) / total_rows, text=f"已处理 {i+1}/{total_rows}")

                        # 写日志
                        write_import_log(
                            module="教材表管理",
                            filename=uploaded.name,
                            total=total_rows,
                            success=success_count,
                            errors=errors
                        )

                        # 展示结果
                        mc_ok, mc_upd, mc_err, mc_dp = st.columns(4)
                        mc_ok.metric("✅ 新增", success_count)
                        mc_upd.metric("🔄 更新", update_count)
                        mc_err.metric("⚠️ 失败", len(errors))
                        mc_dp.metric("🔁 去重跳过", dup_skip)
                        if errors:
                            with st.expander(f"查看 {len(errors)} 条失败详情"):
                                for err in errors:
                                    st.caption(f"• {err}")
                        if skip_count > 0:
                            st.caption(f"⏭️ 跳过 {skip_count} 行（教材名称为空）")
            except Exception as e:
                st.error(f"❌ 读取失败：{e}")
