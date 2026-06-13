def textbook_management():
    """教材征订表管理（双表架构：textbooks_master + textbook_orders）"""
    import sqlite3
    show_header("📖 教材征订表", "教材库 + 征订明细")

    semesters = query_df("SELECT id, name FROM semesters ORDER BY id DESC")
    if semesters.empty:
        st.warning("⚠️ 请先在「学期管理」中添加学期")
        return

    semester_options = ["全部"] + [f"{r['id']}|{r['name']}" for _, r in semesters.iterrows()]
    master_books = query_df("SELECT id, name, isbn, publisher, editor, price, course_name FROM textbooks_master ORDER BY name")
    master_options = [(0, "➕ 新增教材...")] + [(r["id"], r["name"]) for _, r in master_books.iterrows()]

    tab1, tab2, tab3, tab4 = st.tabs(["📋 征订列表", "➕ 新增/编辑征订", "📥 导入 Excel", "✏️ 批量编辑"])

    with tab1:
        col1, col2, col3, col4 = st.columns(4)
        with col1: f_sem = st.selectbox("学期", semester_options, key="t_sem")
        with col2: f_tcollege = st.selectbox("学院", ["全部"] + get_filtered_colleges(), key="t_college")
        with col3: f_tmajor = st.selectbox("专业", ["全部"] + get_filtered_majors(), key="t_major")
        with col4: f_tclass = st.selectbox("班级", ["全部"] + get_filtered_class_names(), key="t_class")
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
            m1, m2 = st.columns([0.6, 0.4])
            with m1: st.caption(f"共 **{len(df)}** 条征订记录")
            with m2: st.metric("💰 总计金额", f"¥{total_amount:,.2f}")
            cols = ["id","semester_name","college","major","class_name","grade",
                    "name","course_name","isbn","publisher","editor","price","quantity","remark"]
            display_df = df[cols]
            rename_map = {"semester_name":"学期","college":"学院","major":"专业","class_name":"班级","grade":"年级",
                "name":"教材名称","course_name":"课程","isbn":"书号","publisher":"出版社","editor":"主编",
                "price":"单价(元)","quantity":"征订数量","remark":"备注"}
            st.dataframe(display_df.rename(columns=rename_map).style.set_properties(**{"text-align": "center"}), use_container_width=True, hide_index=True,
                column_config={"单价(元)": st.column_config.NumberColumn("单价(元)", format="¥%.2f")})
        else:
            st.info("暂无征订数据")

    with tab2:
        st.markdown("#### ➕ 新增征订 / 管理教材库")
        edit_oid = st.number_input("编辑征订ID（留空为新增）", min_value=0, value=0, step=1, key="edit_oid")
        defaults = {}
        if edit_oid > 0:
            row = query_df("SELECT o.*, tm.name as book_name FROM textbook_orders o JOIN textbooks_master tm ON o.textbook_id=tm.id WHERE o.id=%s", (edit_oid,))
            if not row.empty: defaults = row.iloc[0].to_dict()

        with st.form("textbook_order_form"):
            st.caption("🔹 第一步：选择或新增教材")
            selected_master = st.selectbox("选择已有教材", master_options,
                format_func=lambda x: x[1],
                index=0 if edit_oid == 0 else next((i for i, o in enumerate(master_options) if o[0] == defaults.get("textbook_id", 0)), 0),
                key="sel_master")
            is_new_book = (selected_master[0] == 0)

            st.caption("🔹 第二步：教材信息（新增时必填）")
            cols_book = st.columns(3)
            with cols_book[0]:
                bk_name = st.text_input("教材名称*", value="" if is_new_book else selected_master[1], placeholder="新增时必填", key="bk_name")
                bk_isbn = st.text_input("书号(ISBN)", value=next((r.get("isbn","") for _,r in master_books.iterrows() if r["id"]==selected_master[0]), "") if not is_new_book else "", key="bk_isbn")
            with cols_book[1]:
                bk_publisher = st.text_input("出版社", value=next((r.get("publisher","") for _,r in master_books.iterrows() if r["id"]==selected_master[0]), "") if not is_new_book else "", key="bk_pub")
                bk_editor = st.text_input("主编", value=next((r.get("editor","") for _,r in master_books.iterrows() if r["id"]==selected_master[0]), "") if not is_new_book else "", key="bk_ed")
            with cols_book[2]:
                bk_price = st.number_input("单价(元)*", min_value=0.0, value=float(next((r["price"] for _,r in master_books.iterrows() if r["id"]==selected_master[0]), 0) if not is_new_book else 0), step=0.01, format="%.2f", key="bk_price")
                bk_course = st.text_input("课程", value=next((r.get("course_name","") for _,r in master_books.iterrows() if r["id"]==selected_master[0]), "") if not is_new_book else "", key="bk_course")

            st.caption("🔹 第三步：选择征订范围")
            sem_id = st.selectbox("学期*", [(r["id"], r["name"]) for _, r in semesters.iterrows()],
                format_func=lambda x: x[1], index=next((i for i, s in enumerate([(r["id"], r["name"]) for _, r in semesters.iterrows()]) if s[0] == defaults.get("semester_id")), 0))
            sel_grades = st.multiselect("年级（可多选）", options=get_filtered_grades(), default=[defaults.get("grade")] if defaults.get("grade") else [])
            sel_colleges = st.multiselect("学院（可多选）", options=get_filtered_colleges(), default=[defaults.get("college")] if defaults.get("college") else [])
            sel_majors = st.multiselect("专业（可多选）", options=get_filtered_majors(), default=[defaults.get("major")] if defaults.get("major") else [])
            sel_classes = st.multiselect("班级（可多选）", options=get_filtered_class_names(), default=[defaults.get("class_name")] if defaults.get("class_name") else [])
            o_qty = st.number_input("征订数量", min_value=0, value=int(defaults.get("quantity", 0)), step=1)
            o_remark = st.text_input("备注", value=defaults.get("remark", ""))

            combos = max(len(sel_grades) or 1, len(sel_colleges) or 1, len(sel_majors) or 1, len(sel_classes) or 1)
            if edit_oid == 0 and combos > 1:
                st.info(f"📌 将新增 **{combos}** 条征订记录")

            col_btn1, col_btn2 = st.columns(2)
            with col_btn1: submitted = st.form_submit_button("💾 保存", use_container_width=True)
            with col_btn2: delete_btn = st.form_submit_button("🗑️ 删除", use_container_width=True) if edit_oid > 0 else False

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
                    mid = selected_master[0]

                if edit_oid > 0:
                    execute_sql("UPDATE textbook_orders SET semester_id=%s,grade=%s,college=%s,major=%s,class_name=%s,quantity=%s,remark=%s WHERE id=%s",
                        (sem_id[0], (sel_grades or [None])[0], (sel_colleges or [None])[0], (sel_majors or [None])[0], (sel_classes or [None])[0], o_qty, o_remark, edit_oid))
                    st.success("✅ 征订已更新")
                else:
                    grades, colleges, majors, classes = sel_grades or [None], sel_colleges or [None], sel_majors or [None], sel_classes or [None]
                    cnt = 0
                    for g in grades:
                        for c in colleges:
                            for m in majors:
                                for cl in classes:
                                    execute_sql("INSERT INTO textbook_orders (semester_id,textbook_id,grade,college,major,class_name,quantity,remark) VALUES (%s,%s,%s,%s,%s,%s,%s,%s)",
                                        (sem_id[0], mid, g, c, m, cl, o_qty, o_remark))
                                    cnt += 1
                    st.success(f"✅ 已添加 {cnt} 条征订记录")
                st.rerun()
            if delete_btn:
                execute_sql("DELETE FROM textbook_orders WHERE id=%s", (edit_oid,))
                st.success("✅ 已删除")
                st.rerun()

    with tab3:
        st.markdown("#### 📥 从 Excel 导入教材征订")
        template_cols = ["教材名称", "学期", "年级", "学院", "专业", "班级", "书号(ISBN)", "出版社", "主编", "单价(元)", "征订数量", "备注"]
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
                    elif "专业" in cl: col_map["major"] = col
                    elif "班级" in cl: col_map["class_name"] = col
                    elif "单价" in cl: col_map["price"] = col
                    elif "数量" in cl: col_map["quantity"] = col
                    elif "书号" in cl: col_map["isbn"] = col
                    elif "出版社" in cl: col_map["publisher"] = col
                    elif "主编" in cl: col_map["editor"] = col
                    elif "备注" in cl: col_map["remark"] = col
                if "name" not in col_map or "semester_name" not in col_map:
                    st.error("❌ 缺少必要列：教材名称、学期")
                else:
                    if st.button("✅ 确认导入", use_container_width=True, type="primary", key="tb_import3"):
                        mapped = raw_df.rename(columns={v: k for k, v in col_map.items()})
                        sem_map = {r["name"]: r["id"] for _, r in semesters.iterrows()}
                        success, errors = 0, []
                        for _, row in mapped.iterrows():
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
                                execute_sql("INSERT INTO textbook_orders (semester_id,textbook_id,grade,college,major,class_name,quantity,remark) VALUES (%s,%s,%s,%s,%s,%s,%s,%s)",
                                    (sid, tid, str(row.get("grade","") or ""), str(row.get("college","") or ""), str(row.get("major","") or ""), str(row.get("class_name","") or ""), safe_int(row.get("quantity", 0)), str(row.get("remark","") or "")))
                                success += 1
                            except Exception as e: errors.append(f"{nm}: {str(e)[:40]}")
                        st.success(f"✅ 导入 {success} 条")
                        if errors:
                            st.warning(f"⚠️ {len(errors)} 条失败")
                            with st.expander("查看详情"):
                                for e in errors[:10]: st.caption(f"• {e}")
                        st.rerun()
            except Exception as e:
                st.error(f"❌ 读取失败：{e}")

    with tab4:
        st.markdown("#### ✏️ 批量编辑征订")
        be_sems = st.multiselect("学期", options=semester_options[1:],
            format_func=lambda x: x.split("|",1)[1], key="be_sem3")
        be_classes = st.multiselect("班级", options=get_filtered_class_names(), key="be_class3")
        be_grades = st.multiselect("年级", options=get_filtered_grades(), key="be_grade3")
        bf = st.selectbox("要批量设置的字段", ["单价", "征订数量", "班级", "年级", "专业", "备注"])
        bv = st.text_input("设置为", placeholder="新值...", key="bt_val")
        if bf and bv:
            fm = {"单价":"price","征订数量":"quantity","班级":"class_name","年级":"grade","专业":"major","备注":"remark"}
            db_f = fm[bf]
            wc, pp = [], []
            pp.append(safe_float(bv) if bf=="单价" else (safe_int(bv) if bf=="征订数量" else bv))
            for vals, col, is_sem in [(be_sems,"o.semester_id",True),(be_classes,"o.class_name",False),(be_grades,"o.grade",False)]:
                if vals:
                    ph = ", ".join(["%s"]*len(vals))
                    wc.append(f"{col} IN ({ph})")
                    for v in vals: pp.append(int(v.split("|")[0]) if is_sem else v)
            ws = " AND ".join(wc) if wc else "1=1"
            pv = query_df(f"SELECT o.id,tm.name,o.class_name,tm.price,o.quantity FROM textbook_orders o JOIN textbooks_master tm ON o.textbook_id=tm.id WHERE {ws}",
                tuple(pp[1:]) if len(pp)>1 else None)
            st.warning(f"⚠️ 影响 **{len(pv)}** 条")
            if st.button("⚠️ 确认批量更新", use_container_width=True, type="primary", key="bt_conf3"):
                if bf == "单价":
                    execute_sql(f"UPDATE textbooks_master SET price=%s WHERE id IN (SELECT DISTINCT textbook_id FROM textbook_orders WHERE {ws})", tuple(pp))
                else:
                    execute_sql(f"UPDATE textbook_orders o SET {db_f}=%s WHERE {ws}", tuple(pp))
                st.success(f"✅ 已批量更新 {len(pv)} 条")
                st.rerun()
