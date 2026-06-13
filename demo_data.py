"""
演示数据自动初始化模块
======================
当数据库为空时（如 Streamlit Cloud 首次部署），
自动创建演示数据，使系统可直接运行演示。
"""
from database import query_df, execute_sql


def is_database_empty() -> bool:
    """检测数据库是否为空（无学期、无学生）"""
    try:
        sem_count = query_df("SELECT COUNT(*) as cnt FROM semesters")
        stu_count = query_df("SELECT COUNT(*) as cnt FROM students")
        return (sem_count.empty or int(sem_count.iloc[0]["cnt"]) == 0) and \
               (stu_count.empty or int(stu_count.iloc[0]["cnt"]) == 0)
    except Exception:
        return True


def init_demo_data() -> dict:
    """
    自动创建演示数据。

    Returns:
        dict: 创建的各记录数量统计
    """
    if not is_database_empty():
        return {"skipped": True, "reason": "数据库已有数据"}

    stats = {"semesters": 0, "students": 0, "textbooks": 0,
             "master_books": 0, "orders": 0, "distributions": 0}

    # ── 1. 学期 ──
    semesters_data = [
        ("2025-2026 第一学期", "2025-2026", "第一学期"),
        ("2025-2026 第二学期", "2025-2026", "第二学期"),
    ]
    sem_ids = {}
    for name, ay, sn in semesters_data:
        try:
            sid = execute_sql(
                "INSERT INTO semesters (name, academic_year, semester_name) VALUES (%s, %s, %s)",
                (name, ay, sn)
            )
            sem_ids[name] = sid
            stats["semesters"] += 1
        except Exception:
            df = query_df("SELECT id FROM semesters WHERE name = %s", (name,))
            if not df.empty:
                sem_ids[name] = int(df.iloc[0]["id"])

    # ── 2. 学生（3个班级，各8人） ──
    students_data = [
        # 软件工程1班
        ("320101200501010001", "2025001", "张明",   "2025级", "计算机学院", "软件工程", "软件工程1班"),
        ("320101200501010002", "2025002", "李华",   "2025级", "计算机学院", "软件工程", "软件工程1班"),
        ("320101200501010003", "2025003", "王芳",   "2025级", "计算机学院", "软件工程", "软件工程1班"),
        ("320101200501010004", "2025004", "赵强",   "2025级", "计算机学院", "软件工程", "软件工程1班"),
        ("320101200501010005", "2025005", "刘洋",   "2025级", "计算机学院", "软件工程", "软件工程1班"),
        ("320101200501010006", "2025006", "陈静",   "2025级", "计算机学院", "软件工程", "软件工程1班"),
        ("320101200501010007", "2025007", "杨磊",   "2025级", "计算机学院", "软件工程", "软件工程1班"),
        ("320101200501010008", "2025008", "黄丽",   "2025级", "计算机学院", "软件工程", "软件工程1班"),
        # 计算机科学1班
        ("320101200502010001", "2025009", "周杰",   "2025级", "计算机学院", "计算机科学", "计算机科学1班"),
        ("320101200502010002", "2025010", "吴敏",   "2025级", "计算机学院", "计算机科学", "计算机科学1班"),
        ("320101200502010003", "2025011", "郑伟",   "2025级", "计算机学院", "计算机科学", "计算机科学1班"),
        ("320101200502010004", "2025012", "孙燕",   "2025级", "计算机学院", "计算机科学", "计算机科学1班"),
        ("320101200502010005", "2025013", "朱鹏",   "2025级", "计算机学院", "计算机科学", "计算机科学1班"),
        ("320101200502010006", "2025014", "何雪",   "2025级", "计算机学院", "计算机科学", "计算机科学1班"),
        ("320101200502010007", "2025015", "林峰",   "2025级", "计算机学院", "计算机科学", "计算机科学1班"),
        ("320101200502010008", "2025016", "徐婷",   "2025级", "计算机学院", "计算机科学", "计算机科学1班"),
        # 应用数学1班
        ("320101200503010001", "2025017", "马超",   "2025级", "数学学院", "应用数学", "应用数学1班"),
        ("320101200503010002", "2025018", "高颖",   "2025级", "数学学院", "应用数学", "应用数学1班"),
        ("320101200503010003", "2025019", "罗浩",   "2025级", "数学学院", "应用数学", "应用数学1班"),
        ("320101200503010004", "2025020", "梁欣",   "2025级", "数学学院", "应用数学", "应用数学1班"),
        ("320101200503010005", "2025021", "宋磊",   "2025级", "数学学院", "应用数学", "应用数学1班"),
        ("320101200503010006", "2025022", "谢琳",   "2025级", "数学学院", "应用数学", "应用数学1班"),
        ("320101200503010007", "2025023", "韩飞",   "2025级", "数学学院", "应用数学", "应用数学1班"),
        ("320101200503010008", "2025024", "唐梅",   "2025级", "数学学院", "应用数学", "应用数学1班"),
    ]
    for id_card, sid, name, grade, college, major, cls in students_data:
        try:
            execute_sql(
                """INSERT INTO students (id_card, student_id, name, grade, college, major, class_name)
                   VALUES (%s, %s, %s, %s, %s, %s, %s)""",
                (id_card, sid, name, grade, college, major, cls)
            )
            stats["students"] += 1
        except Exception:
            pass

    # ── 3. 教材主表 ──
    master_books = [
        ("Python程序设计",       "978-7-302-60001-1", "清华大学出版社", "张三丰", 59.00, "Python程序设计",       None, 0.85, None),
        ("数据结构与算法",       "978-7-302-60002-8", "清华大学出版社", "李思",   45.00, "数据结构与算法",       None, 0.85, None),
        ("计算机网络",           "978-7-121-60003-5", "电子工业出版社", "王武",   42.00, "计算机网络",           None, 0.80, None),
        ("操作系统原理",         "978-7-111-60004-2", "机械工业出版社", "赵柳",   55.00, "操作系统原理",         None, 0.85, None),
        ("高等数学（上）",      "978-7-04-60005-9",  "高等教育出版社", "陈七",   38.00, "高等数学",             None, 0.90, None),
        ("线性代数",             "978-7-04-60006-6",  "高等教育出版社", "刘八",   32.00, "线性代数",             None, 0.90, None),
        ("数据库系统概论",      "978-7-04-60007-0",  "高等教育出版社", "李思",   48.00, "数据库系统概论",      None, 0.85, None),
        ("计算机组成原理",      "978-7-302-60008-7", "清华大学出版社", "王武",   52.00, "计算机组成原理",      None, 0.85, None),
        ("高等数学（下）",      "978-7-04-60009-3",  "高等教育出版社", "陈七",   41.00, "高等数学（下）",      None, 0.90, None),
        ("概率论与数理统计",    "978-7-04-60010-3",  "高等教育出版社", "刘八",   35.00, "概率论与数理统计",    None, 0.90, None),
        ("软件工程导论",        "978-7-302-60011-7", "清华大学出版社", "周九",   50.00, "软件工程导论",        None, 0.85, None),
        ("Web前端开发",          "978-7-121-60012-4", "电子工业出版社", "吴十",   46.00, "Web前端开发",          None, 0.80, None),
        ("编译原理",             "978-7-111-60013-1", "机械工业出版社", "赵柳",   54.00, "编译原理",             None, 0.85, None),
        ("人工智能导论",        "978-7-302-60014-8", "清华大学出版社", "李思",   49.00, "人工智能导论",        None, 0.85, None),
        ("数学建模",             "978-7-04-60015-5",  "高等教育出版社", "陈七",   37.00, "数学建模",             None, 0.90, None),
        ("离散数学",             "978-7-302-60016-2", "清华大学出版社", "郑伟",   40.00, "离散数学",             None, 0.85, None),
    ]
    master_ids = {}
    for name, isbn, pub, editor, price, course, pub_date, disc, actual in master_books:
        try:
            mid = execute_sql(
                """INSERT INTO textbooks_master (name, isbn, publisher, editor, price, course_name,
                   publication_date, discount_rate, actual_price)
                   VALUES (%s, %s, %s, %s, %s, %s, %s, %s, %s)""",
                (name, isbn, pub, editor, price, course, pub_date, disc, actual)
            )
            master_ids[isbn] = mid
            stats["master_books"] += 1
        except Exception:
            df = query_df("SELECT id FROM textbooks_master WHERE isbn = %s", (isbn,))
            if not df.empty:
                master_ids[isbn] = int(df.iloc[0]["id"])

    # ── 4. 旧版 textbooks 表（兼容）+ 发放记录 ──
    sem1 = sem_ids.get("2025-2026 第一学期")
    sem2 = sem_ids.get("2025-2026 第二学期")

    textbooks_data = [
        # 第一学期
        (sem1, "2025级", "计算机学院", "软件工程",   "软件工程1班",   "Python程序设计",     "清华大学出版社", "张三丰", "978-7-302-60001-1", 59.00, 8,  ""),
        (sem1, "2025级", "计算机学院", "软件工程",   "软件工程1班",   "数据结构与算法",     "清华大学出版社", "李思",   "978-7-302-60002-8", 45.00, 8,  ""),
        (sem1, "2025级", "计算机学院", "软件工程",   "软件工程1班",   "离散数学",           "清华大学出版社", "郑伟",   "978-7-302-60016-2", 40.00, 8,  ""),
        (sem1, "2025级", "计算机学院", "计算机科学", "计算机科学1班", "计算机网络",         "电子工业出版社", "王武",   "978-7-121-60003-5", 42.00, 8,  ""),
        (sem1, "2025级", "计算机学院", "计算机科学", "计算机科学1班", "操作系统原理",       "机械工业出版社", "赵柳",   "978-7-111-60004-2", 55.00, 8,  ""),
        (sem1, "2025级", "数学学院",   "应用数学",   "应用数学1班",   "高等数学（上）",     "高等教育出版社", "陈七",   "978-7-04-60005-9", 38.00, 8,  ""),
        (sem1, "2025级", "数学学院",   "应用数学",   "应用数学1班",   "线性代数",           "高等教育出版社", "刘八",   "978-7-04-60006-6", 32.00, 8,  ""),
        # 第二学期
        (sem2, "2025级", "计算机学院", "软件工程",   "软件工程1班",   "数据库系统概论",     "高等教育出版社", "李思",   "978-7-04-60007-0", 48.00, 8,  ""),
        (sem2, "2025级", "计算机学院", "软件工程",   "软件工程1班",   "软件工程导论",       "清华大学出版社", "周九",   "978-7-302-60011-7", 50.00, 8,  ""),
        (sem2, "2025级", "计算机学院", "计算机科学", "计算机科学1班", "计算机组成原理",     "清华大学出版社", "王武",   "978-7-302-60008-7", 52.00, 8,  ""),
        (sem2, "2025级", "计算机学院", "计算机科学", "计算机科学1班", "人工智能导论",       "清华大学出版社", "李思",   "978-7-302-60014-8", 49.00, 8,  ""),
        (sem2, "2025级", "数学学院",   "应用数学",   "应用数学1班",   "高等数学（下）",     "高等教育出版社", "陈七",   "978-7-04-60009-3", 41.00, 8,  ""),
        (sem2, "2025级", "数学学院",   "应用数学",   "应用数学1班",   "概率论与数理统计",   "高等教育出版社", "刘八",   "978-7-04-60010-3", 35.00, 8,  ""),
    ]
    textbook_ids = {}
    for t in textbooks_data:
        try:
            tid = execute_sql(
                """INSERT INTO textbooks (semester_id, grade, college, major, class_name, name,
                   publisher, editor, isbn, price, quantity, remark)
                   VALUES (%s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s)""",
                t
            )
            key = (t[0], t[5])  # (semester_id, book_name)
            textbook_ids[key] = tid
            stats["textbooks"] += 1
        except Exception:
            pass

    # ── 5. 征订明细 ──
    orders_data = [
        # 第一学期
        (sem1, master_ids.get("978-7-302-60001-1"), "2025级", "计算机学院", "软件工程",   "软件工程1班",   8,  ""),
        (sem1, master_ids.get("978-7-302-60002-8"), "2025级", "计算机学院", "软件工程",   "软件工程1班",   8,  ""),
        (sem1, master_ids.get("978-7-302-60016-2"), "2025级", "计算机学院", "软件工程",   "软件工程1班",   8,  ""),
        (sem1, master_ids.get("978-7-121-60003-5"), "2025级", "计算机学院", "计算机科学", "计算机科学1班", 8,  ""),
        (sem1, master_ids.get("978-7-111-60004-2"), "2025级", "计算机学院", "计算机科学", "计算机科学1班", 8,  ""),
        (sem1, master_ids.get("978-7-04-60005-9"),  "2025级", "数学学院",   "应用数学",   "应用数学1班",   8,  ""),
        (sem1, master_ids.get("978-7-04-60006-6"),  "2025级", "数学学院",   "应用数学",   "应用数学1班",   8,  ""),
        # 第二学期
        (sem2, master_ids.get("978-7-04-60007-0"),  "2025级", "计算机学院", "软件工程",   "软件工程1班",   8,  ""),
        (sem2, master_ids.get("978-7-302-60011-7"), "2025级", "计算机学院", "软件工程",   "软件工程1班",   8,  ""),
        (sem2, master_ids.get("978-7-302-60008-7"), "2025级", "计算机学院", "计算机科学", "计算机科学1班", 8,  ""),
        (sem2, master_ids.get("978-7-302-60014-8"), "2025级", "计算机学院", "计算机科学", "计算机科学1班", 8,  ""),
        (sem2, master_ids.get("978-7-04-60009-3"),  "2025级", "数学学院",   "应用数学",   "应用数学1班",   8,  ""),
        (sem2, master_ids.get("978-7-04-60010-3"),  "2025级", "数学学院",   "应用数学",   "应用数学1班",   8,  ""),
    ]
    for o in orders_data:
        try:
            execute_sql(
                """INSERT INTO textbook_orders (semester_id, textbook_id, grade, college, major, class_name, quantity, remark)
                   VALUES (%s, %s, %s, %s, %s, %s, %s, %s)""",
                o
            )
            stats["orders"] += 1
        except Exception:
            pass

    # ── 6. 发放记录（第一学期：每人领取本班教材各1本） ──
    all_students = query_df("SELECT id, class_name FROM students ORDER BY class_name, id")
    all_textbooks = query_df("SELECT id, semester_id, class_name FROM textbooks ORDER BY semester_id, class_name")

    # 按 (学期, 班级) 分组教材
    sem_class_books = {}
    for _, t in all_textbooks.iterrows():
        sid = int(t["semester_id"])
        cls = t["class_name"]
        sem_class_books.setdefault(sid, {}).setdefault(cls, []).append(int(t["id"]))

    # 按班级分组学生
    class_students = {}
    for _, s in all_students.iterrows():
        cls = s["class_name"]
        class_students.setdefault(cls, []).append(int(s["id"]))

    # 只为第一学期创建发放记录
    if sem1 and sem1 in sem_class_books:
        for cls, book_ids in sem_class_books[sem1].items():
            for stu_id in class_students.get(cls, []):
                for bid in book_ids:
                    try:
                        execute_sql(
                            "INSERT INTO distributions (student_id, textbook_id, quantity, distribute_date, handler) VALUES (%s, %s, 1, %s, '管理员')",
                            (stu_id, bid, "2025-09-10")
                        )
                        stats["distributions"] += 1
                    except Exception:
                        pass

    # ── 7. V2.0 核对确认演示数据 ──
    try:
        from database_v2 import upsert_confirmation, add_notification

        # 为第一学期的部分学生创建确认记录
        if sem1:
            confirm_students = query_df(
                "SELECT id, name FROM students WHERE class_name = '软件工程1班' LIMIT 4"
            )
            for _, s in confirm_students.iterrows():
                stu_id = int(s["id"])
                # 前两个学生已确认
                if stu_id <= int(confirm_students.iloc[1]["id"]):
                    upsert_confirmation(stu_id, sem1, "confirmed")
                else:
                    upsert_confirmation(stu_id, sem1, "pending")

            # 发一条通知
            if not confirm_students.empty:
                add_notification(
                    int(confirm_students.iloc[0]["id"]),
                    sem1,
                    "📢 请尽快完成第一学期教材核对确认",
                    "confirmation_request"
                )
    except Exception:
        pass

    return stats
