"""
演示数据自动初始化模块
======================
当数据库为空时（如 Streamlit Cloud 首次部署），
或关键表缺少数据时，自动创建演示数据，使系统可直接运行演示。

逻辑：逐表检测、按需补全（不会覆盖已有数据）。
"""
from database import query_df, execute_sql


def _count(table: str) -> int:
    """安全查询某表的记录数"""
    try:
        df = query_df(f"SELECT COUNT(*) as cnt FROM {table}")
        return 0 if df.empty else int(df.iloc[0]["cnt"])
    except Exception:
        return 0


def _get_semester_ids():
    """获取学期 ID 映射（第一/第二学期）"""
    sem_df = query_df("SELECT id, name FROM semesters ORDER BY id")
    result = {}
    sem1_id = None
    sem2_id = None
    for _, row in sem_df.iterrows():
        name = row["name"]
        sid = int(row["id"])
        result[name] = sid
        if "第一" in name:
            sem1_id = sid
        elif "第二" in name:
            sem2_id = sid
    return result, sem1_id, sem2_id


def needs_demo_data() -> bool:
    """检测是否需要补充演示数据（任一关键表为空即触发）"""
    return (
        _count("semesters") == 0 or
        _count("students") == 0 or
        _count("textbooks_master") == 0 or
        _count("textbooks") == 0
    )


def init_demo_data() -> dict:
    """
    自动创建/补全演示数据。
    对每个表独立判断是否为空，只插入缺失的数据，不覆盖已有记录。

    Returns:
        dict: 创建的各记录数量统计
    """
    stats = {}

    # ── 1. 学期 ──
    if _count("semesters") == 0:
        semesters_data = [
            ("2025-2026 第一学期", "2025-2026", "第一学期"),
            ("2025-2026 第二学期", "2025-2026", "第二学期"),
        ]
        for name, ay, sn in semesters_data:
            try:
                execute_sql(
                    "INSERT INTO semesters (name, academic_year, semester_name) VALUES (%s, %s, %s)",
                    (name, ay, sn)
                )
                stats["semesters"] = stats.get("semesters", 0) + 1
            except Exception:
                pass

    # 无论学期是否为空，都获取最新 ID 映射
    sem_ids, sem1_id, sem2_id = _get_semester_ids()
    stats["semester_ids"] = sem_ids

    # ── 2. 教材主表（核心！教材管理页面依赖此表）──
    if _count("textbooks_master") == 0:
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
                stats["master_books"] = stats.get("master_books", 0) + 1
            except Exception:
                pass
    else:
        master_df = query_df("SELECT id, isbn FROM textbooks_master")
        master_ids = {row["isbn"]: int(row["id"]) for _, row in master_df.iterrows() if row.get("isbn")}

    # 无论教材主表是否为空，都获取最新映射
    master_df = query_df("SELECT id, isbn FROM textbooks_master")
    master_ids = {row["isbn"]: int(row["id"]) for _, row in master_df.iterrows() if row.get("isbn")}
    stats["master_ids"] = master_ids

    # ── 3. 学生（仅空表时填充）──
    if _count("students") == 0:
        students_data = [
            ("320101200501010001", "2025001", "张明",   "2025级", "计算机学院", "软件工程", "软件工程1班"),
            ("320101200501010002", "2025002", "李华",   "2025级", "计算机学院", "软件工程", "软件工程1班"),
            ("320101200501010003", "2025003", "王芳",   "2025级", "计算机学院", "软件工程", "软件工程1班"),
            ("320101200501010004", "2025004", "赵强",   "2025级", "计算机学院", "软件工程", "软件工程1班"),
            ("320101200501010005", "2025005", "刘洋",   "2025级", "计算机学院", "软件工程", "软件工程1班"),
            ("320101200501010006", "2025006", "陈静",   "2025级", "计算机学院", "软件工程", "软件工程1班"),
            ("320101200501010007", "2025007", "杨磊",   "2025级", "计算机学院", "软件工程", "软件工程1班"),
            ("320101200501010008", "2025008", "黄丽",   "2025级", "计算机学院", "软件工程", "软件工程1班"),
            ("320101200502010001", "2025009", "周杰",   "2025级", "计算机学院", "计算机科学", "计算机科学1班"),
            ("320101200502010002", "2025010", "吴敏",   "2025级", "计算机学院", "计算机科学", "计算机科学1班"),
            ("320101200502010003", "2025011", "郑伟",   "2025级", "计算机学院", "计算机科学", "计算机科学1班"),
            ("320101200502010004", "2025012", "孙燕",   "2025级", "计算机学院", "计算机科学", "计算机科学1班"),
            ("320101200503010001", "2025017", "马超",   "2025级", "数学学院",   "应用数学", "应用数学1班"),
            ("320101200503010002", "2025018", "高颖",   "2025级", "数学学院",   "应用数学", "应用数学1班"),
            ("320101200503010003", "2025019", "罗浩",   "2025级", "数学学院",   "应用数学", "应用数学1班"),
            ("320101200503010004", "2025020", "梁欣",   "2025级", "数学学院",   "应用数学", "应用数学1班"),
        ]
        for id_card, sid, name, grade, college, major, cls in students_data:
            try:
                execute_sql(
                    """INSERT INTO students (id_card, student_id, name, grade, college, major, class_name)
                       VALUES (%s, %s, %s, %s, %s, %s, %s)""",
                    (id_card, sid, name, grade, college, major, cls)
                )
                stats["students"] = stats.get("students", 0) + 1
            except Exception:
                pass

    # ── 4. 旧版 textbooks 表（仅空表时填充）──
    if _count("textbooks") == 0 and sem1_id and master_ids:
        textbooks_data = [
            (sem1_id, "2025级", "计算机学院", "软件工程",   "软件工程1班",   "Python程序设计",     "清华大学出版社", "张三丰", "978-7-302-60001-1", 59.00, 8,  ""),
            (sem1_id, "2025级", "计算机学院", "软件工程",   "软件工程1班",   "数据结构与算法",     "清华大学出版社", "李思",   "978-7-302-60002-8", 45.00, 8,  ""),
            (sem1_id, "2025级", "计算机学院", "软件工程",   "软件工程1班",   "离散数学",           "清华大学出版社", "郑伟",   "978-7-302-60016-2", 40.00, 8,  ""),
            (sem1_id, "2025级", "计算机学院", "计算机科学", "计算机科学1班", "计算机网络",         "电子工业出版社", "王武",   "978-7-121-60003-5", 42.00, 8,  ""),
            (sem1_id, "2025级", "计算机学院", "计算机科学", "计算机科学1班", "操作系统原理",       "机械工业出版社", "赵柳",   "978-7-111-60004-2", 55.00, 8,  ""),
            (sem1_id, "2025级", "数学学院",   "应用数学",   "应用数学1班",   "高等数学（上）",     "高等教育出版社", "陈七",   "978-7-04-60005-9", 38.00, 8,  ""),
            (sem1_id, "2025级", "数学学院",   "应用数学",   "应用数学1班",   "线性代数",           "高等教育出版社", "刘八",   "978-7-04-60006-6", 32.00, 8,  ""),
        ]
        for t in textbooks_data:
            try:
                execute_sql(
                    """INSERT INTO textbooks (semester_id, grade, college, major, class_name, name,
                       publisher, editor, isbn, price, quantity, remark)
                       VALUES (%s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s)""",
                    t
                )
                stats["textbooks"] = stats.get("textbooks", 0) + 1
            except Exception:
                pass

    # ── 5. 发放记录（仅空表且有学生和教材时）──
    if _count("distributions") == 0 and _count("students") > 0 and _count("textbooks") > 0:
        all_students = query_df("SELECT id, class_name FROM students ORDER BY class_name, id")
        all_textbooks = query_df("SELECT id, semester_id, class_name FROM textbooks ORDER BY semester_id, class_name")

        sem_class_books = {}
        for _, t in all_textbooks.iterrows():
            sid = int(t["semester_id"])
            cls = t["class_name"]
            sem_class_books.setdefault(sid, {}).setdefault(cls, []).append(int(t["id"]))

        class_students = {}
        for _, s in all_students.iterrows():
            cls = s["class_name"]
            class_students.setdefault(cls, []).append(int(s["id"]))

        for sid, class_map in sem_class_books.items():
            for cls, book_ids in class_map.items():
                for stu_id in class_students.get(cls, []):
                    for bid in book_ids:
                        try:
                            execute_sql(
                                "INSERT INTO distributions (student_id, textbook_id, quantity, distribute_date, handler) VALUES (%s, %s, 1, '2025-09-10', '管理员')",
                                (stu_id, bid)
                            )
                            stats["distributions"] = stats.get("distributions", 0) + 1
                        except Exception:
                            pass

    # ── 6. V2.0 核对确认演示数据 ──
    try:
        from database_v2 import upsert_confirmation, add_notification
        confirm_count = _count("student_confirmations")
        if confirm_count == 0 and _count("students") > 0 and sem1_id:
            confirm_students = query_df(
                "SELECT id, name FROM students WHERE class_name LIKE '%软件工程%' LIMIT 4"
            )
            if not confirm_students.empty:
                for idx, (_, s) in enumerate(confirm_students.iterrows()):
                    stu_id = int(s["id"])
                    status = "confirmed" if idx < 2 else "pending"
                    upsert_confirmation(stu_id, sem1_id, status=status)

                add_notification(
                    int(confirm_students.iloc[0]["id"]),
                    sem1_id,
                    "📢 请尽快完成第一学期教材核对确认",
                    "confirmation_request"
                )
                stats["confirmations"] = 4
                stats["notifications"] = 1
    except Exception as e:
        stats["v2_error"] = str(e)

    return stats
