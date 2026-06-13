"""
测试数据生成模块

仅在开发/测试环境使用，生产环境可安全移除。
"""
from database import query_df, execute_sql


def generate_test_data() -> tuple:
    """
    生成完整的测试数据：
    - 3个学期（2024-2025 两学期 + 2025-2026 第二学期）
    - 12名学生，分3个班级
    - 每个学期各班级若干教材
    - 每个学期都有发放记录

    Returns:
        (学生数, 教材数, 涉及学期数)
    """
    # 1. 创建学期
    semesters_data = [
        ("2024-2025 第一学期", "2024-2025", "第一学期"),
        ("2024-2025 第二学期", "2024-2025", "第二学期"),
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
        except Exception:
            df = query_df("SELECT id FROM semesters WHERE name = %s", (name,))
            if not df.empty:
                sem_ids[name] = int(df.iloc[0]["id"])

    # 2. 创建学生（3个班级，各4人）
    students_data = [
        ("110101200501010001", "2024001", "张三",   "2024级", "计算机学院", "软件工程", "软件工程1班"),
        ("110101200501010002", "2024002", "李四",   "2024级", "计算机学院", "软件工程", "软件工程1班"),
        ("110101200501010003", "2024003", "王五",   "2024级", "计算机学院", "软件工程", "软件工程1班"),
        ("110101200501010004", "2024004", "赵六",   "2024级", "计算机学院", "软件工程", "软件工程1班"),
        ("110101200502010001", "2024005", "钱七",   "2024级", "计算机学院", "计算机科学", "计算机科学1班"),
        ("110101200502010002", "2024006", "孙八",   "2024级", "计算机学院", "计算机科学", "计算机科学1班"),
        ("110101200502010003", "2024007", "周九",   "2024级", "计算机学院", "计算机科学", "计算机科学1班"),
        ("110101200502010004", "2024008", "吴十",   "2024级", "计算机学院", "计算机科学", "计算机科学1班"),
        ("110101200503010001", "2024009", "郑十一", "2024级", "数学学院",   "应用数学",   "应用数学1班"),
        ("110101200503010002", "2024010", "王十二", "2024级", "数学学院",   "应用数学",   "应用数学1班"),
        ("110101200503010003", "2024011", "李十三", "2024级", "数学学院",   "应用数学",   "应用数学1班"),
        ("110101200503010004", "2024012", "赵十四", "2024级", "数学学院",   "应用数学",   "应用数学1班"),
    ]
    for id_card, sid, name, grade, college, major, cls in students_data:
        try:
            execute_sql(
                """INSERT INTO students (id_card,student_id,name,grade,college,major,class_name)
                   VALUES (%s,%s,%s,%s,%s,%s,%s)""",
                (id_card, sid, name, grade, college, major, cls)
            )
        except Exception:
            pass

    # 3. 创建教材（每个学期，每个班级2-3本教材）
    sem_df = query_df("SELECT id, name FROM semesters ORDER BY id")
    sem_id_map = {row["name"]: int(row["id"]) for _, row in sem_df.iterrows()}

    textbooks_data = [
        (sem_id_map.get("2024-2025 第一学期"), "2024级", "计算机学院", "软件工程",   "软件工程1班", "Python程序设计",       "清华大学出版社", "张三丰", "978-7-302-00001-1", 59.00, 5, ""),
        (sem_id_map.get("2024-2025 第一学期"), "2024级", "计算机学院", "软件工程",   "软件工程1班", "数据结构与算法",     "清华大学出版社", "李思",   "978-7-302-00002-8", 45.00, 3, ""),
        (sem_id_map.get("2024-2025 第一学期"), "2024级", "计算机学院", "计算机科学", "计算机科学1班", "计算机网络",      "电子工业出版社", "王武",   "978-7-121-00003-5", 42.00, 4, ""),
        (sem_id_map.get("2024-2025 第一学期"), "2024级", "计算机学院", "计算机科学", "计算机科学1班", "操作系统原理",    "机械工业出版社", "赵柳",   "978-7-111-00004-2", 55.00, 6, ""),
        (sem_id_map.get("2024-2025 第一学期"), "2024级", "数学学院",   "应用数学",   "应用数学1班",   "高等数学（上）",  "高等教育出版社", "陈七",   "978-7-04-00005-9", 38.00, 4, ""),
        (sem_id_map.get("2024-2025 第一学期"), "2024级", "数学学院",   "应用数学",   "应用数学1班",   "线性代数",       "高等教育出版社", "刘八",   "978-7-04-00006-6", 32.00, 2, ""),
        (sem_id_map.get("2024-2025 第二学期"), "2024级", "计算机学院", "软件工程",   "软件工程1班", "Python程序设计（下）", "清华大学出版社", "张三丰", "978-7-302-00007-3", 62.00, 4, ""),
        (sem_id_map.get("2024-2025 第二学期"), "2024级", "计算机学院", "软件工程",   "软件工程1班", "数据库系统概论",    "高等教育出版社", "李思",   "978-7-04-00008-0", 48.00, 4, ""),
        (sem_id_map.get("2024-2025 第二学期"), "2024级", "计算机学院", "计算机科学", "计算机科学1班", "计算机组成原理",   "清华大学出版社", "王武",   "978-7-302-00009-7", 52.00, 4, ""),
        (sem_id_map.get("2024-2025 第二学期"), "2024级", "数学学院",   "应用数学",   "应用数学1班",   "高等数学（下）",   "高等教育出版社", "陈七",   "978-7-04-00010-3", 41.00, 4, ""),
        (sem_id_map.get("2024-2025 第二学期"), "2024级", "数学学院",   "应用数学",   "应用数学1班",   "概率论与数理统计", "高等教育出版社", "刘八",   "978-7-04-00011-0", 35.00, 4, ""),
        (sem_id_map.get("2025-2026 第二学期"), "2024级", "计算机学院", "软件工程",   "软件工程1班", "软件工程导论",     "清华大学出版社", "周九",   "978-7-302-00012-7", 50.00, 4, ""),
        (sem_id_map.get("2025-2026 第二学期"), "2024级", "计算机学院", "软件工程",   "软件工程1班", "Web前端开发",       "电子工业出版社", "吴十",   "978-7-121-00013-4", 46.00, 4, ""),
        (sem_id_map.get("2025-2026 第二学期"), "2024级", "计算机学院", "计算机科学", "计算机科学1班", "编译原理",         "机械工业出版社", "赵柳",   "978-7-111-00014-1", 54.00, 4, ""),
        (sem_id_map.get("2025-2026 第二学期"), "2024级", "计算机学院", "计算机科学", "计算机科学1班", "人工智能导论",     "清华大学出版社", "李思",   "978-7-302-00015-8", 49.00, 4, ""),
        (sem_id_map.get("2025-2026 第二学期"), "2024级", "数学学院",   "应用数学",   "应用数学1班",   "数学建模",         "高等教育出版社", "陈七",   "978-7-04-00016-5", 37.00, 4, ""),
    ]
    for t in textbooks_data:
        try:
            execute_sql(
                """INSERT INTO textbooks (semester_id,grade,college,major,class_name,name,course_name,publisher,editor,isbn,price,quantity,remark)
                   VALUES (%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s)""",
                t
            )
        except Exception:
            pass

    # 4. 创建发放记录（每个学生在每个学期领取本班所有教材各1本）
    all_students = query_df("SELECT id, student_id, class_name FROM students ORDER BY class_name, student_id")
    all_textbooks = query_df("SELECT id, semester_id, class_name FROM textbooks ORDER BY semester_id, class_name")

    sem_class_books = {}
    for _, t in all_textbooks.iterrows():
        sid = int(t["semester_id"])
        cls = t["class_name"]
        sem_class_books.setdefault(sid, {}).setdefault(cls, []).append(int(t["id"]))

    class_students = {}
    for _, s in all_students.iterrows():
        cls = s["class_name"]
        class_students.setdefault(cls, []).append((int(s["id"]), s["student_id"]))

    distribute_dates = {
        "2024-2025 第一学期": "2024-09-10",
        "2024-2025 第二学期": "2025-02-20",
        "2025-2026 第二学期": "2026-02-25",
    }
    for sid, class_map in sem_class_books.items():
        sem_name = query_df("SELECT name FROM semesters WHERE id = %s", (sid,))["name"].iloc[0]
        d_date = distribute_dates.get(sem_name, "2026-05-01")
        for cls, book_ids in class_map.items():
            for stu_id, stu_sid in class_students.get(cls, []):
                for bid in book_ids:
                    try:
                        execute_sql(
                            "INSERT INTO distributions (student_id,textbook_id,quantity,distribute_date,handler) VALUES (%s,%s,1,%s,'管理员')",
                            (stu_id, bid, d_date)
                        )
                    except Exception:
                        pass

    return len(students_data), len(textbooks_data), len(sem_class_books)
