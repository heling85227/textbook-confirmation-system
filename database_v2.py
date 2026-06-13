"""
数据库 V3.0 扩展模块
====================

交互确认增强功能新增的表和 CRUD：
- student_confirmations: 学生核对确认表
- student_notifications: 学生通知表

使用方式：
    from database_v2 import init_v2_tables, ...
    init_v2_tables(cur, conn)  # 在 init_db 末尾调用
"""

import sqlite3
from datetime import datetime

# pymysql 仅在 MySQL 模式下按需导入
try:
    from config import DB_TYPE
except ImportError:
    DB_TYPE = "sqlite"

if DB_TYPE == "mysql":
    try:
        import pymysql
    except ImportError:
        pymysql = None


# ═════════════════════════════════════════════════════════
# 建表 + 迁移
# ═════════════════════════════════════════════════════════

def init_v2_tables(cur, conn):
    """在 init_db 末尾调用，创建 V3.0 新增表和索引"""
    tables = [
        # 学生核对确认表
        """CREATE TABLE IF NOT EXISTS student_confirmations (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            student_id INTEGER NOT NULL,
            semester_id INTEGER NOT NULL,
            textbook_id INTEGER,
            status VARCHAR(20) DEFAULT 'pending',
            feedback_type VARCHAR(50),
            feedback_detail TEXT,
            admin_response TEXT,
            admin_action VARCHAR(50),
            confirmed_at DATETIME,
            updated_at DATETIME DEFAULT CURRENT_TIMESTAMP,
            created_at DATETIME DEFAULT CURRENT_TIMESTAMP,
            FOREIGN KEY (student_id) REFERENCES students(id) ON DELETE CASCADE,
            FOREIGN KEY (semester_id) REFERENCES semesters(id) ON DELETE CASCADE,
            FOREIGN KEY (textbook_id) REFERENCES textbooks_master(id) ON DELETE SET NULL,
            UNIQUE(student_id, semester_id, textbook_id)
        )""",
        # 学生通知表
        """CREATE TABLE IF NOT EXISTS student_notifications (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            student_id INTEGER NOT NULL,
            semester_id INTEGER NOT NULL,
            type VARCHAR(50) NOT NULL,
            title VARCHAR(200) NOT NULL,
            content TEXT,
            is_read INTEGER DEFAULT 0,
            created_at DATETIME DEFAULT CURRENT_TIMESTAMP,
            read_at DATETIME,
            FOREIGN KEY (student_id) REFERENCES students(id) ON DELETE CASCADE,
            FOREIGN KEY (semester_id) REFERENCES semesters(id) ON DELETE CASCADE
        )""",
    ]

    for sql in tables:
        try:
            cur.execute(sql)
        except Exception as e:
            print(f"[V3.0] 建表警告: {e}")

    # 索引
    indexes = [
        "CREATE INDEX IF NOT EXISTS idx_conf_student ON student_confirmations(student_id)",
        "CREATE INDEX IF NOT EXISTS idx_conf_semester ON student_confirmations(semester_id)",
        "CREATE INDEX IF NOT EXISTS idx_conf_status ON student_confirmations(status)",
        "CREATE INDEX IF NOT EXISTS idx_conf_feedback ON student_confirmations(feedback_type)",
        "CREATE INDEX IF NOT EXISTS idx_notif_student ON student_notifications(student_id)",
        "CREATE INDEX IF NOT EXISTS idx_notif_semester ON student_notifications(semester_id)",
        "CREATE INDEX IF NOT EXISTS idx_notif_read ON student_notifications(is_read)",
        "CREATE INDEX IF NOT EXISTS idx_notif_type ON student_notifications(type)",
    ]
    for sql in indexes:
        try:
            cur.execute(sql)
        except Exception:
            pass

    conn.commit()
    print("[V3.0] 新增表和索引创建完成")


# ═════════════════════════════════════════════════════════
# 确认记录 CRUD
# ═════════════════════════════════════════════════════════

# 反馈类型枚举
FEEDBACK_TYPES = [
    "退书", "补领", "少领", "教材分配错误",
    "重复发放", "价格疑问", "不需要", "其他"
]

# 确认状态枚举
CONFIRMATION_STATUS = ["pending", "confirmed", "disputed"]

# 通知类型枚举
NOTIFICATION_TYPES = [
    "confirmation_required",    # 需要核对确认
    "feedback_processed",       # 反馈已处理
    "feedback_rejected",        # 反馈已驳回
    "new_semester",            # 新学期通知
    "system",                  # 系统通知
]


def upsert_confirmation(student_id, semester_id, textbook_id=None,
                         status="pending", feedback_type=None,
                         feedback_detail=None):
    """
    插入或更新确认记录。
    如果 student_id+semester_id+textbook_id 组合已存在则更新，否则插入。
    返回记录 ID。
    """
    from database import execute_sql, query_df

    # 查找现有记录
    existing = query_df(
        """SELECT id FROM student_confirmations
           WHERE student_id = %s AND semester_id = %s AND (
               (textbook_id IS NULL AND %s IS NULL) OR textbook_id = %s
           )""",
        (student_id, semester_id, textbook_id, textbook_id)
    )

    now = datetime.now().strftime("%Y-%m-%d %H:%M:%S")

    if not existing.empty:
        # 更新
        rid = int(existing.iloc[0]["id"])
        sets = ["status = %s", "updated_at = %s"]
        params = [status, now]
        if feedback_type is not None:
            sets.append("feedback_type = %s")
            params.append(feedback_type)
        if feedback_detail is not None:
            sets.append("feedback_detail = %s")
            params.append(feedback_detail)
        if status == "confirmed":
            sets.append("confirmed_at = %s")
            params.append(now)
        params.append(rid)
        execute_sql(
            f"UPDATE student_confirmations SET {', '.join(sets)} WHERE id = %s",
            tuple(params)
        )
        return rid
    else:
        # 插入
        return execute_sql(
            """INSERT INTO student_confirmations
               (student_id, semester_id, textbook_id, status, feedback_type,
                feedback_detail, created_at, updated_at)
               VALUES (%s, %s, %s, %s, %s, %s, %s, %s)""",
            (student_id, semester_id, textbook_id, status, feedback_type,
             feedback_detail, now, now)
        )


def get_confirmations(student_id=None, semester_id=None, status=None,
                       feedback_type=None):
    """查询确认记录，返回 DataFrame"""
    from database import query_df

    conds = []
    params = []
    if student_id is not None:
        conds.append("sc.student_id = %s")
        params.append(student_id)
    if semester_id is not None:
        conds.append("sc.semester_id = %s")
        params.append(semester_id)
    if status is not None:
        conds.append("sc.status = %s")
        params.append(status)
    if feedback_type is not None:
        conds.append("sc.feedback_type = %s")
        params.append(feedback_type)

    where = " AND ".join(conds) if conds else "1=1"

    sql = f"""
        SELECT sc.*, s.name as student_name, s.student_id as sid, s.class_name,
               sem.name as semester_name,
               tm.name as textbook_name
        FROM student_confirmations sc
        LEFT JOIN students s ON sc.student_id = s.id
        LEFT JOIN semesters sem ON sc.semester_id = sem.id
        LEFT JOIN textbooks_master tm ON sc.textbook_id = tm.id
        WHERE {where}
        ORDER BY sc.updated_at DESC
    """
    return query_df(sql, tuple(params) if params else None)


def get_confirmation_stats(semester_id):
    """获取某学期的确认统计"""
    from database import query_df

    return query_df("""
        SELECT
            COUNT(*) as total,
            SUM(CASE WHEN status='pending' THEN 1 ELSE 0 END) as pending_count,
            SUM(CASE WHEN status='confirmed' THEN 1 ELSE 0 END) as confirmed_count,
            SUM(CASE WHEN status='disputed' THEN 1 ELSE 0 END) as disputed_count
        FROM student_confirmations
        WHERE semester_id = %s
    """, (semester_id,))


# ═════════════════════════════════════════════════════════
# 通知 CRUD
# ═════════════════════════════════════════════════════════

def add_notification(student_id, semester_id, ntype, title, content=None):
    """创建一条通知，返回通知 ID"""
    from database import execute_sql
    now = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    return execute_sql(
        """INSERT INTO student_notifications
           (student_id, semester_id, type, title, content, is_read, created_at)
           VALUES (%s, %s, %s, %s, %s, 0, %s)""",
        (student_id, semester_id, ntype, title, content, now)
    )


def get_notifications(student_id, semester_id=None, unread_only=False):
    """查询学生通知，返回 DataFrame"""
    from database import query_df

    conds = ["sn.student_id = %s"]
    params = [student_id]
    if semester_id is not None:
        conds.append("sn.semester_id = %s")
        params.append(semester_id)
    if unread_only:
        conds.append("sn.is_read = 0")

    where = " AND ".join(conds)
    sql = f"""
        SELECT sn.*, sem.name as semester_name
        FROM student_notifications sn
        LEFT JOIN semesters sem ON sn.semester_id = sem.id
        WHERE {where}
        ORDER BY sn.created_at DESC
    """
    return query_df(sql, tuple(params))


def get_unread_count(student_id, semester_id=None):
    """获取未读通知数"""
    from database import query_df

    if semester_id is not None:
        df = query_df(
            "SELECT COUNT(*) as cnt FROM student_notifications WHERE student_id = %s AND is_read = 0 AND semester_id = %s",
            (student_id, semester_id)
        )
    else:
        df = query_df(
            "SELECT COUNT(*) as cnt FROM student_notifications WHERE student_id = %s AND is_read = 0",
            (student_id,)
        )
    return int(df.iloc[0]["cnt"]) if not df.empty else 0


def mark_notification_read(notification_id):
    """标记通知为已读"""
    from database import execute_sql
    now = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    execute_sql(
        "UPDATE student_notifications SET is_read = 1, read_at = %s WHERE id = %s",
        (now, notification_id)
    )


def mark_all_notifications_read(student_id, semester_id=None):
    """标记学生所有通知为已读"""
    from database import execute_sql
    now = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    if semester_id is not None:
        execute_sql(
            "UPDATE student_notifications SET is_read = 1, read_at = %s WHERE student_id = %s AND is_read = 0 AND semester_id = %s",
            (now, student_id, semester_id)
        )
    else:
        execute_sql(
            "UPDATE student_notifications SET is_read = 1, read_at = %s WHERE student_id = %s AND is_read = 0",
            (now, student_id)
        )


# ═════════════════════════════════════════════════════════
# 反馈处理
# ═════════════════════════════════════════════════════════

def process_feedback(confirmation_id, action, admin_response=None):
    """
    管理员处理学生反馈。

    Args:
        confirmation_id: 确认记录 ID
        action: "accept" 或 "reject"
        admin_response: 管理员回复内容

    Returns:
        (success, message) 元组
    """
    from database import query_df, execute_sql

    # 获取确认记录
    conf = query_df(
        """SELECT sc.*, s.name as student_name, s.student_id as sid, sem.id as sem_id, sem.name as sem_name
           FROM student_confirmations sc
           LEFT JOIN students s ON sc.student_id = s.id
           LEFT JOIN semesters sem ON sc.semester_id = sem.id
           WHERE sc.id = %s""",
        (confirmation_id,)
    )

    if conf.empty:
        return False, "确认记录不存在"

    record = conf.iloc[0].to_dict()
    feedback_type = record.get("feedback_type", "")
    student_id = int(record["student_id"])
    semester_id = int(record["sem_id"])

    now = datetime.now().strftime("%Y-%m-%d %H:%M:%S")

    # 更新确认记录
    new_status = "pending" if action == "accept" else "confirmed"
    admin_action = f"accepted_{feedback_type}" if action == "accept" else "rejected"

    execute_sql(
        """UPDATE student_confirmations
           SET status = %s, admin_response = %s, admin_action = %s, updated_at = %s
           WHERE id = %s""",
        (new_status, admin_response, admin_action, now, confirmation_id)
    )

    # 发送通知给学生
    if action == "accept":
        notif_title = "反馈已处理"
        notif_content = f"您提交的反馈（{feedback_type}）已被接受处理"
        if admin_response:
            notif_content += f"：{admin_response}"
    else:
        notif_title = "反馈已驳回"
        notif_content = f"您提交的反馈（{feedback_type}）已被驳回"
        if admin_response:
            notif_content += f"，原因：{admin_response}"

    add_notification(student_id, semester_id,
                     "feedback_processed" if action == "accept" else "feedback_rejected",
                     notif_title, notif_content)

    return True, f"反馈已{'接受' if action == 'accept' else '驳回'}"


# ═════════════════════════════════════════════════════════
# 批量操作
# ═════════════════════════════════════════════════════════

def batch_create_confirmations(semester_id, class_name=None, college=None):
    """
    为指定学期的学生批量创建待确认记录。

    根据征订表中的班级-教材关联，为每个学生×教材创建一条 pending 记录。
    如果已存在则跳过。

    Returns:
        (created_count, skipped_count)
    """
    from database import query_df

    # 查询该学期有征订记录的学生
    conditions = ["o.semester_id = %s"]
    params = [semester_id]

    if class_name:
        conditions.append("s.class_name = %s")
        params.append(class_name)
    if college:
        conditions.append("s.college = %s")
        params.append(college)

    where = " AND ".join(conditions)

    # 学生-教材对
    pairs = query_df(f"""
        SELECT DISTINCT s.id as student_id, o.textbook_id
        FROM students s
        JOIN textbook_orders o ON s.class_name = o.class_name
        WHERE {where}
        ORDER BY s.id, o.textbook_id
    """, tuple(params))

    if pairs.empty:
        return 0, 0

    created = 0
    skipped = 0

    for _, row in pairs.iterrows():
        sid = int(row["student_id"])
        tid = int(row["textbook_id"])

        # 检查是否已存在
        existing = query_df(
            """SELECT id FROM student_confirmations
               WHERE student_id = %s AND semester_id = %s AND textbook_id = %s""",
            (sid, semester_id, tid)
        )
        if not existing.empty:
            skipped += 1
            continue

        upsert_confirmation(sid, semester_id, tid, status="pending")
        created += 1

    return created, skipped


def send_confirmation_notification(semester_id, class_name=None, college=None):
    """
    向学生发送核对确认通知。

    Returns:
        发送的通知数量
    """
    from database import query_df

    conditions = ["o.semester_id = %s"]
    params = [semester_id]
    if class_name:
        conditions.append("s.class_name = %s")
        params.append(class_name)
    if college:
        conditions.append("s.college = %s")
        params.append(college)

    where = " AND ".join(conditions)

    students = query_df(f"""
        SELECT DISTINCT s.id, s.name
        FROM students s
        JOIN textbook_orders o ON s.class_name = o.class_name
        WHERE {where}
    """, tuple(params))

    if students.empty:
        return 0

    sem_df = query_df("SELECT name FROM semesters WHERE id = %s", (semester_id,))
    sem_name = sem_df.iloc[0]["name"] if not sem_df.empty else "本学期"

    count = 0
    for _, stu in students.iterrows():
        add_notification(
            int(stu["id"]), semester_id, "confirmation_required",
            f"请核对{sem_name}教材",
            f"您好，{sem_name}的教材已发放完毕，请登录系统核对您的教材信息并签名确认。"
        )
        count += 1

    return count
