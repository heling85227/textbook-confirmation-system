"""
数据库层（双引擎：MySQL + SQLite）

提供统一的数据库访问接口：
- get_connection: 上下文管理器，自动处理连接生命周期
- init_db: 初始化表结构 + 自动迁移
- query_df: 执行查询，返回 DataFrame
- execute_sql: 执行写操作，返回 lastrowid
"""
import os
import sys
import sqlite3
import pandas as pd
from contextlib import contextmanager
from config import DB_TYPE, DB_CONFIG
from pathlib import Path

# pymysql 仅在 MySQL 模式下按需导入
if DB_TYPE == "mysql":
    try:
        import pymysql
    except ImportError:
        pymysql = None

# ── 价格计算 SQL 片段 ──
# 所有涉及教材结算价的查询统一通过 textbooks_master 实时计算
# 优先级：实洋(actual_price) > 定价×折扣率 > 定价(old textbooks)
PRICE_CALC = (
    "COALESCE(tm.actual_price, tm.price * COALESCE(tm.discount_rate, 1), t.price, 0)"
)
PRICE_JOIN = "LEFT JOIN textbooks_master tm ON tm.isbn = t.isbn"


def get_sqlite_path() -> str:
    """返回 SQLite 数据库文件路径
    Streamlit Cloud 容器重启后数据会丢失，但单次运行内可正常读写。
    检测逻辑：如果项目目录下不存在 textbook_data.db（如云端首次部署），
    则使用 /tmp 目录；本地有 .db 文件则继续使用。
    """
    local_db = os.path.join(os.path.dirname(os.path.abspath(__file__)), "textbook_data.db")
    if os.path.exists(local_db):
        # 本地已有数据库文件，直接使用
        return local_db
    else:
        # 云端首次部署（无 .db 文件）或全新本地环境，使用 /tmp
        return "/tmp/textbook_data.db"


@contextmanager
def get_connection():
    """
    统一获取数据库连接（上下文管理器）。

    使用示例:
        with get_connection() as conn:
            df = pd.read_sql("SELECT ...", conn)
    """
    if DB_TYPE == "sqlite":
        conn = sqlite3.connect(get_sqlite_path())
        conn.row_factory = sqlite3.Row
        conn.execute("PRAGMA journal_mode=WAL")
        conn.execute("PRAGMA foreign_keys=ON")
        try:
            yield conn
        finally:
            conn.close()
    else:
        if pymysql is None:
            raise ImportError("MySQL 模式需要 pymysql，请安装：pip install pymysql")
        conn = pymysql.connect(**DB_CONFIG, cursorclass=pymysql.cursors.DictCursor)
        try:
            yield conn
        finally:
            conn.close()


def init_db():
    """初始化数据库表结构 + 自动迁移（MySQL 建库 + SQLite 建表/迁移）"""
    # ── MySQL: 先确保数据库存在 ──
    if DB_TYPE != "sqlite":
        if pymysql is None:
            raise ImportError("MySQL 模式需要 pymysql，请安装：pip install pymysql")
        try:
            tmp_conn = pymysql.connect(
                host=DB_CONFIG["host"], port=DB_CONFIG["port"],
                user=DB_CONFIG["user"], password=DB_CONFIG["password"],
                charset="utf8mb4"
            )
            tmp_cur = tmp_conn.cursor()
            tmp_cur.execute(
                f"CREATE DATABASE IF NOT EXISTS `{DB_CONFIG['database']}` "
                "CHARACTER SET utf8mb4 COLLATE utf8mb4_unicode_ci"
            )
            tmp_conn.commit()
            tmp_cur.close()
            tmp_conn.close()
        except Exception:
            pass

    with get_connection() as conn:
        cur = conn.cursor()

        # ── 核心表定义 ──
        tables_sql = [
            # 学期
            """CREATE TABLE IF NOT EXISTS semesters (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                name VARCHAR(100) UNIQUE NOT NULL,
                academic_year VARCHAR(20),
                semester_name VARCHAR(50),
                created_at DATETIME DEFAULT CURRENT_TIMESTAMP
            )""",
            # 学生
            """CREATE TABLE IF NOT EXISTS students (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                id_card VARCHAR(18) UNIQUE NOT NULL,
                student_id VARCHAR(50) UNIQUE NOT NULL,
                name VARCHAR(50) NOT NULL,
                grade VARCHAR(20),
                college VARCHAR(100),
                major VARCHAR(100),
                class_name VARCHAR(100),
                created_at DATETIME DEFAULT CURRENT_TIMESTAMP
            )""",
            # 教材（旧版，兼容保留）
            """CREATE TABLE IF NOT EXISTS textbooks (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                semester_id INTEGER NOT NULL,
                grade VARCHAR(20),
                college VARCHAR(100),
                major VARCHAR(100),
                class_name VARCHAR(100),
                name VARCHAR(200) NOT NULL,
                publisher VARCHAR(200),
                editor VARCHAR(200),
                course_name VARCHAR(200),
                isbn VARCHAR(50),
                price DECIMAL(10,2) NOT NULL DEFAULT 0,
                quantity INTEGER DEFAULT 0,
                remark TEXT,
                created_at DATETIME DEFAULT CURRENT_TIMESTAMP,
                FOREIGN KEY (semester_id) REFERENCES semesters(id) ON DELETE CASCADE
            )""",
            # 教材主表（新版核心表）
            """CREATE TABLE IF NOT EXISTS textbooks_master (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                name VARCHAR(200) NOT NULL,
                isbn VARCHAR(50),
                publisher VARCHAR(200),
                editor VARCHAR(200),
                price DECIMAL(10,2) DEFAULT 0,
                course_name VARCHAR(200),
                publication_date VARCHAR(50),
                edition VARCHAR(100),
                discount_rate REAL DEFAULT NULL,
                actual_price REAL DEFAULT NULL,
                textbook_type TEXT DEFAULT NULL,
                remark TEXT,
                created_at DATETIME DEFAULT CURRENT_TIMESTAMP
            )""",
            # 教材征订明细（关联主表）
            """CREATE TABLE IF NOT EXISTS textbook_orders (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                semester_id INTEGER NOT NULL,
                textbook_id INTEGER NOT NULL,
                grade VARCHAR(20),
                college VARCHAR(100),
                major VARCHAR(100),
                class_name VARCHAR(100),
                quantity INTEGER DEFAULT 0,
                remark TEXT,
                created_at DATETIME DEFAULT CURRENT_TIMESTAMP,
                FOREIGN KEY (textbook_id) REFERENCES textbooks_master(id)
            )""",
            # 免领标记
            """CREATE TABLE IF NOT EXISTS student_exemptions (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                semester_id INTEGER NOT NULL,
                student_id INTEGER NOT NULL,
                is_exempt INTEGER NOT NULL DEFAULT 0,
                remark TEXT,
                created_at DATETIME DEFAULT CURRENT_TIMESTAMP,
                FOREIGN KEY (semester_id) REFERENCES semesters(id) ON DELETE CASCADE,
                FOREIGN KEY (student_id) REFERENCES students(id) ON DELETE CASCADE,
                UNIQUE(semester_id, student_id)
            )""",
            # 发放记录
            """CREATE TABLE IF NOT EXISTS distributions (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                student_id INTEGER NOT NULL,
                textbook_id INTEGER NOT NULL,
                quantity INTEGER NOT NULL DEFAULT 1,
                distribute_date DATE,
                handler VARCHAR(50),
                created_at DATETIME DEFAULT CURRENT_TIMESTAMP,
                FOREIGN KEY (student_id) REFERENCES students(id) ON DELETE CASCADE,
                FOREIGN KEY (textbook_id) REFERENCES textbooks(id) ON DELETE CASCADE
            )""",
            # 征订总表
            """CREATE TABLE IF NOT EXISTS textbook_subscriptions (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                semester_id INTEGER NOT NULL,
                textbook_id INTEGER,
                book_name VARCHAR(200) NOT NULL,
                isbn VARCHAR(50),
                publisher VARCHAR(200),
                editor VARCHAR(200),
                price DECIMAL(10,2) DEFAULT 0,
                course_name VARCHAR(200),
                college VARCHAR(100),
                major VARCHAR(100),
                grade VARCHAR(20),
                class_scope TEXT,
                class_names TEXT,
                total_qty INTEGER DEFAULT 0,
                teacher_qty INTEGER DEFAULT 0,
                status VARCHAR(20) DEFAULT 'pending',
                remark TEXT,
                source VARCHAR(50) DEFAULT 'manual',
                created_at DATETIME DEFAULT CURRENT_TIMESTAMP,
                dispatched_at DATETIME,
                FOREIGN KEY (semester_id) REFERENCES semesters(id) ON DELETE CASCADE
            )""",
            # 学生签名
            """CREATE TABLE IF NOT EXISTS signatures (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                student_id INTEGER NOT NULL,
                semester_id INTEGER NOT NULL,
                signature_data TEXT NOT NULL,
                signed_at DATETIME DEFAULT CURRENT_TIMESTAMP,
                FOREIGN KEY (student_id) REFERENCES students(id) ON DELETE CASCADE,
                FOREIGN KEY (semester_id) REFERENCES semesters(id) ON DELETE CASCADE,
                UNIQUE(student_id, semester_id)
            )""",
        ]

        for sql in tables_sql:
            cur.execute(sql)

        # ── 索引 ──
        indexes = [
            "CREATE INDEX IF NOT EXISTS idx_students_grade ON students(grade)",
            "CREATE INDEX IF NOT EXISTS idx_students_college ON students(college)",
            "CREATE INDEX IF NOT EXISTS idx_students_major ON students(major)",
            "CREATE INDEX IF NOT EXISTS idx_students_class ON students(class_name)",
            "CREATE INDEX IF NOT EXISTS idx_students_id_card ON students(id_card)",
            "CREATE INDEX IF NOT EXISTS idx_students_student_id ON students(student_id)",
            "CREATE INDEX IF NOT EXISTS idx_textbooks_semester ON textbooks(semester_id)",
            "CREATE INDEX IF NOT EXISTS idx_textbooks_class ON textbooks(class_name)",
            "CREATE INDEX IF NOT EXISTS idx_textbooks_name ON textbooks(name)",
            "CREATE INDEX IF NOT EXISTS idx_exemptions_semester ON student_exemptions(semester_id)",
            "CREATE INDEX IF NOT EXISTS idx_exemptions_student ON student_exemptions(student_id)",
            "CREATE INDEX IF NOT EXISTS idx_distributions_student ON distributions(student_id)",
            "CREATE INDEX IF NOT EXISTS idx_distributions_textbook ON distributions(textbook_id)",
            "CREATE INDEX IF NOT EXISTS idx_signatures_student ON signatures(student_id)",
            "CREATE INDEX IF NOT EXISTS idx_signatures_semester ON signatures(semester_id)",
        ]
        for sql in indexes:
            try:
                cur.execute(sql)
            except Exception:
                pass
        conn.commit()

        # ── 自动迁移 ──
        _run_migrations(cur, conn)

        # ── V3.0 新增表 ──
        from database_v2 import init_v2_tables
        init_v2_tables(cur, conn)


def _run_migrations(cur, conn):
    """执行数据库迁移（添加缺失列、表）"""
    # 旧版 textbooks 表迁移
    try:
        cur.execute("PRAGMA table_info(textbooks)")
        tb_cols = {r[1] for r in cur.fetchall()}
        for col, col_type in [
            ("isbn", "VARCHAR(50)"),
            ("course_name", "VARCHAR(200)"),
            ("publication_date", "VARCHAR(50)"),
            ("editor", "VARCHAR(200)"),
        ]:
            if col not in tb_cols:
                cur.execute(f"ALTER TABLE textbooks ADD COLUMN {col} {col_type}")
                conn.commit()
                print(f"Migration: added {col} to textbooks")
    except Exception:
        pass

    # textbooks_master 迁移
    try:
        cur.execute("PRAGMA table_info(textbooks_master)")
        master_cols = {r[1] for r in cur.fetchall()}
        for col, col_type in [
            ("publication_date", "VARCHAR(50)"),
            ("discount_rate", "REAL DEFAULT NULL"),
            ("actual_price", "REAL DEFAULT NULL"),
            ("textbook_type", "TEXT DEFAULT NULL"),
        ]:
            if col not in master_cols:
                cur.execute(f"ALTER TABLE textbooks_master ADD COLUMN {col} {col_type}")
                conn.commit()
                print(f"Migration: added {col} to textbooks_master")
        # 旧字段 is_national_standard → textbook_type
        if "is_national_standard" in master_cols and "textbook_type" not in master_cols:
            cur.execute("ALTER TABLE textbooks_master ADD COLUMN textbook_type TEXT DEFAULT NULL")
            cur.execute("UPDATE textbooks_master SET textbook_type='国规教材' WHERE is_national_standard=1")
            conn.commit()
            print("Migration: migrated is_national_standard -> textbook_type")
    except Exception:
        pass

    # textbook_subscriptions 迁移
    try:
        cur.execute("PRAGMA table_info(textbook_subscriptions)")
        sub_cols = {r[1] for r in cur.fetchall()}
        if "class_names" not in sub_cols:
            cur.execute("ALTER TABLE textbook_subscriptions ADD COLUMN class_names TEXT")
            conn.commit()
            print("Migration: added class_names to textbook_subscriptions")
    except Exception:
        pass

    # distributions 表：删除 unit_price 列（改用实时计算）
    try:
        cur.execute("PRAGMA table_info(distributions)")
        dist_cols = {r[1] for r in cur.fetchall()}
        if "unit_price" in dist_cols:
            # SQLite 需重建表来删除列
            cur.execute("""
                CREATE TABLE distributions_new (
                    id INTEGER PRIMARY KEY AUTOINCREMENT,
                    student_id INTEGER NOT NULL,
                    textbook_id INTEGER NOT NULL,
                    quantity INTEGER NOT NULL DEFAULT 1,
                    distribute_date DATE,
                    handler VARCHAR(50),
                    created_at DATETIME DEFAULT CURRENT_TIMESTAMP,
                    FOREIGN KEY (student_id) REFERENCES students(id) ON DELETE CASCADE,
                    FOREIGN KEY (textbook_id) REFERENCES textbooks(id) ON DELETE CASCADE
                )
            """)
            cur.execute("""
                INSERT INTO distributions_new
                (id, student_id, textbook_id, quantity, distribute_date, handler, created_at)
                SELECT id, student_id, textbook_id, quantity, distribute_date, handler, created_at
                FROM distributions
            """)
            cur.execute("DROP TABLE distributions")
            cur.execute("ALTER TABLE distributions_new RENAME TO distributions")
            # 重建索引
            for idx_sql in [
                "CREATE INDEX IF NOT EXISTS idx_distributions_student ON distributions(student_id)",
                "CREATE INDEX IF NOT EXISTS idx_distributions_textbook ON distributions(textbook_id)",
            "CREATE INDEX IF NOT EXISTS idx_signatures_student ON signatures(student_id)",
            "CREATE INDEX IF NOT EXISTS idx_signatures_semester ON signatures(semester_id)",
            ]:
                cur.execute(idx_sql)
            conn.commit()
            print("Migration: dropped unit_price from distributions")
    except Exception:
        pass


def query_df(sql: str, params: tuple = None) -> pd.DataFrame:
    """
    执行 SELECT 查询，返回 DataFrame。

    Args:
        sql: SQL 语句（使用 %s 作为占位符）
        params: 参数元组

    Returns:
        pd.DataFrame，结果为空时返回空 DataFrame
    """
    if DB_TYPE == "sqlite":
        sql = sql.replace("%s", "?")
    with get_connection() as conn:
        return pd.read_sql(sql, conn, params=params)


def execute_sql(sql: str, params: tuple = None) -> int:
    """
    执行 INSERT/UPDATE/DELETE 操作。

    Args:
        sql: SQL 语句（使用 %s 作为占位符）
        params: 参数元组

    Returns:
        lastrowid（INSERT 时为新 ID，其他为 0）
    """
    if DB_TYPE == "sqlite":
        sql = sql.replace("%s", "?")
    with get_connection() as conn:
        cur = conn.cursor()
        cur.execute(sql, params or ())
        conn.commit()
        return cur.lastrowid
