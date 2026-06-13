"""
工具函数模块

提供：日期处理、安全类型转换、Excel 读写、过滤查询、导入日志等功能
"""
import io
import os
import re
import pandas as pd
from datetime import datetime
from database import query_df

# ── 常量 ──
MAX_UPLOAD_MB = 5
LOG_DIR = os.path.join(os.path.dirname(os.path.abspath(__file__)), "logs")


# ═══════════════════════════════════════════════════════
# 日期/学期
# ═══════════════════════════════════════════════════════

def get_current_academic_info() -> tuple:
    """自动识别当前学年和学期，返回 (学年字符串, 学期名)"""
    now = datetime.now()
    year, month = now.year, now.month
    if month >= 9:
        return f"{year}-{year+1}", "第一学期"
    elif month <= 2:
        return f"{year-1}-{year}", "第一学期"
    else:
        return f"{year-1}-{year}", "第二学期"


# ═══════════════════════════════════════════════════════
# 安全类型转换
# ═══════════════════════════════════════════════════════

def safe_int(v, default=0) -> int:
    """安全转换为 int，失败返回 default"""
    try:
        return int(float(v))
    except (ValueError, TypeError):
        return default


def safe_float(v, default=0.0) -> float:
    """安全转换为 float，失败返回 default"""
    try:
        return float(v)
    except (ValueError, TypeError):
        return default


def safe_str(v, default="") -> str:
    """
    安全转换为字符串。
    - NaN/None → default
    - datetime 对象 → 仅保留日期部分 (YYYY-MM-DD)
    - 自动去除 " 00:00:00" 后缀
    """
    if v is None or (isinstance(v, float) and pd.isna(v)):
        return default
    if hasattr(v, "strftime"):
        return v.strftime("%Y-%m-%d")
    s = str(v).strip()
    if " 00:00:00" in s:
        s = s.replace(" 00:00:00", "")
    return s


def safe_field(v, default="") -> str:
    """
    将字段值转为字符串，处理数字（如 2023.0→2023）、NaN、None。
    用于 Excel 导入时清洗字段值。
    """
    if v is None or (isinstance(v, float) and pd.isna(v)):
        return default
    if isinstance(v, float) and v == int(v):
        return str(int(v))
    return str(v).strip()


# ═══════════════════════════════════════════════════════
# 过滤/查询
# ═══════════════════════════════════════════════════════

def get_filtered_list(table: str, column: str, where: str = "1=1", params: tuple = ()) -> list:
    """
    通用不重复值查询。

    Args:
        table: 表名
        column: 列名
        where: WHERE 条件子句（不含 "WHERE" 关键字）
        params: 条件参数

    Returns:
        不重复值列表（已排序）
    """
    df = query_df(
        f"SELECT DISTINCT {column} FROM {table} "
        f"WHERE {column} IS NOT NULL AND {column} != '' AND {where} "
        f"ORDER BY {column}",
        params
    )
    return df[column].tolist() if not df.empty else []


def get_filtered_class_names() -> list:
    return get_filtered_list("students", "class_name")


def get_filtered_grades() -> list:
    return get_filtered_list("students", "grade")


def get_filtered_majors() -> list:
    return get_filtered_list("students", "major")


def get_filtered_colleges() -> list:
    return get_filtered_list("students", "college")


def normalize_grade(g: str) -> str:
    """
    统一年级格式为 "202X级"。
    兼容 "202X"（无"级"后缀）和 "202X级" 两种写法。
    """
    if not g:
        return g
    g = str(g).strip()
    if not g.endswith("级"):
        g = g + "级"
    return g


def get_class_student_counts(grade=None, college=None, major=None, class_names=None) -> pd.DataFrame:
    """
    查询指定条件下各班级的实际学生人数分布。

    Args:
        grade: 年级（如 "2024级"）
        college: 学院
        major: 专业
        class_names: 班级名列表

    Returns:
        DataFrame，含 class_name, student_count 列
    """
    where = "1=1"
    params = []
    if grade:
        g_raw = str(grade).rstrip("级")
        where += " AND (grade = %s OR grade = %s)"
        params.extend([g_raw + "级", g_raw])
    if college:
        where += " AND college = %s"
        params.append(college)
    if major:
        where += " AND major = %s"
        params.append(major)
    if class_names:
        ph = ",".join(["%s"] * len(class_names))
        where += f" AND class_name IN ({ph})"
        params.extend(class_names)
    return query_df(
        f"SELECT class_name, COUNT(*) as student_count FROM students WHERE {where} GROUP BY class_name ORDER BY class_name",
        tuple(params)
    )


def split_qty_by_class(total_qty: int, grade=None, college=None, major=None, class_names=None) -> list:
    """
    【已废弃】将总数量按各班级实际人数比例分摊到每个班。
    现改为按人数直接下发（每人1本）。

    返回: [(class_name, student_count, allocated_qty), ...]
    如无法获取班级分布，返回 [(None, 0, total_qty)]
    """
    df = get_class_student_counts(grade, college, major, class_names)
    if df.empty:
        return [(None, 0, total_qty)]
    total_students = df["student_count"].sum()
    result = []
    for _, row in df.iterrows():
        cn = row["class_name"]
        cnt = int(row["student_count"])
        alloc = round(total_qty * cnt / total_students)
        result.append((cn, cnt, alloc))
    # 四舍五入校正
    diff = total_qty - sum(r[2] for r in result)
    if diff != 0 and result:
        max_idx = max(range(len(result)), key=lambda i: result[i][2])
        result[max_idx] = (result[max_idx][0], result[max_idx][1], result[max_idx][2] + diff)
    return result


# ═══════════════════════════════════════════════════════
# 专业/年级解析（征订总表用）
# ═══════════════════════════════════════════════════════

def parse_major_grade_from_scope(class_scope: str, college: str = "") -> tuple:
    """
    从班级范围说明中智能解析专业和年级。

    支持的格式:
        - "会计1241-1244" → 从班级号解析年级（1241→24→2024级）
        - "机械25级"     → major="机械", grades=["2025级"]
        - "25市场营销级"  → major="市场营销", grades=["2025级"]
        - "24级、25级学生" → grades=["2024级","2025级"]
        - "教师用书"      → ("", [])

    Returns: (major_full, [grades]) — grades 始终为列表
    """
    if not class_scope:
        return "", []
    scope = str(class_scope).strip()
    if "教师" in scope:
        return "", []

    major_short = ""
    grades = []

    # 优先级1：从「专业名+班级号」模式提取年级
    # 如 "会计1241-1244", "风电1251-1255", "会计1241-会计1251"（跨年级）
    class_num_m = re.findall(r'(\d{4})', scope)
    if class_num_m:
        years_from_class = set()
        for cn in class_num_m:
            if len(cn) == 4:
                yy = cn[1:3]  # 第2-3位是年份
                if yy.isdigit():
                    years_from_class.add(int(yy))
        if years_from_class:
            grades = sorted([f"20{yy}级" for yy in years_from_class])
        mm = re.match(r'^([\u4e00-\u9fff]+)', scope)
        if mm:
            major_short = mm.group(1)
        return _resolve_major_name(major_short, college), grades

    # 优先级2: 匹配「XX级」模式（支持多个"XX级"）
    yy_matches = re.findall(r'(\d{2})\s*级', scope)
    if yy_matches:
        grades = sorted([f"20{yy}级" for yy in yy_matches])
        remaining = re.sub(r'\d{2}\s*级', '', scope).strip()
        mm = re.match(r'^([\u4e00-\u9fff]+)', remaining)
        if mm:
            major_short = mm.group(1)
        return _resolve_major_name(major_short, college), grades

    # 优先级3: 匹配「开头数字+中文+级」(如 25市场营销级)
    m = re.match(r'^(\d{2})\s*([\u4e00-\u9fff]+)\s*级', scope)
    if m:
        grades = ["20" + m.group(1) + "级"]
        major_short = m.group(2)
        return _resolve_major_name(major_short, college), grades

    # 兜底: 提取开头中文，年级从学生表查
    mm = re.match(r'^([\u4e00-\u9fff]+)', scope)
    if mm:
        major_short = mm.group(1)
    return _resolve_major_name(major_short, college), grades


def _resolve_major_name(major_short: str, college: str = "") -> str:
    """
    通过学生表将短专业名解析为完整专业名。
    先用 major LIKE 匹配，再通过 class_name 反查。
    """
    if not major_short:
        return ""

    conditions = ["major LIKE %s"]
    params = [f"%{major_short}%"]
    if college:
        conditions.append("college = %s")
        params.append(college)

    where = " AND ".join(conditions)
    students_df = query_df(
        f"SELECT DISTINCT major FROM students WHERE {where} LIMIT 3",
        tuple(params)
    )
    if not students_df.empty:
        return str(students_df.iloc[0]["major"])

    # 通过班级名反查
    conds2 = ["class_name LIKE %s"]
    params2 = [f"{major_short}%"]
    if college:
        conds2.append("college = %s")
        params2.append(college)
    students_df = query_df(
        f"SELECT DISTINCT major FROM students WHERE {' AND '.join(conds2)} LIMIT 3",
        tuple(params2)
    )
    if not students_df.empty:
        return str(students_df.iloc[0]["major"])

    return major_short


# ═══════════════════════════════════════════════════════
# Excel 读写
# ═══════════════════════════════════════════════════════

def read_excel_upload(uploaded_file) -> pd.DataFrame:
    """
    读取上传的 Excel 文件，带安全校验。
    - 文件大小 ≤ 5MB
    - 非空
    - 行数 ≤ 5000
    """
    size_mb = len(uploaded_file.getvalue()) / (1024 * 1024)
    if size_mb > MAX_UPLOAD_MB:
        raise ValueError(f"文件过大（{size_mb:.1f}MB），上限 {MAX_UPLOAD_MB}MB")
    uploaded_file.seek(0)
    df = pd.read_excel(uploaded_file, engine="openpyxl")
    if df.empty:
        raise ValueError("文件为空或无有效数据")
    if len(df) > 5000:
        raise ValueError(f"行数过多（{len(df)}），上限 5000 行")
    return df


def make_template_df(columns: list) -> pd.DataFrame:
    """生成空模板 DataFrame（只有表头，无数据行）"""
    return pd.DataFrame(columns=columns)


# ═══════════════════════════════════════════════════════
# 导入日志
# ═══════════════════════════════════════════════════════

def write_import_log(module: str, filename: str, total: int, success: int, errors: list) -> None:
    """记录导入操作到日志文件"""
    os.makedirs(LOG_DIR, exist_ok=True)
    log_file = os.path.join(LOG_DIR, "import.log")
    now = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    error_detail = "; ".join(errors[:20]) if errors else "无"
    line = (
        f"[{now}]\n"
        f"  模块: {module}\n"
        f"  文件: {filename}\n"
        f"  总计: {total} 条 | 成功: {success} 条 | 失败: {len(errors)} 条\n"
        f"  错误: {error_detail}\n"
        f"  {'─' * 60}\n"
    )
    with open(log_file, "a", encoding="utf-8") as f:
        f.write(line)


def read_import_logs(n: int = 50) -> list:
    """读取最近的导入日志，返回字符串列表"""
    log_file = os.path.join(LOG_DIR, "import.log")
    if not os.path.exists(log_file):
        return ["暂无日志记录"]
    with open(log_file, "r", encoding="utf-8") as f:
        lines = f.readlines()
    blocks = []
    block = []
    for line in reversed(lines):
        if line.startswith("["):
            if block:
                blocks.append("".join(reversed(block)))
                block = []
        block.append(line)
    if block:
        blocks.append("".join(reversed(block)))
    return blocks[:n]
