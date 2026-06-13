"""
认证模块

提供登录页面、管理员/学生侧边栏导航
"""
import streamlit as st
import bcrypt
from config import ADMIN_PASSWORD_HASH
from database import query_df
from utils import get_current_academic_info
from components import show_header


def login_page() -> None:
    """登录页面：管理员密码登录 / 学生学号身份证查询"""
    st.markdown("""
    <div style="text-align:center; padding: 60px 0 20px 0;">
        <div style="font-size:64px; margin-bottom:12px;">📚</div>
        <h1 style="color:#1e40af; font-size:36px; font-weight:800; margin:0;">学生教材费用核对系统</h1>
        <p style="color:#6b7280; font-size:16px; margin-top:8px;">Textbook Fee Management System</p>
    </div>
    """, unsafe_allow_html=True)

    ay, sem = get_current_academic_info()
    st.markdown(f"""
    <div style="text-align:center; margin-bottom:32px;">
        <span style="background:#eff6ff; color:#1e40af; padding:6px 20px; border-radius:20px; font-weight:600;">
            📅 当前：{ay} 学年 {sem}
        </span>
    </div>
    """, unsafe_allow_html=True)

    col_l, col_c, col_r = st.columns([1, 1.5, 1])
    with col_c:
        tab1, tab2 = st.tabs(["🔑 管理员登录", "🎓 学生查询"])

        # ── 管理员登录 ──
        with tab1:
            with st.form("admin_login", clear_on_submit=False):
                st.markdown("### 管理员入口")
                pwd = st.text_input("管理员密码", type="password", placeholder="请输入管理员密码")
                submitted = st.form_submit_button("🔓 登录系统", use_container_width=True)
                if submitted:
                    if not ADMIN_PASSWORD_HASH:
                        st.error("⚠️ 未配置管理员密码（ADMIN_PASSWORD_HASH 环境变量）")
                    elif bcrypt.checkpw(pwd.encode(), ADMIN_PASSWORD_HASH.encode()):
                        st.session_state.role = "admin"
                        st.session_state.user = "管理员"
                        st.session_state.page = "semester"
                        st.rerun()
                    else:
                        st.error("❌ 密码错误，请重试")

        # ── 学生查询 ──
        with tab2:
            with st.form("student_login", clear_on_submit=False):
                st.markdown("### 学生查询入口")
                login_id = st.text_input("请输入学号或身份证号", placeholder="学号 或 身份证号")
                submitted = st.form_submit_button("🔍 查询费用", use_container_width=True)
                if submitted:
                    if login_id.strip():
                        df = query_df(
                            "SELECT * FROM students WHERE student_id = %s OR id_card = %s",
                            (login_id.strip(), login_id.strip())
                        )
                        if not df.empty:
                            student = df.iloc[0].to_dict()
                            st.session_state.role = "student"
                            st.session_state.user = student
                            st.session_state.page = "student_query"
                            st.rerun()
                        else:
                            st.error("❌ 未找到该学生信息，请检查学号或身份证号")
                    else:
                        st.warning("请输入学号或身份证号")

    st.markdown("""
    <div style="text-align:center; color:#9ca3af; font-size:12px; margin-top:48px;">
        v2.5 &nbsp;|&nbsp; 支持 MySQL / SQLite &nbsp;|&nbsp; 按班级分页打印
    </div>
    """, unsafe_allow_html=True)


def admin_sidebar() -> None:
    """管理员侧边栏：学期/学生/教材/征订/发放/确认/统计/日志导航"""
    with st.sidebar:
        st.markdown("""
        <div style="text-align:center; padding:16px 0 8px 0;">
            <div style="font-size:40px;">📚</div>
            <h3 style="color:#1e40af; margin:8px 0 0 0; font-weight:700;">教材费用系统</h3>
        </div>
        """, unsafe_allow_html=True)

        st.markdown('<div class="user-badge">👤 管理员模式</div>', unsafe_allow_html=True)

        ay, sem = get_current_academic_info()
        st.caption(f"🏫 {ay} {sem}")

        st.divider()

        pages = [
            ("📅", "学期管理", "semester", "管理学年学期"),
            ("👨‍🎓", "学生管理", "students", "导入/编辑学生信息"),
            ("📖", "教材表管理", "textbook_master", "管理教材库（名称/ISBN/出版社/主编/单价/课程）"),
            ("📋", "征订总表", "subscriptions", "导入原始征订数据，一键下发到班级"),
            ("📖", "教材征订表", "textbooks", "按班级查看/编辑征订明细"),
            ("📦", "教材发放表", "distribution", "录入发放记录"),
            ("✅", "领书确认表", "confirmation", "学生领书确认（免领标记）"),
            ("📢", "反馈处理", "feedback", "查看和处理学生教材核对反馈 [V2.0新增]"),
            ("📊", "费用统计", "statistics", "多维度费用汇总"),
            ("📋", "系统日志", "logs", "查看导入/操作记录"),
        ]

        for emoji, label, key, desc in pages:
            active = st.session_state.get("page") == key
            btn_type = "primary" if active else "secondary"
            if st.button(f"{emoji} {label}", use_container_width=True, type=btn_type,
                         help=desc, key=f"nav_{key}"):
                st.session_state.page = key
                st.rerun()

        st.divider()
        if st.button("🚪 退出登录", use_container_width=True):
            for k in ["role", "user", "page"]:
                st.session_state.pop(k, None)
            st.rerun()


def student_sidebar() -> None:
    """学生侧边栏：显示学生信息和导航"""
    from database_v2 import get_unread_count

    with st.sidebar:
        student = st.session_state.user
        st.markdown(f"""
        <div style="text-align:center; padding:16px 0 8px 0;">
            <div style="font-size:48px;">🎓</div>
            <h3 style="color:#1e40af; margin:8px 0 0 0;">学生端</h3>
        </div>
        """, unsafe_allow_html=True)

        st.info(f"**{student['name']}**\n学号：{student['student_id']}\n班级：{student.get('class_name', '-')}")

        st.divider()

        # 学生端页面导航
        unread = get_unread_count(int(student["id"]))
        student_pages = [
            ("📊", "费用查询", "student_query", "查看个人教材费用"),
            ("✅", "核对确认" + (f"({unread})" if unread > 0 else ""), "student_confirm", "核对教材信息并签名 [V2.0新增]"),
        ]

        for emoji, label, key, desc in student_pages:
            active = st.session_state.get("page") == key
            btn_type = "primary" if active else "secondary"
            if st.button(f"{emoji} {label}", use_container_width=True, type=btn_type,
                         help=desc, key=f"nav_{key}"):
                st.session_state.page = key
                st.rerun()

        st.divider()
        if st.button("🚪 退出登录", use_container_width=True):
            for k in ["role", "user", "page"]:
                st.session_state.pop(k, None)
            st.rerun()
