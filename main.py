"""
学生教材交互确认系统 V3.0 — 主入口
====================================

多模块架构：
    main.py             入口 + 路由
    config.py           全局配置 / CSS
    database.py         数据库连接 / 建表 / 查询
    utils.py            工具函数（日期 / 安全转换 / 过滤 / Excel）
    components.py       共享 UI 组件（页头 / Excel 导出）
    auth.py             登录 / 侧边栏
    test_data.py        测试数据生成
    pages/              各功能页面模块

启动方式：
    streamlit run main.py
"""
import streamlit as st
from config import get_custom_css
from database import init_db
from auth import login_page, admin_sidebar, student_sidebar
from pages.semester import semester_management
from pages.students import student_management
from pages.subscriptions import subscription_management
from pages.textbook_master import textbook_master_management
from pages.textbooks import textbook_management
from pages.distribution import distribution_management
from pages.confirmation import confirmation_page
from pages.statistics import statistics_page
from pages.student_query import student_query_page
from pages.logs import system_logs_page
from pages.feedback_v2 import feedback_page
from pages.student_confirm_v2 import student_confirm_page


# ── 页面配置 ──
st.set_page_config(
    page_title="学生教材费用核对系统",
    page_icon="📚",
    layout="wide",
    initial_sidebar_state="expanded"
)

# ── 注入全局 CSS ──
st.markdown(get_custom_css(), unsafe_allow_html=True)

# ── 通过 components.html 注入 STYLE 到 parent document（唯一能穿透 Streamlit 表格的方式）──
# CSS 直接加在 parent.head 里，对 st.data_editor 和 st.dataframe 的网格单元格生效
st.components.v1.html("""
<script>
(function() {
    try {
        var pdoc = window.parent.document;
        var style = pdoc.createElement('style');
        style.id = 'wb-center-fix';
        style.textContent = [
            '[data-testid="stDataFrame"] [role="gridcell"] { text-align:center !important }',
            '[data-testid="stDataEditor"] [role="gridcell"] { text-align:center !important }',
            '[data-testid="stDataFrame"] [role="columnheader"] { text-align:center !important }',
            '[data-testid="stDataEditor"] [role="columnheader"] { text-align:center !important }',
            '[data-testid="stDataFrame"] .dvn-scroller [class*="cell"] { text-align:center !important; justify-content:center !important }',
            '[data-testid="stDataEditor"] .dvn-scroller [class*="cell"] { text-align:center !important; justify-content:center !important }',
            '[data-testid="stDataFrame"] div[class*="gdg-"] { text-align:center !important }',
            '[data-testid="stDataEditor"] div[class*="gdg-"] { text-align:center !important }',
            // Fix: Streamlit tabs collapse iframe height to 0 — force st_canvas iframe to have visible height
            '.stIFrame { min-height: 320px !important; height: auto !important; }',
            'iframe[src*=\"streamlit_drawable_canvas\"], iframe[src*=\"st_canvas\"] { min-height: 320px !important; height: 320px !important; display: block !important; }',
        ].join('\\n');
        var old = pdoc.getElementById('wb-center-fix');
        if (old) old.remove();
        pdoc.head.appendChild(style);
        console.log('[WB] Center style injected into parent document');
        // Also fix inline styles on st_canvas iframes directly
        setTimeout(function(){
            var ifs = pdoc.querySelectorAll('iframe[src*=\"streamlit_drawable_canvas\"], iframe[src*=\"st_canvas\"]');
            for (var i=0; i<ifs.length; i++) {
                ifs[i].style.setProperty('height', '320px', 'important');
                ifs[i].style.setProperty('min-height', '320px', 'important');
                ifs[i].style.setProperty('display', 'block', 'important');
            }
        }, 1000);
        // Retry since Streamlit might reset the height
        setInterval(function(){
            var ifs = pdoc.querySelectorAll('iframe[src*=\"streamlit_drawable_canvas\"], iframe[src*=\"st_canvas\"]');
            for (var i=0; i<ifs.length; i++) {
                if (ifs[i].offsetHeight < 100) {
                    ifs[i].style.setProperty('height', '320px', 'important');
                    ifs[i].style.setProperty('min-height', '320px', 'important');
                }
            }
        }, 2000);
    } catch(e) {
        console.error('[WB] Style injection failed:', e);
    }
})();
</script>
""", height=1)


def main():
    """应用主入口：初始化会话 → 连接数据库 → 路由分发"""

    # ── 初始化 session state ──
    defaults = {
        "role": None,           # 用户角色：None / "admin" / "student"
        "page": None,           # 当前页面标识
        "show_semester_form": False,
        "show_gen_confirm": False,
        "pending_action": "",
    }
    for key, val in defaults.items():
        if key not in st.session_state:
            st.session_state[key] = val

    # ── 初始化数据库（建表 + 迁移）──
    try:
        init_db()
    except Exception as e:
        st.error(f"❌ 数据库初始化失败：{e}")
        st.stop()

    # ── 空库/缺数据时自动填充演示数据（Streamlit Cloud 等） ──
    try:
        from demo_data import needs_demo_data, init_demo_data
        if needs_demo_data():
            result = init_demo_data()
            if not result.get("skipped"):
                st.session_state.demo_initialized = True
            # 显示初始化结果（方便确认）
            if "demo_shown" not in st.session_state:
                summary = {k: v for k, v in result.items() if isinstance(v, int)}
                if any(v > 0 for v in summary.values()):
                    st.toast(f"\U0001f4e6 演示数据已加载: 学期={summary.get('semesters',0)}, "
                            f"学生={summary.get('students',0)}, 教材主表={summary.get('master_books',0)}", icon="\u2705")
                    st.session_state.demo_shown = True
    except Exception as e:
        import traceback
        st.error(f"\u274c 演示数据初始化失败：{e}")
        st.code(traceback.format_exc())

    # ── 路由分发 ──
    if st.session_state.role is None:
        # 未登录 → 显示登录页
        login_page()

    elif st.session_state.role == "admin":
        # 管理员 → 侧边栏 + 页面路由
        admin_sidebar()
        page_map = {
            "semester": semester_management,
            "students": student_management,
            "textbook_master": textbook_master_management,
            "subscriptions": subscription_management,
            "textbooks": textbook_management,
            "distribution": distribution_management,
            "confirmation": confirmation_page,
            "statistics": statistics_page,
            "logs": system_logs_page,
            "feedback": feedback_page,
        }
        current_page = st.session_state.get("page", "semester")
        page_func = page_map.get(current_page, semester_management)
        page_func()

    elif st.session_state.role == "student":
        # 学生 → 侧边栏 + 查询页
        student_sidebar()
        if st.session_state.get("page") == "student_confirm":
            student_confirm_page()
        else:
            student_query_page()


if __name__ == "__main__":
    main()
