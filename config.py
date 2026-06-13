"""
配置模块
- 加载 config.ini 数据库配置
- 管理员密码
- 全局 CSS 样式
"""
import configparser
import os


def load_config() -> configparser.ConfigParser:
    """从 config.ini 加载配置（兼容 Streamlit Cloud 无文件环境）"""
    config = configparser.ConfigParser()
    config_path = os.path.join(os.path.dirname(__file__), "config.ini")
    if os.path.exists(config_path):
        config.read(config_path, encoding="utf-8")
    else:
        # Streamlit Cloud 环境或无配置文件：使用默认值 + 环境变量
        config["database"] = {
            "type": os.environ.get("DB_TYPE", "sqlite"),
            "host": os.environ.get("DB_HOST", "localhost"),
            "port": os.environ.get("DB_PORT", "3306"),
            "user": os.environ.get("DB_USER", "root"),
            "password": "",
            "database": os.environ.get("DB_NAME", "textbook_fee"),
        }
        config["admin"] = {"password": "admin123"}
    return config


# ── 数据库配置 ──
CONFIG = load_config()
DB_TYPE = CONFIG.get("database", "type", fallback="mysql")

DB_CONFIG = {
    "host": os.environ.get("DB_HOST", CONFIG.get("database", "host", fallback="localhost")),
    "port": int(os.environ.get("DB_PORT", CONFIG.get("database", "port", fallback="3306"))),
    "user": os.environ.get("DB_USER", CONFIG.get("database", "user", fallback="root")),
    "password": os.environ.get("DB_PASSWORD", CONFIG.get("database", "password", fallback="")),
    "database": os.environ.get("DB_NAME", CONFIG.get("database", "database", fallback="textbook_fee")),
    "charset": "utf8mb4",
}

# ── 管理员密码（默认 admin123）──
ADMIN_PASSWORD_HASH = os.environ.get(
    "ADMIN_PASSWORD_HASH",
    "$2b$12$Tt41dNlGAIe8dGGt5ybUGu2OALT7E26IaBpIiQJtybKrtnNL5wv62"
)


def get_custom_css() -> str:
    """返回自定义 CSS 样式（供 main.py 注入）"""
    return """
<style>
    @import url('https://fonts.googleapis.com/css2?family=Inter:wght@400;500;600;700&display=swap');

    html, body, [class*="css"] {
        font-family: 'Inter', -apple-system, 'Microsoft YaHei', sans-serif;
    }

    :root {
        --primary: #1e40af;
        --primary-light: #3b82f6;
        --primary-bg: #eff6ff;
        --success: #059669;
        --success-bg: #ecfdf5;
        --warning: #d97706;
        --warning-bg: #fffbeb;
        --danger: #dc2626;
        --danger-bg: #fef2f2;
        --gray-50: #f9fafb;
        --gray-100: #f3f4f6;
        --gray-200: #e5e7eb;
        --gray-600: #4b5563;
        --gray-700: #374151;
        --gray-800: #1f2937;
        --radius: 12px;
        --shadow: 0 1px 3px rgba(0,0,0,0.08), 0 1px 2px rgba(0,0,0,0.06);
        --shadow-md: 0 4px 12px rgba(0,0,0,0.08);
    }

    .app-header {
        background: linear-gradient(135deg, #1e40af 0%, #3b82f6 50%, #6366f1 100%);
        color: white;
        padding: 24px 32px;
        border-radius: var(--radius);
        margin-bottom: 24px;
        box-shadow: 0 4px 20px rgba(30,64,175,0.2);
    }
    .app-header h1 {
        color: white !important;
        font-size: 28px;
        font-weight: 700;
        margin: 0;
        padding: 0;
    }
    .app-header p {
        color: rgba(255,255,255,0.85);
        margin: 4px 0 0 0;
        font-size: 14px;
    }

    .info-card {
        background: white;
        border: 1px solid var(--gray-200);
        border-radius: var(--radius);
        padding: 20px 24px;
        margin-bottom: 16px;
        box-shadow: var(--shadow);
        transition: all 0.2s;
    }
    .info-card:hover {
        box-shadow: var(--shadow-md);
        transform: translateY(-1px);
    }
    .info-card h3 {
        font-size: 16px;
        font-weight: 600;
        color: var(--gray-800);
        margin: 0 0 4px 0;
    }
    .info-card p {
        font-size: 13px;
        color: var(--gray-600);
        margin: 0;
    }

    .student-card {
        background: linear-gradient(135deg, #eff6ff 0%, #f0fdf4 100%);
        border: 1px solid #bfdbfe;
        border-radius: var(--radius);
        padding: 20px;
        margin-bottom: 12px;
    }

    .stat-number {
        font-size: 32px;
        font-weight: 700;
        color: var(--primary);
        line-height: 1.2;
    }
    .stat-label {
        font-size: 13px;
        color: var(--gray-600);
        text-transform: uppercase;
        letter-spacing: 0.5px;
    }

    [data-testid="stMetricValue"] {
        font-size: 28px !important;
        font-weight: 700 !important;
    }
    [data-testid="stMetricLabel"] {
        font-size: 13px !important;
        color: var(--gray-600) !important;
    }
    [data-testid="stMetricDelta"] {
        font-size: 14px !important;
    }

    .stButton > button {
        border-radius: 8px !important;
        font-weight: 500 !important;
        transition: all 0.2s !important;
        border: none !important;
    }
    .stButton > button:hover {
        transform: translateY(-1px);
        box-shadow: 0 4px 12px rgba(0,0,0,0.15) !important;
    }
    .stButton > button:active {
        transform: translateY(0);
    }

    [data-testid="stSidebar"] {
        background: linear-gradient(180deg, #f8fafc 0%, #eff6ff 100%);
        border-right: 1px solid var(--gray-200);
    }
    [data-testid="stSidebar"] h1 {
        font-size: 22px !important;
        color: var(--primary) !important;
    }

    .stTabs [data-baseweb="tab-list"] {
        gap: 8px;
        background: var(--gray-50);
        padding: 6px;
        border-radius: 10px;
    }
    .stTabs [data-baseweb="tab"] {
        border-radius: 8px !important;
        padding: 8px 16px !important;
        font-weight: 500 !important;
        transition: all 0.2s;
    }
    .stTabs [aria-selected="true"] {
        background: white !important;
        box-shadow: var(--shadow);
    }

    div[data-testid="stForm"] {
        border: 1px solid var(--gray-200);
        border-radius: var(--radius);
        padding: 24px;
        background: white;
        box-shadow: var(--shadow);
    }

    [data-testid="stDataFrame"],
    [data-testid="stDataEditor"] {
        border-radius: var(--radius) !important;
        overflow: hidden;
        border: 1px solid var(--gray-200) !important;
    }
    /* 表头 */
    [data-testid="stDataFrame"] th,
    [data-testid="stDataEditor"] th,
    [data-testid="stDataFrame"] [role="columnheader"],
    [data-testid="stDataEditor"] [role="columnheader"] {
        background: var(--primary-bg) !important;
        color: var(--primary) !important;
        font-weight: 600 !important;
        font-size: 13px !important;
        white-space: nowrap !important;
        text-align: center !important;
    }
    /* 单元格 */
    [data-testid="stDataFrame"] td,
    [data-testid="stDataEditor"] td,
    [data-testid="stDataFrame"] [role="gridcell"],
    [data-testid="stDataEditor"] [role="gridcell"] {
        font-size: 13px !important;
        text-align: center !important;
    }
    /* 单元格内部 div/span 强制居中 */
    [data-testid="stDataFrame"] td > div,
    [data-testid="stDataEditor"] td > div,
    [data-testid="stDataFrame"] td > span,
    [data-testid="stDataEditor"] td > span,
    [data-testid="stDataFrame"] th > div,
    [data-testid="stDataEditor"] th > div,
    [data-testid="stDataFrame"] th > span,
    [data-testid="stDataEditor"] th > span {
        text-align: center !important;
        justify-content: center !important;
    }
    /* Streamlit 新版 data_editor 内部可能用 class 而非 td/th */
    [data-testid="stDataFrame"] [class*="cell"] > div,
    [data-testid="stDataEditor"] [class*="cell"] > div {
        text-align: center !important;
        justify-content: center !important;
    }

    .streamlit-expanderHeader {
        background: var(--gray-50);
        border-radius: 8px !important;
        font-weight: 600 !important;
    }

    hr {
        border-color: var(--gray-200) !important;
        margin: 24px 0 !important;
    }

    input, textarea, .stSelectbox > div > div {
        border-radius: 8px !important;
    }

    [data-testid="stFileUploader"] {
        border-radius: var(--radius) !important;
        border: 2px dashed var(--gray-200) !important;
        padding: 20px !important;
        transition: all 0.2s;
    }
    [data-testid="stFileUploader"]:hover {
        border-color: var(--primary-light) !important;
        background: var(--primary-bg) !important;
    }

    .element-container:has([data-testid="stSuccess"]) {
        animation: slideIn 0.3s ease;
    }
    @keyframes slideIn {
        from { opacity: 0; transform: translateY(-10px); }
        to { opacity: 1; transform: translateY(0); }
    }

    @media print {
        [data-testid="stSidebar"],
        .stButton,
        .stDownloadButton,
        [data-testid="stFileUploader"],
        header,
        .stTabs [data-baseweb="tab-list"],
        .st-emotion-cache-ocsh0s {
            display: none !important;
        }
        .stTabs [role="tabpanel"] {
            display: block !important;
        }
        [data-testid="stVerticalBlock"] {
            gap: 0 !important;
        }
        body {
            font-size: 12px;
        }
        table {
            font-size: 11px;
        }
    }

    .user-badge {
        display: inline-flex;
        align-items: center;
        gap: 8px;
        background: var(--primary-bg);
        color: var(--primary);
        padding: 8px 16px;
        border-radius: 20px;
        font-weight: 600;
        font-size: 14px;
        margin-bottom: 16px;
    }

    .class-tag {
        display: inline-block;
        background: var(--primary-bg);
        color: var(--primary);
        padding: 4px 12px;
        border-radius: 999px;
        font-size: 13px;
        font-weight: 500;
        margin: 2px;
    }

    .stMultiSelect [data-baseweb="tag"] {
        background: var(--primary-bg) !important;
        color: var(--primary) !important;
    }

    .template-btn {
        background: #f0fdf4 !important;
        color: #059669 !important;
        border: 1px solid #bbf7d0 !important;
    }

    [data-testid="stDataEditor"] input[type="checkbox"],
    [data-testid="stDataFrame"] input[type="checkbox"] {
        accent-color: #1677ff !important;
        width: 18px !important;
        height: 18px !important;
        opacity: 1 !important;
        visibility: visible !important;
    }
    [data-testid="stDataEditor"] input[type="checkbox"]:checked,
    [data-testid="stDataFrame"] input[type="checkbox"]:checked {
        accent-color: #1677ff !important;
    }
    [data-testid="stDataEditor"] td:has(input[type="checkbox"]),
    [data-testid="stDataFrame"] td:has(input[type="checkbox"]) {
        text-align: center !important;
    }

    button[kind="secondary"] {
        background-color: #ef4444 !important;
        color: #fff !important;
        border: 1px solid #dc2626 !important;
    }
    button[kind="secondary"]:hover:not(:disabled) {
        background-color: #dc2626 !important;
        border-color: #b91c1c !important;
    }
    button[kind="secondary"]:disabled {
        background-color: #e5e7eb !important;
        color: #9ca3af !important;
        border-color: #d1d5db !important;
    }
    [data-testid="stSidebar"] button[kind="secondary"] {
        background-color: inherit !important;
        color: inherit !important;
        border: inherit !important;
    }

    .pagination-bar {
        display: flex;
        align-items: center;
        justify-content: flex-end;
        gap: 12px;
        padding: 10px 0;
        font-size: 14px;
        color: #4b5563;
    }
    .pagination-bar .page-info {
        font-weight: 500;
        min-width: 80px;
        text-align: center;
    }
    .pagination-bar .total-info {
        color: #6b7280;
        font-size: 13px;
    }
    .pagination-bar button {
        min-width: 32px !important;
        height: 32px !important;
        padding: 0 8px !important;
    }
    
    /* Fix: Streamlit tabs collapse iframe height to 0 */
    iframe[src*="streamlit_drawable_canvas"],
    iframe[src*="st_canvas"] {
        min-height: 320px !important;
        height: 320px !important;
        display: block !important;
    }
</style>
"""
