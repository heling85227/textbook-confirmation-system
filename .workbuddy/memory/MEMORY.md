# 项目记忆

## 项目概述
- 学生教材交互确认系统，基于 Streamlit + Python 构建
- 双数据库引擎：SQLite（开发）/ MySQL（生产）
- 版本：v2.5 基础版 + V2.0 交互确认增强（核对确认工作流）

## 系统架构
- 前端：Streamlit >= 1.28
- 后端：Python 3.13
- 数据处理：pandas + openpyxl
- 密码：bcrypt
- 签名：streamlit-drawable-canvas + Pillow
- 配置：config.ini + .env

## 核心数据表（8张）
1. semesters - 学期表
2. students - 学生表
3. textbooks_master - 教材主表
4. textbook_orders - 教材征订明细表
5. textbook_subscriptions - 征订总表
6. distributions - 发放记录表
7. student_exemptions - 免领标记表
8. signatures - 签名表

## V2.0 新增数据表（2张）
1. student_confirmations - 学生核对确认表
2. student_notifications - 学生通知表

## 价格计算规则
PRICE_CALC: 实洋价 > 定价×折扣率 > 定价 > 0

## 文件结构
- main.py: 主入口+路由（原v2.5 + V2.0路由）
- auth.py: 认证模块（管理员密码+学生学号/身份证登录）+ V2.0导航
- config.py: 全局配置+CSS样式
- database.py: 数据库层（8张原表CRUD + 调用V2.0建表）
- database_v2.py: ⭐ V2.0新增 — 2张表CRUD + 反馈处理 + 批量操作
- components.py: 共享UI组件
- utils.py: 工具函数
- pages/: 各功能页面
  - 管理端（原）：semester/students/textbooks/subscriptions/textbook_master/distribution/confirmation/statistics/logs
  - 管理端（V2.0新增）：feedback_v2
  - 学生端（原）：student_query（费用+领书+签名）
  - 学生端（V2.0新增）：student_confirm_v2（核对确认工作流）

## 原代码来源
- 原代码在 D:\2026-05-30-22-19-46 (1)\，含 textbook_data.db（2341名学生、152条教材等）
- 已备份到 D:\教材管理系统code\backups\v2.5_baseline\

## V2.0 代码实现状态（2026-06-13）
- 新文件（按版本命名）：database_v2.py, pages/feedback_v2.py, pages/student_confirm_v2.py
- 修改的原文件（最小改动）：database.py(+2行)、main.py(+4行)、auth.py(侧边栏扩展)
- ✅ 核对确认工作流（pending→confirmed/disputed→pending状态机）
- ✅ 8种反馈类型（退书/补领/少领/分配错误/重复发放/价格疑问/不需要/其他）
- ✅ 通知系统（创建/已读/未读计数/批量通知）
- ✅ 反馈处理（接受/驳回+自动通知）
- ✅ 签名板+签名保存/读取（复用原有signatures表）
- ✅ 真实数据CRUD测试通过

## 依赖包
streamlit>=1.28, pandas, openpyxl, python-dotenv, bcrypt, streamlit-drawable-canvas, Pillow, pymysql

## 启动方式
cd D:/教材管理系统code && D:/python/python.exe -m streamlit run main.py
