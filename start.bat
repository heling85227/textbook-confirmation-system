@echo off
REM 学生教材费用核对系统 — 启动脚本
REM 启动方式: 双击此文件或命令行执行
set ADMIN_PASSWORD_HASH=$2b$12$Tt41dNlGAIe8dGGt5ybUGu2OALT7E26IaBpIiQJtybKrtnNL5wv62
C:\Users\HL\.workbuddy\binaries\python\versions\3.13.12\python.exe -m streamlit run main.py --server.port 8501 --server.headless true
