#!/bin/bash

# 日志文件路径
LOG_FILE="/home/toolsDev/log_file.log"

# 记录脚本开始时间
echo "Script started at $(date)" >> $LOG_FILE

# 切换到指定目录
cd /home/toolsDev/ >> $LOG_FILE 2>&1

# 激活 Python 虚拟环境
source venv/bin/activate >> $LOG_FILE 2>&1

# 杀掉占用 5002 端口的进程
kill -9 $(lsof -t -i :5002) >> $LOG_FILE 2>&1

# 使用 Gunicorn 启动应用
gunicorn -c gunicorn.py app:app & >> $LOG_FILE 2>&1

# 记录脚本结束时间
echo "Script ended at $(date)" >> $LOG_FILE

