#!/bin/bash
# 安裝 Node.js 套件（docx）
npm install docx
# 啟動 Flask
gunicorn main:app --bind 0.0.0.0:$PORT
