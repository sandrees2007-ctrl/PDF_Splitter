@echo off
chcp 65001 > nul
echo Запуск веб-сервера StroiLab Intelligence Suite...
start http://localhost:5000
py -3.14 app.py
pause
