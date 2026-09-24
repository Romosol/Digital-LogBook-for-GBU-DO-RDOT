@echo off
chcp 65001 > nul
echo ====================================================
echo  Запуск графического интерфейса классного журнала
echo ====================================================
echo Проверка библиотеки openpyxl...
python -m pip install -q openpyxl
echo Запуск приложения...
python app_gui.py
pause
