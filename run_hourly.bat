@echo off
"C:\Users\Zen\PyCharmMiscProject\.venv\Scripts\python.exe" ^
 "D:\Investments\update_portfolio_unified.py" ^
 "D:\Investments\Portfolio.xlsx"

start "" "excel.exe" "D:\Investments\Portfolio.xlsx"