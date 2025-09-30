@echo off
setlocal
cd /d %~dp0

:: Подгружаем .env (простенький парсер)
for /f "usebackq tokens=1,2 delims==" %%A in (".env") do (
  if not "%%A"=="" set "%%A=%%B"
)

:: Активация venv (если другой — поправь путь)
if exist venv311\Scripts\activate.bat (
  call venv311\Scripts\activate.bat
)

:: Запуск
python school_autho_egor.py

endlocal
