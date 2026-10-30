@echo off
chcp 65001 >nul
cd /d "%~dp0"
python -m pip install -q -r requirements.txt
if "%~1"=="" (
  python excel2word.py
) else (
  python excel2word.py %*
  pause
)
