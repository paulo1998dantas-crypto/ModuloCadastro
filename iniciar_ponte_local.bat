@echo off
cd /d "%~dp0"
py -3 local_bridge.py
if errorlevel 1 (
  echo.
  echo A ponte local encerrou com erro. Confira o ponte_config.json ou rode:
  echo py -3 local_bridge.py --server https://SEU-APP.onrender.com --token SEU_TOKEN
  echo.
  pause
)
