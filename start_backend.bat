@echo off
setlocal

set "ROOT_DIR=%~dp0"
cd /d "%ROOT_DIR%backend" || goto :cd_error

set "BACKEND_PORT=%~1"
if "%BACKEND_PORT%"=="" set "BACKEND_PORT=8001"

if exist ".venv311\Scripts\python.exe" (
  set "PYTHON_EXE=.venv311\Scripts\python.exe"
) else if exist ".venv\Scripts\python.exe" (
  set "PYTHON_EXE=.venv\Scripts\python.exe"
) else (
  echo [ERROR] Python virtual environment not found.
  echo Expected one of:
  echo   backend\.venv311\Scripts\python.exe
  echo   backend\.venv\Scripts\python.exe
  exit /b 1
)

set "AIPPT_EXPORT_ENGINE=ai_to_pptx"

echo [INFO] Working dir: %CD%
echo [INFO] Export engine: %AIPPT_EXPORT_ENGINE%
echo [INFO] Starting backend at http://127.0.0.1:%BACKEND_PORT%

"%PYTHON_EXE%" -m uvicorn app.main:app --reload --host 127.0.0.1 --port %BACKEND_PORT%
exit /b %ERRORLEVEL%

:cd_error
echo [ERROR] Failed to enter backend directory.
exit /b 1
