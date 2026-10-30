@echo off
REM Сборка instrOT в папку dist\instrOT
REM Требуется Python 3.10+ и установленные зависимости.

setlocal
cd /d "%~dp0.."

if not exist ".venv\Scripts\python.exe" (
    echo [1/3] Создание окружения...
    python -m venv .venv || goto :error
)

echo [2/3] Установка зависимостей...
".venv\Scripts\python.exe" -m pip install -q --upgrade pip || goto :error
".venv\Scripts\python.exe" -m pip install -q -r requirements.txt pyinstaller || goto :error

echo [3/3] Сборка...
".venv\Scripts\python.exe" scripts\build.py || goto :error

echo.
echo Готово. Папка: dist\instrOT
echo Для запуска ярлыка перетащите dist\instrOT\instrOT.exe на рабочий стол
echo (создать ярлык) и запускайте его. Данные создаются отдельно от программы.
goto :eof

:error
echo.
echo Сборка не выполнена. Смотрите сообщения выше.
exit /b 1
