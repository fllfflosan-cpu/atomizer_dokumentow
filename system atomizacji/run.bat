@echo off
echo Sprawdzanie i instalacja wymaganych bibliotek...
python -m pip install --upgrade pip
python -m pip install -r requirements.txt

echo.
echo Uruchamianie Systemu Atomizacji Aktow Prawnych...
python -m app.gui
pause