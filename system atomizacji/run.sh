#!/bin/bash

echo "=================================================="
echo " Sprawdzanie i instalacja wymaganych bibliotek... "
echo "=================================================="

# Sprawdzenie czy python3 jest dostępny
if ! command -v python3 &> /dev/null
then
    echo "[BŁĄD] Python3 nie jest zainstalowany w systemie."
    exit 1
fi

# Upgrade pip oraz instalacja paczek z requirements.txt
python3 -m pip install --upgrade pip
python3 -m pip install -r requirements.txt

echo ""
echo "Uruchamianie Systemu Atomizacji Aktów Prawnych..."
python3 -m app.gui