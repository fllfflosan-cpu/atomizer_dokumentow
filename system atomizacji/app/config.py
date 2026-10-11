# ==========================================
# MODUŁ: KONFIGURACJA APLIKACJI
# ==========================================

import os

# Główny katalog projektu (katalog nadrzędny względem folderu 'app')
BASE_DIR = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))

# Katalog na pliki (dokumenty .docx, wygenerowane JSON/XML)
FILES_DIR = os.path.join(BASE_DIR, "files")

# Ścieżka do bazy danych SQLite
DB_PATH = os.path.join(FILES_DIR, "legal_acts.db")

# Upewnienie się, że folder na pliki fizycznie istnieje przy starcie aplikacji
os.makedirs(FILES_DIR, exist_ok=True)