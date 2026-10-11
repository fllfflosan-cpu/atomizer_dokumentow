# ==========================================
# MODUŁ: BAZA DANYCH AKTÓW PRAWNYCH (SQLite)
# ==========================================

import sqlite3
import json
import os
from app.config import DB_PATH, FILES_DIR

class LegalDatabase:
    """
    Klasa zarządzająca bazą danych SQLite dla zatomizowanych aktów prawnych.
    Każdy akt prawny jest jednoznacznie identyfikowany przez unikalny string (ID).
    """

    def __init__(self, db_path="files/legal_acts.db"):
        """
        Inicjalizacja bazy danych.
        :param db_path: Ścieżka do pliku bazy danych SQLite.
        """
        self.db_path = db_path
        # Upewniamy się, że katalog na bazę danych istnieje
        os.makedirs(os.path.dirname(self.db_path), exist_ok=True)
        self.setup_database()

    def _get_connection(self):
        """Pomocnicza metoda zwracająca połączenie z bazą SQLite."""
        return sqlite3.connect(self.db_path)

    def setup_database(self):
        """
        Tworzy niezbędne tabele, w tym tabelę główną oraz tabelę historii zmian.
        """
        with self._get_connection() as conn:
            cursor = conn.cursor()
            # Tabela aktualnego stanu aktu
            cursor.execute("""
                CREATE TABLE IF NOT EXISTS acts (
                    act_id TEXT PRIMARY KEY,
                    title TEXT NOT NULL,
                    json_data TEXT NOT NULL,
                    updated_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP
                )
            """)
            # Tabela historii zmian (wersjonowanie)
            cursor.execute("""
                CREATE TABLE IF NOT EXISTS act_history (
                    id INTEGER PRIMARY KEY AUTOINCREMENT,
                    act_id TEXT NOT NULL,
                    title TEXT NOT NULL,
                    json_data TEXT NOT NULL,
                    saved_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP
                )
            """)
            conn.commit()

    def save_act(self, act_id, title, atoms_data):
        """
        Zapisuje akt w tabeli głównej ORAZ tworzy wpis w historii zmian.
        """
        json_string = json.dumps(atoms_data, ensure_ascii=False, indent=4)

        with self._get_connection() as conn:
            cursor = conn.cursor()
            
            # 1. Zapisz do tabeli głównej (aktualny stan)
            cursor.execute("""
                INSERT OR REPLACE INTO acts (act_id, title, json_data, updated_at)
                VALUES (?, ?, ?, CURRENT_TIMESTAMP)
            """, (act_id, title, json_string))

            # 2. Dodaj wpis do historii zmian
            cursor.execute("""
                INSERT INTO act_history (act_id, title, json_data)
                VALUES (?, ?, ?)
            """, (act_id, title, json_string))
            
            conn.commit()

    def get_act_history(self, act_id):
        """
        Zwraca historię modyfikacji dla konkretnego aktu prawnego.
        """
        with self._get_connection() as conn:
            cursor = conn.cursor()
            cursor.execute("""
                SELECT id, title, saved_at 
                FROM act_history 
                WHERE act_id = ? 
                ORDER BY saved_at DESC
            """, (act_id,))
            rows = cursor.fetchall()
            
            history = []
            for row in rows:
                history.append({
                    "version_id": row[0],
                    "title": row[1],
                    "saved_at": row[2]
                })
            return history

    def get_act(self, act_id):
        """
        Pobiera zatomizowany akt prawny z bazy danych po jego unikalnym stringu (ID).
        
        :param act_id: Unikalny string identyfikujący akt.
        :return: Słownik/lista z danymi JSON aktu lub None, jeśli nie znaleziono.
        """
        with self._get_connection() as conn:
            cursor = conn.cursor()
            cursor.execute("SELECT title, json_data, updated_at FROM acts WHERE act_id = ?", (act_id,))
            row = cursor.fetchone()
            
            if row:
                title, json_string, updated_at = row
                return {
                    "act_id": act_id,
                    "title": title,
                    "updated_at": updated_at,
                    "data": json.loads(json_string) # Odkodowanie JSON z powrotem do struktury Pythona
                }
            return None

    def list_all_acts(self):
        """
        Zwraca listę wszystkich zapisanych aktów prawnych (ich ID, tytuły i daty modyfikacji).
        Przydatne do wyświetlenia historii lub wyboru aktu.
        """
        with self._get_connection() as conn:
            cursor = conn.cursor()
            cursor.execute("SELECT act_id, title, updated_at FROM acts")
            rows = cursor.fetchall()
            
            acts_list = []
            for row in rows:
                acts_list.append({
                    "act_id": row[0],
                    "title": row[1],
                    "updated_at": row[2]
                })
            return acts_list

    