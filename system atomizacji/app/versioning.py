# ==========================================
# MODUŁ: WERSJONOWANIE I HISTORIA ZMIAN (A'la Git)
# ==========================================

import difflib
import json
from app.config import DB_PATH, FILES_DIR

class ActVersioning:
    """
    Klasa odpowiedzialna za śledzenie historii zmian aktów prawnych,
    generowanie różnic (diffów) w stylu gita oraz zarządzanie rewizjami.
    """

    @staticmethod
    def generate_diff(old_atoms, new_atoms):
        """
        Porównuje starą i nową wersję zatomizowanego aktu prawnego 
        i generuje różnice w formacie uniwersalnego diffa (jak 'git diff').
        
        :param old_atoms: Poprzednia struktura atomów (lista/słownik).
        :param new_atoms: Nowa struktura atomów (lista/słownik).
        :return: Lista linii opisujących zmiany (diff).
        """
        # Zamieniamy struktury JSON na sformatowane stringi w liniach
        old_lines = json.dumps(old_atoms, ensure_ascii=False, indent=4).splitlines()
        new_lines = json.dumps(new_atoms, ensure_ascii=False, indent=4).splitlines()

        # Generowanie diffa linia po linii
        diff_generator = difflib.unified_diff(
            old_lines,
            new_lines,
            fromfile="wersja_poprzednia (HEAD~1)",
            tofile="wersja_nowa (WORKING)",
            lineterm=""
        )

        return list(diff_generator)

    @staticmethod
    def get_diff_summary(diff_lines):
        """
        Analizuje wygenerowany diff i podlicza liczbę dodanych (+) i usuniętych (-) linii,
        co daje fajne podsumowanie w stylu Gita (np. '+5 -2').
        """
        added = 0
        removed = 0
        for line in diff_lines:
            if line.startswith('+') and not line.startswith('+++'):
                added += 1
            elif line.startswith('-') and not line.startswith('---'):
                removed += 1
                
        return {"added_lines": added, "removed_lines": removed}