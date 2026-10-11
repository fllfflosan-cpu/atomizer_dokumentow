from app.atomizer import LegalDocumentAtomizer
from app.database import LegalDatabase
from app.versioning import ActVersioning
from app.config import DB_PATH, FILES_DIR

# 1. Inicjalizacja bazy i atomizatora
db = LegalDatabase()
act_id = "USTAWA_2024_001"

# 2. Parsujemy nowy dokument z pliku .docx
atomizer = LegalDocumentAtomizer("files/ustawa_nowa.docx")
new_atoms = atomizer.parse_to_atoms()

# 3. Sprawdzamy, czy w bazie jest już stara wersja tego aktu
existing_act = db.get_act(act_id)

if existing_act:
    # Jeśli akt już istnieje, generujemy diff ("git diff")
    old_atoms = existing_act["data"]
    diff_lines = ActVersioning.generate_diff(old_atoms, new_atoms)
    summary = ActVersioning.get_diff_summary(diff_lines)
    
    print(f"Wykryto zmiany w akcie {act_id}!")
    print(f"Dodano linii: {summary['added_lines']}, Usunięto linii: {summary['removed_lines']}")
    
    # Możesz zapisać diff do historii lub wyświetlić prowadzącemu
    print("\nSzczegóły zmian:")
    print("\n".join(diff_lines[:15])) # Wyświetlamy pierwsze linie diffa
else:
    print(f"Dodano nowy akt prawny do bazy: {act_id}")

# 4. Zapisujemy nową wersję do bazy danych
db.save_act(act_id, "Ustawa o systemach atomizacji", new_atoms)