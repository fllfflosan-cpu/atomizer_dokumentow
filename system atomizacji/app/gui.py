import os
import json
import tkinter as tk
from tkinter import ttk, filedialog, messagebox

# Pełne importy pakietowe wymagane przy uruchamianiu przez -m app.gui
from app.atomizer import LegalDocumentAtomizer
from app.database import LegalDatabase
from app.versioning import ActVersioning

class LegalApp:
    def __init__(self, root):
        self.root = root
        self.root.title("System Atomizacji Aktów Prawnych")
        self.root.geometry("950x700")

        self.db = LegalDatabase()

        # Główny kontener na zakładki (Widok podzielony na ekrany)
        self.notebook = ttk.Notebook(self.root)
        self.notebook.pack(fill="both", expand=True, padx=10, pady=10)

        # Tworzenie zakładek
        self.tab_import = ttk.Frame(self.notebook)
        self.tab_hierarchy = ttk.Frame(self.notebook)
        self.history_tab = ttk.Frame(self.notebook)  # <-- Zarejestrowana ramka

        self.notebook.add(self.tab_import, text="📁 Import i Przetwarzanie")
        self.notebook.add(self.tab_hierarchy, text="🌳 Widok Hierarchiczny")
        self.notebook.add(self.history_tab, text="📜 Historia Zmian")  # <-- Dodana zakładka

        # Inicjalizacja zawartości zakładek
        self.setup_import_tab()
        self.setup_hierarchy_tab()
        self._setup_history_tab()  # <-- Wywołanie konfiguracji widoku historii

    # ==========================================
    # ZAKŁADKA 1: IMPORT I WPROWADZANIE DANYCH
    # ==========================================
    def setup_import_tab(self):
        form_frame = ttk.LabelFrame(self.tab_import, text=" Parametry Dokumentu ", padding=15)
        form_frame.pack(fill="x", padx=10, pady=10)

        ttk.Label(form_frame, text="Plik DOCX:").grid(row=0, column=0, sticky="w", pady=5)
        self.file_entry = ttk.Entry(form_frame, width=50)
        self.file_entry.grid(row=0, column=1, padx=5, pady=5)
        ttk.Button(form_frame, text="Przeglądaj...", command=self.browse_file).grid(row=0, column=2, padx=5, pady=5)

        ttk.Label(form_frame, text="Unikalne ID Aktu:").grid(row=1, column=0, sticky="w", pady=5)
        self.id_entry = ttk.Entry(form_frame, width=30)
        self.id_entry.grid(row=1, column=1, sticky="w", padx=5, pady=5)

        ttk.Label(form_frame, text="Tytuł / Nazwa:").grid(row=2, column=0, sticky="w", pady=5)
        self.name_entry = ttk.Entry(form_frame, width=50)
        self.name_entry.grid(row=2, column=1, sticky="w", padx=5, pady=5)

        ttk.Button(self.tab_import, text="🚀 Zatomizuj i Zapisz Akt", command=self.process_document).pack(pady=10)

        log_frame = ttk.LabelFrame(self.tab_import, text=" Logi Systemowe ", padding=10)
        log_frame.pack(fill="both", expand=True, padx=10, pady=10)

        self.log_text = tk.Text(log_frame, wrap="word", height=12, bg="#1e1e1e", fg="#d4d4d4", font=("Consolas", 10))
        self.log_text.pack(side="left", fill="both", expand=True)

        scrollbar = ttk.Scrollbar(log_frame, orient="vertical", command=self.log_text.yview)
        scrollbar.pack(side="right", fill="y")
        self.log_text.configure(yscrollcommand=scrollbar.set)

    def browse_file(self):
        filename = filedialog.askopenfilename(
            filetypes=[("Dokumenty Word", "*.docx;*.doc"), ("Wszystkie pliki", "*.*")]
        )
        if filename:
            self.file_entry.delete(0, tk.END)
            self.file_entry.insert(0, filename)

    def log(self, message):
        self.log_text.insert(tk.END, message + "\n")
        self.log_text.see(tk.END)

    def process_document(self):
        file_path = self.file_entry.get().strip()
        act_id = self.id_entry.get().strip()
        title = self.name_entry.get().strip()

        if not file_path or not os.path.exists(file_path):
            messagebox.showerror("Błąd", "Wskaż poprawny plik .docx / .doc!")
            return
        if not act_id:
            messagebox.showerror("Błąd", "Musisz podać unikalny string ID aktu!")
            return

        try:
            self.log(f"\n--- Przetwarzanie: {act_id} ---")
            
            atomizer = LegalDocumentAtomizer(file_path)
            new_atoms = atomizer.parse_to_atoms()
            self.log(f"[OK] Zatomizowano pomyślnie. Elementów głównych: {len(new_atoms)}")

            existing_act = self.db.get_act(act_id)
            if existing_act:
                old_atoms = existing_act["data"]
                diff_lines = ActVersioning.generate_diff(old_atoms, new_atoms)
                summary = ActVersioning.get_diff_summary(diff_lines)
                self.log(f"[GIT DIFF] Wykryto poprzednią wersję! Dodano: {summary['added_lines']} | Usunięto: {summary['removed_lines']}")
            else:
                self.log(f"[INFO] Tworzysz nowy akt w bazie.")

            self.db.save_act(act_id, title if title else "Brak tytułu", new_atoms)
            self.log(f"[SUKCES] Akt '{act_id}' zapisany w bazie danych!")
            messagebox.showinfo("Sukces", f"Akt '{act_id}' zatomizowany i zapisany pomyślnie!")

            self.load_acts_list()

        except Exception as e:
            messagebox.showerror("Błąd krytyczny", str(e))
            self.log(f"[BŁĄD] {str(e)}")

    # ==========================================
    # ZAKŁADKA 2: WIDOK HIERARCHICZNY (DRZEWO)
    # ==========================================
    def setup_hierarchy_tab(self):
        top_frame = ttk.Frame(self.tab_hierarchy, padding=5)
        top_frame.pack(fill="x", padx=5, pady=5)

        ttk.Label(top_frame, text="Wybierz zapisany akt:").pack(side="left", padx=5)
        
        self.act_combobox = ttk.Combobox(top_frame, state="readonly", width=40)
        self.act_combobox.pack(side="left", padx=5)
        self.act_combobox.bind("<<ComboboxSelected>>", self.display_hierarchy)

        ttk.Button(top_frame, text="🔄 Odśwież listę", command=self.load_acts_list).pack(side="left", padx=5)

        tree_frame = ttk.Frame(self.tab_hierarchy, padding=5)
        tree_frame.pack(fill="both", expand=True, padx=5, pady=5)

        self.tree = ttk.Treeview(tree_frame, columns=("type"), selectmode="browse")
        self.tree.heading("#0", text="Struktura / Treść elementu", anchor="w")
        self.tree.heading("type", text="Typ atomu", anchor="w")
        self.tree.column("#0", width=600, anchor="w")
        self.tree.column("type", width=150, anchor="w")

        tree_scroll = ttk.Scrollbar(tree_frame, orient="vertical", command=self.tree.yview)
        self.tree.configure(yscrollcommand=tree_scroll.set)

        self.tree.pack(side="left", fill="both", expand=True)
        tree_scroll.pack(side="right", fill="y")

        self.load_acts_list()

    def load_acts_list(self):
        """Pobiera listę ID aktów z bazy i wrzuca do rozwijanego menu"""
        try:
            acts_data = self.db.list_all_acts()
            act_list = [f"{act['act_id']} - {act['title']}" for act in acts_data]
            self.act_combobox["values"] = act_list
            
            if act_list:
                self.act_combobox.current(0)
                self.display_hierarchy(None)
            else:
                self.log("[INFO] Baza danych nie zawiera jeszcze żadnych aktów.")
        except Exception as e:
            self.log(f"[BŁĄD BAZY] Nie udało się pobrać listy aktów: {e}")

    def display_hierarchy(self, event):
        for item in self.tree.get_children():
            self.tree.delete(item)

        selected = self.act_combobox.get()
        if not selected:
            return

        act_id = selected.split(" - ")[0]
        act_data = self.db.get_act(act_id)

        if not act_data or "data" not in act_data:
            return

        atoms = act_data["data"]
        for atom in atoms:
            self.add_atom_to_tree("", atom)

    def add_atom_to_tree(self, parent_node, atom):
        atom_type = atom.get("type", "GENERAL")
        content = atom.get("content", "Brak treści")
        display_text = content[:80] + ("..." if len(content) > 80 else "")

        # Tworzenie węzła w drzewie Tkinter
        node_id = self.tree.insert(parent_node, "end", text=display_text, values=(atom_type,))

        # Rekurencyjne dodawanie dzieci z uniwersalnego klucza "children"
        children = atom.get("children", [])
        if children and isinstance(children, list):
            for child in children:
                self.add_atom_to_tree(node_id, child)

    # ==========================================
    # ZAKŁADKA 3: HISTORIA ZMIAN
    # ==========================================
    def _setup_history_tab(self):
        """
        Tworzy elementy interfejsu dla zakładki historii zmian.
        """
        top_frame = ttk.LabelFrame(self.history_tab, text="Wyszukaj historię aktu")
        top_frame.pack(fill="x", padx=10, pady=10)

        ttk.Label(top_frame, text="ID Aktu:").pack(side="left", padx=5, pady=5)
        self.history_id_entry = ttk.Entry(top_frame, width=20)
        self.history_id_entry.pack(side="left", padx=5, pady=5)

        load_btn = ttk.Button(top_frame, text="Pokaż historię", command=self.load_act_history)
        load_btn.pack(side="left", padx=10, pady=5)

        table_frame = ttk.Frame(self.history_tab)
        table_frame.pack(fill="both", expand=True, padx=10, pady=5)

        columns = ("version_id", "title", "saved_at")
        self.history_tree = ttk.Treeview(table_frame, columns=columns, show="headings")
        
        self.history_tree.heading("version_id", text="ID Wersji")
        self.history_tree.heading("title", text="Tytuł / Nazwa Aktu")
        self.history_tree.heading("saved_at", text="Data Modyfikacji")
        
        self.history_tree.column("version_id", width=80, anchor="center")
        self.history_tree.column("title", width=350, anchor="w")
        self.history_tree.column("saved_at", width=180, anchor="center")

        scrollbar = ttk.Scrollbar(table_frame, orient="vertical", command=self.history_tree.yview)
        self.history_tree.configure(yscrollcommand=scrollbar.set)

        self.history_tree.pack(side="left", fill="both", expand=True)
        scrollbar.pack(side="right", fill="y")

    def load_act_history(self):
        """
        Pobiera z bazy danych historię zmian dla wskazanego ID aktu i ładuje do tabeli.
        """
        act_id = self.history_id_entry.get().strip()
        if not act_id:
            self.log("⚠️ Podaj ID aktu, aby sprawdzić jego historię.")
            return

        for row in self.history_tree.get_children():
            self.history_tree.delete(row)

        history_records = self.db.get_act_history(act_id)

        if not history_records:
            self.log(f"ℹ️ Brak historii zmian dla aktu o ID: {act_id}")
            return

        for record in history_records:
            self.history_tree.insert(
                "", 
                "end", 
                values=(record["version_id"], record["title"], record["saved_at"])
            )
        
        self.log(f"📂 Wczytano historię dla aktu '{act_id}' (znaleziono wersji: {len(history_records)}).")

if __name__ == "__main__":
    root = tk.Tk()
    app = LegalApp(root)
    root.mainloop()