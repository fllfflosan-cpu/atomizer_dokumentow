# ==========================================
# MODUŁ: ATOMIZER DOKUMENTÓW PRAWNYCH (.docx / .doc)
# ==========================================

import os
import json
import xml.etree.ElementTree as ET
import zipfile
import re
import tempfile
from app.config import DB_PATH, FILES_DIR

class LegalDocumentAtomizer:
    """
    Klasa odpowiedzialna za wczytywanie dokumentów prawnych w formacie .docx lub .doc 
    (stare pliki .doc są automatycznie konwertowane w tle przez MS Word), 
    wyciąganie stylów oraz rozbijanie struktury tekstu na pojedyncze atomy (artykuły, ustępy, punkty).
    """

    def __init__(self, file_path):
        """
        Inicjalizacja atomizatora. Obsługuje pliki .docx oraz automatycznie konwertuje .doc -> .docx.
        :param file_path: Ścieżka do pliku wejściowego.
        """
        self.original_path = file_path
        self.docx_path = self._prepare_docx(file_path)
        
        # Definicja przestrzeni nazw XML wymagana przez standard OpenXML (Word)
        self.namespaces = {
            'w': 'http://schemas.openxmlformats.org/wordprocessingml/2006/main'
        }

    def _prepare_docx(self, file_path):
        """
        Sprawdza format pliku. Jeśli to stary .doc, kopiuje go do folderu tymczasowego 
        (co zdejmuje flagę widoku chronionego systemu Windows) i konwertuje przez MS Word.
        """
        ext = os.path.splitext(file_path)[1].lower()
        
        if ext == '.docx':
            return file_path
        elif ext == '.doc':
            try:
                import win32com.client as win32
                import shutil
            except ImportError:
                raise Exception("Aby obsługiwać pliki .doc, zainstaluj: pip install pywin32")
            
            temp_dir = tempfile.gettempdir()
            
            # Krok 1: Kopiujemy oryginalny .doc do bezpiecznej lokalizacji tymczasowej,
            # żeby system Windows zdjął z niego blokadę "Widoku chronionego" (Mark of the Web).
            safe_local_doc = os.path.join(temp_dir, "temp_source_legal_act.doc")
            shutil.copy2(file_path, safe_local_doc)
            
            temp_docx = os.path.join(temp_dir, "converted_legal_act.docx")
            if os.path.exists(temp_docx):
                try:
                    os.remove(temp_docx)
                except Exception:
                    pass

            word = win32.Dispatch("Word.Application")
            word.Visible = False
            word.DisplayAlerts = 0
            
            try:
                # Otwieramy lokalną, bezpieczną kopie pliku .doc
                doc = word.Documents.Open(os.path.abspath(safe_local_doc), ReadOnly=True, AddToRecentFiles=False)
                doc.SaveAs2(temp_docx, FileFormat=16) # Format 16 = .docx
                doc.Close(SaveChanges=0)
            except Exception as e:
                word.Quit()
                raise Exception(f"Błąd konwersji pliku .doc przez MS Word: {e}")
            
            word.Quit()
            
            # Sprzątamy plik roboczy .doc
            try:
                os.remove(safe_local_doc)
            except Exception:
                pass
                
            return temp_docx
        else:
            raise ValueError(f"Nieobsługiwany format pliku: {ext}. Dozwolone: .docx, .doc")

    def _get_style_mapping(self, docx_zip):
        """
        Wczytuje słownik stylów z pliku styles.xml wewnątrz archiwum docx.
        """
        style_map = {}
        try:
            styles_xml = docx_zip.read('word/styles.xml')
            root = ET.fromstring(styles_xml)
            
            for style in root.findall('w:style', self.namespaces):
                style_id = style.get('{http://schemas.openxmlformats.org/wordprocessingml/2006/main}styleId')
                name_elem = style.find('w:name', self.namespaces)
                if name_elem is not None:
                    style_name = name_elem.get('{http://schemas.openxmlformats.org/wordprocessingml/2006/main}val')
                    style_map[style_id] = style_name.upper()
        except Exception:
            pass
            
        return style_map

    def parse_to_atoms(self):
        atoms = []
        stack = []

        with zipfile.ZipFile(self.docx_path, 'r') as docx_zip:
            document_xml = docx_zip.read('word/document.xml')
            root = ET.fromstring(document_xml)

            for p in root.iter('{http://schemas.openxmlformats.org/wordprocessingml/2006/main}p'):
                texts = []
                for t in p.iter('{http://schemas.openxmlformats.org/wordprocessingml/2006/main}t'):
                    if t.text:
                        texts.append(t.text)
                paragraph_text = "".join(texts).strip()

                if not paragraph_text:
                    continue  

                # Elastyczne dopasowanie typów (uwzględnia ewentualne spacje/tabulacje na początku)
                atom_type = "GENERAL"
                
                if re.match(r'^(?:[ \t]*)(Rozdział|Oddział)\s+\d+', paragraph_text, re.IGNORECASE):
                    atom_type = "CHAPTER_OR_SECTION"
                elif re.match(r'^(?:[ \t]*)(?:Art\.|Artykuł|§)\s*\d+', paragraph_text, re.IGNORECASE):
                    atom_type = "SECTION_OR_ARTICLE"
                elif re.match(r'^(?:[ \t]*)\d+\.\s+\S', paragraph_text):
                    atom_type = "UST"
                elif re.match(r'^(?:[ \t]*)\d+\)\s*', paragraph_text):
                    atom_type = "POINT"
                elif re.match(r'^(?:[ \t]*)[a-zA-Z]\)\s*', paragraph_text):
                    atom_type = "SUBPOINT_LETTER"

                new_atom = {
                    "type": atom_type,
                    "content": paragraph_text,
                    "children": []
                }

                # Hierarchia poziomów (0 - najwyższy, 4 - najniższy)
                level_map = {
                    "CHAPTER_OR_SECTION": 0,
                    "SECTION_OR_ARTICLE": 1,
                    "UST": 2,
                    "POINT": 3,
                    "SUBPOINT_LETTER": 4,
                    "GENERAL": 5
                }
                current_level = level_map.get(atom_type, 5)

                if atom_type == "GENERAL":
                    if stack:
                        stack[-1]["content"] += " " + paragraph_text
                    else:
                        atoms.append(new_atom)
                    continue

                # Oczyszczanie stosu dla elementów wyższego lub równego rzędu
                while stack and level_map.get(stack[-1]["type"], 5) >= current_level:
                    stack.pop()

                if stack:
                    stack[-1]["children"].append(new_atom)
                else:
                    atoms.append(new_atom)

                stack.append(new_atom)

        return atoms

    def save_to_json(self, output_path):
        data = self.parse_to_atoms()
        os.makedirs(os.path.dirname(output_path), exist_ok=True)
        with open(output_path, 'w', encoding='utf-8') as f:
            json.dump(data, f, ensure_ascii=False, indent=4)

    def save_to_xml(self, output_path):
        data = self.parse_to_atoms()
        os.makedirs(os.path.dirname(output_path), exist_ok=True)
        
        root_elem = ET.Element("LegalAct")
        
        def _build_xml_rec(parent_xml, children_list):
            for item in children_list:
                item_elem = ET.SubElement(parent_xml, item.get("type", "ITEM"))
                content_elem = ET.SubElement(item_elem, "Content")
                content_elem.text = item.get("content", "")
                
                if item.get("children"):
                    children_container = ET.SubElement(item_elem, "Children")
                    _build_xml_rec(children_container, item["children"])

        _build_xml_rec(root_elem, data)

        tree = ET.ElementTree(root_elem)
        ET.indent(tree, space="    ", level=0)
        tree.write(output_path, encoding="utf-8", xml_declaration=True)