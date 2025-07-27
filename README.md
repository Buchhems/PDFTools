# 🛠️ PDF-Tools – PDF aus Word erzeugen & Metadaten entfernen

**PDF-Tools** ist ein praktisches Windows-Tool zur automatisierten Erstellung von PDF-Dateien aus Word-Dokumenten (.docx) sowie zur Entfernung sensibler Metadaten aus bestehenden PDF-Dateien.

<img width="267" alt="Screenshot 2023-03-05 163200" src="https://user-images.githubusercontent.com/75378632/222970050-cfb7194c-1ebb-46a3-95fc-bf6127d8d1a4.png">

Icon by Some icon firm (https://icon-icons.com/de/pack/Online-Learning/3480)

---

## 🚀 Funktionen

- 📄 Konvertiert `.docx`-Dateien in `.pdf` mit Microsoft Word
- 🧹 Entfernt Metadaten wie Autor, Titel, Erstellungsdatum aus PDFs
- 🗑️ Optional: Kommentare löschen und Änderungsverfolgung beenden
- 🖥️ Benutzerfreundliche Oberfläche mit Tkinter
- 🧠 Automatische Erkennung und Beendigung laufender Word-Instanzen
- 🔐 Datenschutzfreundlich durch Metadatenbereinigung

---

## 🖥️ Benutzeroberfläche

| Element                          | Beschreibung                                           |
|----------------------------------|--------------------------------------------------------|
| **PDF aus Docx erzeugen**        | Startet die Konvertierung von Word zu PDF             |
| **Metadaten aus PDF entfernen**  | Bereinigt ausgewählte PDFs von sensiblen Informationen|
| **Checkbox**                     | Aktiviert das Löschen von Kommentaren und Revisionen  |
| **Logo/Icon**                    | Optionales Bild zur optischen Gestaltung              |

---

## 📦 Voraussetzungen

- Windows-PC mit installiertem **Microsoft Word**
- Python 3.x
- Module:
  - `comtypes`
  - `pypdf`
  - `psutil`
  - `tkinter` (Standardmodul)

## 📁 Dateistruktur

Die wichtigsten Dateien im Projekt:

| Datei             | Zweck                                      |
|-------------------|--------------------------------------------|
| `pdf_tools.py`    | Hauptskript mit GUI und Konvertierungslogik |
| `PDFTools.ico`    | Fenster-Icon für die Anwendung              |
| `hla.png`         | Bild für die grafische Oberfläche (z. B. Logo) |

---

## 🧑‍🏫 Anleitung

### 📄 PDF aus Word erzeugen

1. Schließe alle geöffneten Word-Fenster.
2. Klicke auf **„PDF aus Docx erzeugen“**.
3. Wähle die gewünschten `.docx`-Dateien aus.
4. Optional: Checkbox aktivieren, um Kommentare zu löschen und Änderungen anzunehmen.
5. Die PDFs werden im gleichen Ordner wie die Word-Dateien gespeichert.

### 🧹 Metadaten entfernen

1. Klicke auf **„Metadaten aus PDF entfernen“**.
2. Wähle die gewünschten `.pdf`-Dateien aus.
3. Die bereinigten PDFs werden überschrieben (Originale werden temporär gesichert).

---

## ⚠️ Hinweise

- ❗ Bei der PDF-Erzeugung dürfen keine anderen Word-Instanzen aktiv sein.
- 🔐 Die Metadatenbereinigung überschreibt die Originaldateien.
- 🧼 Temporäre Dateien werden automatisch gelöscht.
- Benutzung auf eigene Gefahr.
- Keine Haftung für Datenverlust oder fehlerhafte Konvertierungen.

---

## 🧩 Erweiterungsideen

- Auswahl eines Zielordners für PDF-Ausgabe
- Fortschrittsanzeige bei großen Dateimengen
- Drag & Drop Unterstützung

---

## 🧪 Kompilierung mit PyInstaller

Zur Erstellung einer ausführbaren Datei:

```bash
pyinstaller --onefile --windowed pdf_tools.py
