import os
import sys
import psutil
from tkinter import (N, W, E, BooleanVar, Button, Canvas, Checkbutton, Label, PhotoImage, Tk, Toplevel, StringVar, OptionMenu, Frame, messagebox, filedialog, Menu)
from pypdf import PdfReader, PdfWriter
import threading
import comtypes.client
import subprocess
import shutil


class ToolTip:
    """Create Tooltip for a widget"""
    def __init__(self, widget, text='Widget Info'):
        self.widget = widget
        self.text = text
        self.widget.bind("<Enter>", self.enter)
        self.widget.bind("<Leave>", self.leave)
        self.tooltip_window = None

    def enter(self, event=None):
        x = self.widget.winfo_rootx() + 25
        y = self.widget.winfo_rooty() + 25

        self.tooltip_window = tw = Toplevel(self.widget)
        tw.wm_overrideredirect(True)
        tw.wm_geometry("+%d+%d" % (x, y))

        label = Label(tw, text=self.text, justify='left',
                     background="#FFFFE0", relief='solid', borderwidth=1,
                     font=("Segoe UI", "8", "normal"))
        label.pack(ipadx=1)

    def leave(self, event=None):
        if self.tooltip_window:
            self.tooltip_window.destroy()
            self.tooltip_window = None


def resource_path(relative_path: str) -> str:
    #Get absolute path to resource, works for dev and for PyInstaller
    try:
        # PyInstaller creates a temp folder and stores path in _MEIPASS
        base_path = sys._MEIPASS
    except Exception:
        base_path = os.path.abspath(".")

    return os.path.join(base_path, relative_path)
    
    
def docx_to_pdf(docx_path: str, pdf_path: str, disable_track_changes: bool) -> bool:
    word = comtypes.client.CreateObject('Word.Application')
    word.Visible = False
    word.DisplayAlerts = False

    try:
        doc = word.Documents.Open(os.path.abspath(docx_path))
    except comtypes.COMError as e:
        messagebox.showerror("Word öffnen fehlgeschlagen", f"{docx_path}\n{e}")
        word.Quit()
        return False

    if disable_track_changes:
        delete_comments(doc)

    try:
        doc.SaveAs(os.path.abspath(pdf_path), FileFormat=17)
    except comtypes.COMError as e:
        messagebox.showerror("PDF-Speicherung fehlgeschlagen", f"{pdf_path}\n{e}")
        doc.Close()
        word.Quit()
        return False

    doc.Close()
    word.Quit()
    return True


def delete_comments(doc):
    # Iterate over all comments and delete them
    for comment in doc.Comments:
        comment.Delete()
        
    # Only disable all TrackRevisions based on the user's choice
    doc.TrackRevisions = False

    # Accept all revisions
    doc.AcceptAllRevisions()

def select_docx_files(convert_btn, meta_btn, disable_track_changes_var):
    pdf_count = 0

    # set Button status
    convert_btn.config(state="disabled", text="... einen Moment bitte ...", bg="#E5E7EB", fg="black")
    root.update()

    # parallel word instances warning
    messagebox.showwarning(
        title="Word-Fenster schließen!",
        message=(
            "Schließen Sie alle Word-Fenster und klicken Sie dann auf OK.\n\n"
            "Erläuterung:\nUm fehlerfrei PDF zu erzeugen, dürfen keine anderen Word Instanzen parallel laufen."
        )
    )

    # thread to close all (former) word processes
    kill_thread = threading.Thread(target=kill_all_word, daemon=True)
    kill_thread.start()
    kill_thread.join()  # wait till all dead :)

    # choose files
    docx_filenames = filedialog.askopenfilenames(
        title='Word-Dateien zur Erzeugung von PDF auswählen',
        filetypes=[('Word Dokumente', '*.docx')]
    )

    # conversion
    for docx_path in docx_filenames:
        pdf_path = os.path.splitext(docx_path)[0] + ".pdf"
        success = docx_to_pdf(docx_path, pdf_path, disable_track_changes_var.get())
        if success:
            pdf_count += 1

        convert_btn.config(text=f"Erzeugte PDF: {pdf_count}")
        root.update()
    
    revert_button_text(convert_button, meta_button)  
    
    if pdf_count:
        messagebox.showinfo("Erledigt", f"{pdf_count} PDF wurden erstellt.")
    
    
        
def pdf_edit(meta_btn, pdf_format_var):
    pdfcount = 0
    meta_button.config(state="disabled", text="Bearbeite PDF...", bg="#E5E7EB", fg="black")

    files = filedialog.askopenfilenames(
        title='PDF auswählen',
        filetypes=[('PDF Dokumente', '*.pdf')]
    )

    for file in files:
        name, extension = os.path.splitext(file)
        temp_name = name + "_todo" + extension
        temp_clean = name + "_clean" + extension

        try:
            os.rename(file, temp_name)
        except (PermissionError, FileExistsError) as e:
            messagebox.showerror("Fehler", f"Datei {file} konnte nicht umbenannt werden: {e}")
            continue

        # 1. clean metadata with pypdf
        try:
            reader = PdfReader(temp_name)
            writer = PdfWriter()
            for page in reader.pages:
                writer.add_page(page)
            writer.add_metadata({})
            if "/Metadata" in writer._root_object:
                del writer._root_object["/Metadata"]

            with open(temp_clean, "wb") as f_out:
                writer.write(f_out)
        except Exception as e:
            messagebox.showerror("Fehler", f"pypdf-Fehler bei {file}: {e}")
            os.rename(temp_name, file)
            continue

        # 2. if Blista: Ghostscript on temp_clean PDF
        if pdf_format_var.get() == "Blista":
            if shutil.which("gswin64c") is None:
                messagebox.showerror(
                    "Ghostscript nicht gefunden",
                    "Bitte installieren Sie Ghostscript oder fügen Sie es zum PATH hinzu."
                )
                os.rename(temp_name, file)
                os.remove(temp_clean)
                continue

            gs_cmd = [
                "gswin64c",
                "-dPDFA=1",
                "-dBATCH",
                "-dNOPAUSE",
                "-dNOOUTERSAVE",
                "-sProcessColorModel=DeviceCMYK",
                "-sDEVICE=pdfwrite",
                "-dPDFACompatibilityPolicy=1",
                f"-sOutputFile={file}",
                temp_clean
            ]
            try:
                subprocess.run(gs_cmd, check=True)
            except subprocess.CalledProcessError as e:
                messagebox.showerror("Fehler", f"Ghostscript-Fehler bei {file}: {e}")
                os.rename(temp_name, file)
                continue
        else:
            # Nur säubern → temp_clean ins Ziel kopieren
            shutil.move(temp_clean, file)

        pdfcount += 1
        os.remove(temp_name)
        if os.path.exists(temp_clean):
            os.remove(temp_clean)

    revert_button_text(convert_button, meta_button)  
    
    if pdfcount > 0:
        messagebox.showinfo("Erledigt", f"{pdfcount} PDF wurden bearbeitet und gespeichert.")



def show_temp_message(title: str, message: str, seconds: int = 5):
    # Create a new top-level window for the message.
    temp_window = Toplevel()
    temp_window.overrideredirect(True)  # no topic for window
    
    #screen center
    screen_width = temp_window.winfo_screenwidth()
    screen_height = temp_window.winfo_screenheight()

    # position of message
    x = (screen_width // 2) - 150
    y = (screen_height // 2) - 50

    # set window size
    temp_window.geometry(f"300x100+{x}+{y}")
    temp_window.title(title)

    # Label for message
    label = Label(temp_window, text=message, font=("Helvetica", 12), pady=20)
    label.pack()

    # kill message after x seconds
    temp_window.after(seconds * 1000, temp_window.destroy)


def revert_button_text(convert_btn: Button, meta_btn: Button):
    # Reset button to original text
    meta_btn.config(state="active", text = "PDF bearbeiten", bg="#1E40AF", fg="white")
    # Enable the button again to create another batch of PDF files.
    convert_btn.config(state="active", text="PDF aus Word erzeugen", bg="#2563EB", fg="white")

def kill_all_word():
    # Iterate over all running processes
    for proc in psutil.process_iter():
        try:
            # Get process details as a named tuple
            process_info = proc.as_dict(attrs=['pid', 'name'])
            process_name = process_info['name'].lower()
        
            # Check if the process is Microsoft Word
            if 'winword.exe' == process_name:
            # Terminate the process
                process = psutil.Process(process_info['pid'])
                process.terminate()
        except (psutil.NoSuchProcess, psutil.AccessDenied, psutil.ZombieProcess):
                pass


def show_info():
    messagebox.showinfo(
        "Über das Tool",
        "PDF-Tools v2.1\n\n"
        "Beschreibung: \n"
        "- erstellt PDF aus Word-Dateien mittels COM-Schnittstelle\n"
        "- entfernt Metadaten aus bestehenden PDF\n"
        "- wandelt PDF ins Format PDF/A-1a mittels Ghostscript um (Blista)\n\n"
        "Erstellt von Sebastian Buch\n"
        "Kontakt: buc@hems.de"
    )

# Create the main window
root = Tk()
root.grid_columnconfigure(1, weight=1)
root.iconbitmap(resource_path("PDFTools.ico"))

# Set the window title
root.title("PDF-Tools v2.1")

# Make the window not resizeable
root.resizable(0, 0)

# Menu bar with Info
menubar = Menu(root)
root.config(menu=menubar)

help_menu = Menu(menubar, tearoff=0)
help_menu.add_command(label="Info", command=show_info)
menubar.add_cascade(label="About", menu=help_menu)
# -------------------------------

# Place image
pimage = PhotoImage(file=resource_path("hla.png"))
label1 = Label(root, image=pimage)
label1.image = pimage
label1.grid(row=0, column=0, sticky=N, columnspan=2, pady=(0,10))

# PDF Creation Frame
create_frame = Frame(root, relief="ridge", bd=2, bg="#FFFFFF")
create_frame.grid(row=1, column=0, columnspan=2, padx=10, pady=5, sticky="ew")
create_frame.grid_columnconfigure(1, weight=1)

# Frame title
create_title = Label(create_frame, text="PDF aus Word erstellen", font=("Helvetica", 12, "bold"), bg="#FFFFFF")
create_title.grid(row=0, column=0, columnspan=2, pady=(10,10))

# Check if Revisionmode should be disabled and all changes accepted.
disable_track_changes_var = BooleanVar()
disable_track_changes_cb = Checkbutton(create_frame, text="Kommentare löschen,\nNachverfolgung beenden\nund Änderungen annehmen", 
                                    variable=disable_track_changes_var, font=("Helvetica", 10), bg="#FFFFFF")
disable_track_changes_cb.grid(row=1, column=0, columnspan=2, padx=10, pady=5)

format_frame = Frame(create_frame, bg="#FFFFFF")
format_frame.grid(row=2, column=0, columnspan=2, pady=5)

# Add a button to start converting the docx
convert_button = Button(create_frame, text="PDF aus Word erzeugen", width=25, 
                       command=lambda: select_docx_files(convert_button, meta_button, disable_track_changes_var), 
                       font=("Helvetica", 12), bg="#2563EB", fg="white", relief="raised", bd=2)
convert_button.grid(row=3, column=0, columnspan=2, padx=10, pady=(0,15))

# Separator
separator_frame = Frame(root, height=20)
separator_frame.grid(row=2, column=0, columnspan=2, pady=15)
canvas = Canvas(separator_frame, height=2, bg="white", highlightthickness=0)
canvas.pack(fill="x", padx=20)
canvas.create_line(0, 1, 400, 1, fill="#cccccc", width=2) 

# PDF Processing Frame
process_frame = Frame(root, relief="ridge", bd=2, bg="#FFFFFF")
process_frame.grid(row=3, column=0, columnspan=2, padx=10, pady=5, sticky="ew")
process_frame.grid_columnconfigure(1, weight=1)

# Frame title
process_title = Label(process_frame, text="PDF bearbeiten", font=("Helvetica", 12, "bold"), bg="#FFFFFF")
process_title.grid(row=0, column=0, columnspan=2, pady=(10,5))

pdf_mode_frame = Frame(process_frame, bg="#FFFFFF")
pdf_format_var = StringVar(value="Nur säubern")
pdf_format_label = Label(pdf_mode_frame, text="PDF‑Modus:", font=("Helvetica", 11), bg="#FFFFFF")
pdf_format_label.pack(side="left", padx=(0,5))

pdf_format_menu = OptionMenu(pdf_mode_frame, pdf_format_var, "Nur säubern", "Blista")
pdf_format_menu.config(font=("Helvetica", 11), width=10, bg="white")
pdf_format_menu.pack(side="left")
pdf_mode_frame.grid(row=1, column=0, columnspan=2, pady=(0,10))

meta_button = Button(process_frame, text="PDF bearbeiten", width=25,
                    command=lambda: pdf_edit(meta_button, pdf_format_var),
                    font=("Helvetica", 12), bg="#1E40AF", fg="white", relief="raised", bd=2)
meta_button.grid(row=2, column=0, columnspan=2, padx=10, pady=(0,15))

# Run the Tkinter event loop
root.mainloop()
