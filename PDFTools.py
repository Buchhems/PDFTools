import os
import tkinter as tk
from tkinter import *
from tkinter import filedialog
from tkinter import messagebox
from PIL import Image, ImageTk
import os
import comtypes.client
import pypdf

def docx_to_pdf(docx_filename, pdf_filename):
    # Create a COM object for Word.
    word = comtypes.client.CreateObject('Word.Application')
    try:
        doc = word.Documents.Open(docx_filename)
    except comtypes.COMError:
        messagebox.showerror(title = "Fehler", message = f'Kann {docx_filename} nicht finden.\nWord-Instanz, die die Datei geöffnet hat oder Leerzeichen im Namen?')
        doc.Close()
        word.Quit()
        return

    try:
        doc.SaveAs(pdf_filename, FileFormat=17)
    except comtypes.COMError:
        messagebox.showerror(title = "Fehler", message = f'{pdf_filename} existiert bereits.\nÜbersprungen...')

    doc.Close()
    word.Quit()

def select_docx_files(select_button):

    # Disable the button for unintentional clicks of certain users ;)
    select_button.config(state="disabled", text="Erzeuge PDFs...")
    
    # close all instances of word to not create a mess...
    messagebox.showwarning(title = "Bitte alle Word-Fenster schließen!", message = "Bitte schließen Sie alle Word-Fenster und klicken dann auf OK. Dies ist für einen sauberen Ablauf notwendig.\n\nHinweis:\nIm Extremfall schließen Sie diese per Taskmanager.")

    # Open a file selection dialog and get the selected files.
    docx_filenames = filedialog.askopenfilenames(title='Word-Dateien zur Konvertierung auswählen', filetypes=[('Word Dokumente', '*.docx')]) 
              
    # Convert each Word document to a PDF.
    for docx_filename in docx_filenames:
        # Get the base and extension of the file.
        base, ext = os.path.splitext(docx_filename)

        # Create the PDF filename.
        pdf_filename = base + '.pdf'
                 
        # Convert the Word document to a PDF.
        docx_to_pdf(docx_filename, pdf_filename)
    
    tk.messagebox.showinfo('Erledigt', 'Die PDFs wurden erzeugt.')
    #enable the button again to create another batch of PDF files.
    select_button.config(state="active", text ="Word > PDF")
        
def remove_metadata():
    # Open the PDF files in read-binary mode
    files = filedialog.askopenfilenames(title='PDFs zum Entfernen von Metadaten auswählen', filetypes=[('PDF Dokumente', '*.pdf')])

    # Loop through all selected PDFs
    for file in files:
        # Get the original file name and extension
        name, extension = os.path.splitext(file)
        # Generate the new name for the original PDF
        new_name = os.path.join(name + '_orig' + extension)
        # Rename the original PDF
        try:
            os.rename(file, new_name)
        except PermissionError:
            # If the file is in use, skip it and move on to the next file
            tk.messagebox.showerror('Fehler', 'Die Datei "{}" wird von einem anderen Programm verwendet und wird daher übersprungen.'.format(file))
        except FileExistsError:
            tk.messagebox.showerror('Fehler', 'Die Datei "{}" gibt es bereits. Die Originaldatei wurde daher nicht umbenannt. Vorgang abgebrochen.'.format(new_name))
            continue

        # Open the PDF in read-binary mode
        with open(new_name, 'rb') as file:
            # Create a PDF object
            pdf = pypdf.PdfReader(new_name)

            # Create a PDF object to write the output to
            output_pdf = pypdf.PdfWriter()

            # Iterate through all pages in the PDF
            for page in pdf.pages:
                output_pdf.add_page(page)
            
            output_pdf.add_metadata(
            {
                "/Creator": "",
                "/Producer": "",
                "/Author": "",
                "/Title": "",
                "/Subject": "",
                "/Keywords": "",
                "/CreationDate": "",
                "/ModDate": "",
            }
            )
            
            output_file = name + extension
            with open(output_file, 'wb') as f:
                output_pdf.write(f)

    tk.messagebox.showinfo('Erledigt', 'Die Metadaten der PDFs wurden gesäubert. Die originalen Dateien wurden in DATEINAME_orig.pdf umbenannt.')

def show_temp_message(title, message, seconds=3):
    # Create a new top-level window for the message.
    window = Toplevel()
    window.title(title)
    
    # Create a label for the message.
    label = Label(window, text=message, font=("Arial", 25))
    label.pack()
    
    # Close the window after a certain number of seconds.
    window.after_idle(lambda: window.after(seconds * 1000, window.destroy))
            
# Create the main window
window = tk.Tk()

# Set the window title
window.title("PDF-Tools v1.2 (buc@hems.de)")

# Set the window size
window.geometry("560x240")

#load picture
image1 = Image.open (os.path.dirname(__file__) +"\hla.png")
pimage = ImageTk.PhotoImage(image1)

label1 = tk.Label(image=pimage)
label1.image = pimage

#position image
label1.place(x=0, y= 0)

# Add a button to start cleaning the PDFs
button = tk.Button(text="Metadaten aus PDF löschen", command=remove_metadata, font=("Helvetica", 14))
button.pack(side=BOTTOM, pady=10)

# Add a button to start converting the docx
select_button = tk.Button(text ="Word => PDF", command=lambda: select_docx_files(select_button), font=("Helvetica", 14))
select_button.pack(side=BOTTOM)

# Run the Tkinter event loop
window.mainloop()