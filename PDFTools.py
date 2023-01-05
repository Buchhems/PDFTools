import os
from tkinter import BOTTOM, Button, Label, Tk, Toplevel, filedialog, messagebox, PhotoImage
import comtypes.client
from pypdf import PdfReader, PdfWriter

def docx_to_pdf(docx_filename, pdf_filename):
    # Create a COM object for Word
    word = comtypes.client.CreateObject('Word.Application')

    # Read the Word Document.
    try:
        doc = word.Documents.Open(f'"{docx_filename}"')
    except comtypes.COMError as e:
        messagebox.showerror(title = "Fehler", message = f'Kann {docx_filename} nicht finden.\nWord-Datei bereits geöffnet oder Leerzeichen im Namen?')
        word.Quit()
        return

    # Write the pdf
    try:
        doc.SaveAs(pdf_filename, FileFormat=17)
    except comtypes.COMError:
        messagebox.showerror(title = "Fehler", message = f'{pdf_filename} existiert bereits.\nÜbersprungen...')
        doc.Close()
        word.Quit()
        return

    # necessary to quit word instances after a correct run of the pdf generator
    doc.Close()
    word.Quit()

def select_docx_files(convert_button):

    # Disable the button for unintentional clicks of certain users ;)
    convert_button.config(state="disabled", text="Erzeuge PDFs...")
    
    # close all instances of word to not create a mess...
    messagebox.showwarning(title = "Bitte alle Word-Fenster schließen!", message = "Bitte schließen Sie alle Word-Fenster und klicken dann auf OK. Dies ist für einen sauberen Ablauf notwendig.\n\nHinweis:\nIm Extremfall schließen Sie Word-Instanzen per Taskmanager.")

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
    
    show_temp_message('Erledigt', 'Die PDFs wurden erzeugt.')
    #enable the button again to create another batch of PDF files.
    convert_button.config(state="active", text ="Word > PDF")
        
def remove_metadata(meta_button):

    meta_button.config(state="disabled", text="Entferne Metadaten...")

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
            messagebox.showerror('Fehler', 'Die Datei "{}" wird von einem anderen Programm verwendet und wird daher übersprungen.'.format(file))
        except FileExistsError:
            messagebox.showerror('Fehler', 'Die Datei "{}" gibt es bereits. Die Ursprungsdatei wurde daher nicht umbenannt. Vorgang abgebrochen.'.format(new_name))
            continue

        # Open the PDF in read-binary mode
        with open(new_name, 'rb') as file:
            # Create a PDF object
            pdf = PdfReader(new_name)

            # Create a PDF object to write the output to
            output_pdf = PdfWriter()

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

    # Reset button to original text
    meta_button.config(state="active", text ="Metadaten aus PDF löschen")

    show_temp_message('Erledigt', 'Die Metadaten der PDFs wurden gelöscht.\nDie Ursprungsdateien wurden in\n DATEINAME_orig.pdf umbenannt.')

def show_temp_message(title, message, seconds=5):
    # Create a new top-level window for the message.
    window = Toplevel()
    window.title(title)
   
    # Create a label for the message.
    label = Label(window, text=message, font=("Helvetica", 12))
    label.pack()
    
    # Close the window after a certain number of seconds.
    window.after_idle(lambda: window.after(seconds * 1000, window.destroy))
            
# Create the main window
window = Tk()

# Set the window title
window.title("PDF-Tools v1.2 (buc @ hems.de)")

# Set the window size
window.geometry("560x240")

pimage = PhotoImage(file="./hla.png")

label1 = Label(image=pimage)
label1.image = pimage

#position image
label1.place(x=0, y= 0)

# Add a button to start cleaning the PDFs
meta_button = Button(text="Metadaten aus PDF löschen", command=lambda: remove_metadata(meta_button), font=("Helvetica", 14))
meta_button.pack(side=BOTTOM, pady=10)

# Add a button to start converting the docx
convert_button = Button(text ="Word => PDF", command=lambda: select_docx_files(convert_button), font=("Helvetica", 14))
convert_button.pack(side=BOTTOM)

# Run the Tkinter event loop
window.mainloop()