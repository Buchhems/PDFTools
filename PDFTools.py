import os
import sys
import psutil
from tkinter import (BOTTOM, Button, Label, PhotoImage, Tk, Toplevel,
                     filedialog, messagebox)

import comtypes.client
from pypdf import PdfReader, PdfWriter


def docx_to_pdf(docx_filename, pdf_filename):
    # Create a COM object for Word
    word = comtypes.client.CreateObject('Word.Application')

    # Read the Word Document.
    try:
        docx_filename = os.path.abspath(docx_filename)
        doc = word.Documents.Open(f'"{docx_filename}"')

    except comtypes.COMError as e:
        messagebox.showerror(title = "Fehler", message = "Die Datei " + docx_filename + " kann nicht geöffnet werden.")
        word.Quit()
        return 0

    # Write the pdf
    try:
        pdf_filename = os.path.abspath(pdf_filename)
        doc.SaveAs(pdf_filename, FileFormat=17)
      
    except comtypes.COMError as e:
        messagebox.showerror(title = "Fehler", message = "Die Datei " + pdf_filename + " kann nicht angelegt werden.")
        doc.Close()
        word.Quit()
        return 0

    # necessary to quit word instances after a correct run of the pdf generator
    doc.Close()
    word.Quit()
    return 1

def select_docx_files(convert_button):
    #for message later
    pdfcount = 0

    # Disable the button for unintentional clicks of certain users ;)
    convert_button.config(state="disabled", text="...einen Moment bitte...")
    
    # close all instances of word to not create a mess...
    messagebox.showwarning(title = "Word-Fenster schließen!", 
                           message = "Schließen Sie alle Word-Fenster\nund klicken Sie dann auf OK.")
 
    # kill all word instances (user has been warned)
    killallword()

    # Open a file selection dialog and get the selected files.
    docx_filenames = filedialog.askopenfilenames(title='Word-Dateien zur Erzeugung von PDF auswählen', filetypes=[('Word Dokumente', '*.docx')]) 

    # Convert each Word document to a PDF.
    for docx_filename in docx_filenames:
        # Get the base and extension of the file.
        base, ext = os.path.splitext(docx_filename)

        # Create the PDF filename.
        pdf_filename = base + '.pdf'

        # Convert the Word document to a PDF.
        count = docx_to_pdf(docx_filename, pdf_filename)
        pdfcount = pdfcount + count

        #print(pdfcount)
        convert_button.config(state="disabled", text="Erzeugte PDF: " + str(pdfcount)+".")
        window.update()
    
    if pdfcount > 0:
        show_temp_message('Erledigt', 'Es wurde(n)\n' + str(pdfcount) + ' PDF erzeugt.')
    
    revert_button_text()
        
def remove_metadata(meta_button):
    pdfcount = 0

    meta_button.config(state="disabled", text="Entferne gerade Metadaten...")

    # Open the PDF files in read-binary mode
    files = filedialog.askopenfilenames(title='PDF auswählen', filetypes=[('PDF Dokumente', '*.pdf')])

    # Loop through all selected PDFs
    for file in files:
        # Get the original file name and extension
        name, extension = os.path.splitext(file)
        # Generate the temp name for the original PDF
        temp_name = os.path.join(name + "_todo" + extension)
        # Rename the original PDF
        try:
            os.rename(file, temp_name)
        except PermissionError:
            # If the file is in use, skip it and move on to the next file
            messagebox.showerror('Fehler', 'Die Datei "{}" wird von einem anderen Programm verwendet und wird daher übersprungen.'.format(file))
            break
        except FileExistsError:
            messagebox.showerror('Fehler', 'Die temporäre Datei "{}" gibt es bereits. Datei übersprungen. Bitte löschen Sie diese Datei.'.format(temp_name))
            break

        # Open the PDF in read-binary mode
        with open(temp_name, 'rb') as file:
            # Create a PDF object
            pdf = PdfReader(temp_name)

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
                pdfcount+=1
            
        os.remove(temp_name)
            
    revert_button_text()

    if pdfcount >0:
        show_temp_message("erledigt...", "Die Metadaten von\n" + str(pdfcount) + " PDF wurden entfernt.")

def show_temp_message(title, message, seconds=5):
    # Create a new top-level window for the message.
    window = Toplevel()
    window.overrideredirect(True)
    #window.geometry("300x200")
    window.title(title)
   

    # Create a label for the message.
    label = Label(window, text=message, font=("Helvetica", 50))
    label.pack()
    
    # Close the window after a certain number of seconds.
    window.after_idle(lambda: window.after(seconds * 1000, window.destroy))

def revert_button_text():
    # Reset button to original text
    meta_button.config(state="active", text = "Metadaten aus PDF entfernen")
    #enable the button again to create another batch of PDF files.
    convert_button.config(state="active", text ="PDF aus DOCX erzeugen")

def killallword():
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
                #time.sleep(2)
        except (psutil.NoSuchProcess, psutil.AccessDenied, psutil.ZombieProcess):
                pass

def resource_path(relative_path):
    """ Get absolute path to resource, works for dev and for PyInstaller """
    try:
        # PyInstaller creates a temp folder and stores path in _MEIPASS
        base_path = sys._MEIPASS
    except Exception:
        base_path = os.path.abspath(".")

    return os.path.join(base_path, relative_path)


# Create the main window
window = Tk()
window.iconbitmap(resource_path("PDFTools.ico"))

# Set the window title
window.title("PDF-Tools v1.4 (buc @ hems.de)")

# Set the window size
window.geometry("560x240")

pimage = PhotoImage(file=resource_path("hla.png"))

label1 = Label(image=pimage)
label1.image = pimage

#position image
label1.place(x=0, y= 0)

# Add a button to start cleaning the PDFs
meta_button = Button(text="Metadaten aus PDF entfernen", command=lambda: remove_metadata(meta_button), font=("Helvetica", 14))
meta_button.pack(side=BOTTOM, pady=10)

# Add a button to start converting the docx
convert_button = Button(text ="PDF aus DOCX erzeugen", command=lambda: select_docx_files(convert_button), font=("Helvetica", 14))
convert_button.pack(side=BOTTOM)

# Run the Tkinter event loop
window.mainloop()