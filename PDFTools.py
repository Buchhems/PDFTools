import os
import sys
import time
import tkinter
import psutil
from tkinter import (messagebox, filedialog)
import comtypes.client
from pypdf import PdfReader, PdfWriter


def docx_to_pdf(docx_filename, pdf_filename, disable_track_changes_var):
    # Create a COM object for Word
    word = comtypes.client.CreateObject('Word.Application')

    # Read the Word Document.
    try:
        docx_filename = os.path.abspath(docx_filename)
        doc = word.Documents.Open(f'"{docx_filename}"')

    except comtypes.COMError as e:
        messagebox.showerror(title = "Fehler", message = "Die Datei " + docx_filename + " kann nicht geöffnet werden.\nPDF gerade geöffnet?")
        word.Quit()
        return 0

    #lookup if RevisionMode should be disabled or not
    if disable_track_changes_var.get():
            # Iterate over all comments and delete them
        for comment in doc.Comments:
                #print(comment.Range.Text)
                # Select the entire comment thread (including any replies) using the Range.Start and Range.End properties
            range_start = comment.Scope.Start
            range_end = comment.Scope.End
            for reply in comment.Replies:
                range_start = min(range_start, reply.Scope.Start)
                range_end = max(range_end, reply.Scope.End)
                # Delete the comment thread using the Range.Delete() method
            doc.Range(range_start, range_end).Delete()
            # Only Disable all TrackRevisions based on the user's choice
            doc.TrackRevisions = False
            # Accept all revisions
            doc.AcceptAllRevisions()

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

def select_docx_files(convert_button, enable_track_changes_cb):
    #for message later
    pdfcount = 0

    # Disable the button for unintentional clicks of certain users ;)
    convert_button.config(state="disabled", text="... einen Moment bitte ...")
    
    # close all instances of word to not create a mess...
    messagebox.showwarning(title = "Word-Fenster schließen!", 
                           message = "Schließen Sie alle Word-Fenster\nund klicken Sie dann auf OK.")
 
    # kill all word instances (user has been warned)
    killallword()
    time.sleep(3)

    # Open a file selection dialog and get the selected files.
    docx_filenames = tkinter.filedialog.askopenfilenames(title='Word-Dateien zur Erzeugung von PDF auswählen', filetypes=[('Word Dokumente', '*.docx')]) 

    # Convert each Word document to a PDF.
    for docx_filename in docx_filenames:
        # Get the base and extension of the file.
        base, ext = os.path.splitext(docx_filename)

        # Create the PDF filename.
        pdf_filename = base + '.pdf'

        # Convert the Word document to a PDF.
        count = docx_to_pdf(docx_filename, pdf_filename, enable_track_changes_cb)
        pdfcount = pdfcount + count

        #print(pdfcount)
        convert_button.config(state="disabled", text="Erzeugte PDF: " + str(pdfcount))
        root.update()
    
    if pdfcount > 0:
        show_temp_message('Erledigt', 'Es wurde(n)\n' + str(pdfcount) + ' PDF erzeugt.')
    
    revert_button_text()
        
def remove_metadata(meta_button):
    pdfcount = 0

    meta_button.config(state="disabled", text="Entferne gerade Metadaten...")

    # Open the PDF files in read-binary mode
    files = tkinter.filedialog.askopenfilenames(title='PDF auswählen', filetypes=[('PDF Dokumente', '*.pdf')])

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
        show_temp_message("erledigt...", "Die Metadaten\nvon " + str(pdfcount) + " PDF\nwurden entfernt.")

def show_temp_message(title, message, seconds=5):
    # Create a new top-level window for the message.
    root = tkinter.Toplevel()
    root.overrideredirect(True)
    #window.geometry("300x200")
    root.title(title)

    # get the screen width and height
    screen_width = root.winfo_screenwidth()
    screen_height = root.winfo_screenheight()

    # calculate the x and y coordinates of the message box
    x = (screen_width // 2) - 235
    y = (screen_height // 2) - 95

    # Create a new top-level window for the message.
    
    root.title(title)
    root.geometry(f"+{x}+{y}")  
   
    # Create a label for the message.
    label = tkinter.Label(root, text=message, font=("Helvetica", 50))
    label.pack()
    
    # Close the window after a certain number of seconds.
    root.after_idle(lambda: root.after(seconds * 1000, root.destroy))

def revert_button_text():
    # Reset button to original text
    meta_button.config(state="active", text = "Metadaten aus PDF entfernen")
    #enable the button again to create another batch of PDF files.
    convert_button.config(state="active", text ="PDF aus Docx erzeugen")

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
root = tkinter.Tk()
root.iconbitmap(resource_path("PDFTools.ico"))

# Set the window title
root.title("PDF-Tools v1.5 (buc @ hems.de)")

# Set the window size
#root.geometry("430x250")
#make the window not resizeable
root.resizable(0, 0)

#place image
pimage = tkinter.PhotoImage(file=resource_path("hla.png"))
label1 = tkinter.Label(image=pimage)
label1.image = pimage
label1.grid(row=0, column=0, sticky=tkinter.N,columnspan=2)

# Add a button to start converting the docx
convert_button = tkinter.Button(root, text ="PDF aus Docx erzeugen", width = 20, command=lambda: select_docx_files(convert_button, enable_track_changes_var), font=("Helvetica", 14))
convert_button.grid(row=1, column=1, padx=5, pady=5)

# check if Revisionmode should be disabled and all changes accepted.
enable_track_changes_var = tkinter.BooleanVar()
enable_track_changes_cb = tkinter.Checkbutton(root, text="Evtl. Kommentare löschen,\nNachverfolgung beenden\nund Änderungen annehmen", variable=enable_track_changes_var)
enable_track_changes_cb.grid(row=1, column=0, padx=5, pady=5)

canvas = tkinter.Canvas(root, height=1)
canvas.create_line(2, 2, 500, 2, dash=(4,2))
canvas.grid(row=2, column=0, columnspan=2, padx=5, pady=5)

# Add a button to start cleaning the PDFs
meta_button = tkinter.Button(root, text="Metadaten aus PDF entfernen", command=lambda: remove_metadata(meta_button), font=("Helvetica", 14))
meta_button.grid(row=3, column=0, columnspan=2, padx=5, pady=10)

# Run the Tkinter event loop
root.mainloop()