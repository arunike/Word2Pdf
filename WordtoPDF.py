from logging import NOTSET
from tkinter import *
from tkinter import messagebox
from tkinter import filedialog
import os, sys
import docx2pdf
from tkinter import Tk, Label

def WordToPDF():
    window = Tk()
    window.resizable(False, False)
    window.geometry("600x400")
    window.title("PDF Converter")
    window.config(background = "#6495ed")

    paths = []
    files = []

    basedir = os.path.dirname(os.path.abspath(__file__))

    def chooseFiles():
        global filepath
        global filename
        global frame
    
        filepath = filedialog.askopenfilename(initialdir = "../", title = "Select File :", 
        filetypes = (("Available Files","* docx txt"), ("All Files", "*.*")))
        filename = os.path.basename(filepath)
        files.append(filepath)
        frame = Frame(window, bg = "#ffffff")
        frame.place(relx = 0.38, rely = 0.4, width = 355, height = 135,)

        yscrollbar = Scrollbar(frame, orient = VERTICAL)
        yscrollbar.pack(side = RIGHT, fill=  Y)
        xscrollbar = Scrollbar(frame, orient = HORIZONTAL)
        xscrollbar.pack(side = BOTTOM, fill = X)

        fbox = Canvas(frame, width = 355, height = 135, scrollregion = (0, 0, 500, 500), xscrollcommand = xscrollbar.set, yscrollcommand = yscrollbar.set, bg = 'white')
        fbox.pack()
        
        yscrollbar.config(command = fbox.yview)
        xscrollbar.config(command = fbox.xview)

        if filename.lower().endswith(('.docx')):
            global fileLogo
            fileLogo = PhotoImage(file = os.path.join(basedir, "img/file.png"))

            for i in files: 
                selectedFile = Label(frame,
                                        text = filename,
                                        bg = "#fff",
                                        fg = "#828282",
                                        image = fileLogo,
                                        compound = TOP       
                                    )
                selectedFile.place(relx = 0,
                                    rely = 0.01,
                                    width = 350,
                                    height = 130,
                                )
    
    def savePath():
        global sPath
        sPath = filedialog.askdirectory()

        paths.append(sPath)

        destination = Label(canvas, 
                                    text=  "Save to : " + sPath, 
                                    bg = "#2f4a7c", 
                                    fg = "#fff", 
                                    padx = 6.5,
                                    image = saveTo,
                                    compound = LEFT
                                )
        destination.place(rely = 0.85, relx = 0.01)

    def fileToPdf():
        try:
            filepath
        except NameError:
            messagebox.showwarning("Warning", "You must choose a file first")
        else:
            if(filename.lower().endswith('.docx')):
                docxToPdf()

    def docxToPdf():
        nn = entry.get()
        name = nn

        try:
            sPath
        except NameError:
            messagebox.showwarning("Warning", "New PDF file need a new path to save")
        else:
            with open(f"{sPath}/{name}.pdf", "wb") as f:
                docx2pdf.convert(filepath, f"{sPath}/{name}.pdf")

            messagebox.showinfo("ALl Set","Congrats! Your PDF is ready")

    def removefiles():
        print ("Before", files)

        files.clear()

        for widgets in frame.winfo_children():
            widgets.destroy()

        print ("After", files)

    logo = PhotoImage(file = os.path.join(basedir, 'img/logo.png'))
    header = Label(window,
                        text = "Convert Your Files To .PDF",
                        fg = "#111111" ,font=("Arial", 20, 'bold'),
                        bg = "#6495ed",
                        pady = 5,
                        image = logo,
                        compound = TOP,
                        )
    header.place(relx = 0.27, rely = 0.05)

    canvas = Canvas(window,
                        width = 580, 
                        height = 230, 
                        bg = "#2f4a7c"
                    )
    canvas.place(rely = 0.35, relx = 0.01)

    chooseFile = Button(canvas, 
                            text = "Choose File",
                            bg = 'white', fg='black',
                            padx = 50,
                            pady = 5,
                            command = chooseFiles
                        )
    chooseFile.place(relx = 0.025, rely = 0.1)

    saveTo = PhotoImage(file = os.path.join(basedir, "img/saveTo.png"))
    savePath = Button(canvas, 
                        text = "Save Destination",
                        bg = 'white', fg = 'black',
                        padx = 35,
                        pady = 5,
                        command = savePath,
                    )
    savePath.place(relx = 0.025, rely = 0.3)

    convert = Button(canvas, 
                        text = "Convert",
                        bg = 'white', fg='black',
                        padx = 62,
                        pady = 5,
                        command = fileToPdf
                    )
    convert.place(relx = 0.025, rely = 0.5)

    remove = Button(canvas, 
                        text = "Remove",
                        bg = 'white', fg='black',
                        padx = 62,
                        pady = 5,
                        command = removefiles
                    )
    remove.place(relx = 0.025, rely = 0.7)

    preview = Label(canvas,
                        text = "Drag and Drop your files",
                        bg = "#fff",
                        fg = "#828282",       
                    )
    preview.place(relx = 0.38,
                    rely = 0.1,
                    width = 350,
                    height = 130,
                )

    newFileAs = Label(canvas, 
                        text = "Save File As : ", 
                        bg = "#2f4a7c",
                        fg = "#fff"
                    )
    newFileAs.place(relx = 0.38, rely = 0.7)

    entry = Entry(width = 20, background = "#3a4c77", fg = "#fff",border = None)
    entry.place(relx = 0.55, rely = 0.75)
    
    window.mainloop()
    sys.exit()

WordToPDF()