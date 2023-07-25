from docx import Document
from tkinter import *
from tkinter import filedialog
from tkinter import messagebox
import customtkinter as ct
import pandas as pd
def pdf():
    root = ct.CTkToplevel()
    root.grab_set()
    root.title('MULTIPLEXER')
    width = root.winfo_screenwidth()
    height = root.winfo_screenheight()
    root.geometry("{}x{}".format(width, height))
    root.config(background='#FFB6C1')

    global file_list
    file_list = []

    def file_open():
        filename = filedialog.askopenfilename(
            title='Select a file',
            filetypes=(('PDF', '*.pdf'),))

        file_list.append(filename)
        for i in range(0, len(file_list)):
            filename_lbl = ct.CTkLabel(root, text=file_list[i],fg_color='#FFB6C1',text_color='black').place(x=550, y=(i * 20) + 450)

    def file_print():
        print(file_list)
        f = ''
        f = combine_filename.get()
        if f == "":
            messagebox.showerror("showerror", "Enter file name !!!")
        elif len(file_list) < 2:
            messagebox.showerror("showerror", "Select 2 file to combine !!!")
        else:
            new = Pdf.new()
            for i in file_list:
                with Pdf.open(i.strip()) as pdf:
                    new.pages.extend(pdf.pages)
            new.save(f + '.pdf')
            messagebox.showinfo("showinfo", "Success")
    combine_filename = StringVar()
    btn_select = ct.CTkButton(root, text="Select file",bg_color='#FFB6C1', fg_color="#800080", hover_color="green", width=200, height=50,
                              corner_radius=40, command=file_open).place(x=650, y=250)
    filename_entry = ct.CTkEntry(root, textvariable=combine_filename,text_color='black', bg_color='#FFB6C1',corner_radius=20,fg_color='white', width=400).place(x=550, y=335)
    btn_print = ct.CTkButton(root, text="combine", fg_color="#800080", bg_color='#FFB6C1', hover_color="green", command=file_print,
                             width=200, height=50, corner_radius=40).place(x=650, y=400)
    label = ct.CTkLabel(root, text='MULTIPLEXER', fg_color='#FFB6C1', text_color='black', font=('times', 50)).place(x=580,
                                                                                                                 y=100)
    root.mainloop()


def doc():
    root = ct.CTkToplevel()
    root.title('MULTIPLEXER')
    height = root.winfo_screenheight()
    width = root.winfo_screenwidth()
    root.geometry("{}x{}".format(width, height))
    root.grab_set()

    global ety
    global file_lst
    file_lst = []

    def selectdoc():
        doc = filedialog.askopenfilename(title='select a document',
                                         filetypes=(('Word Documents', '*.docx'), ('All files', '*.*')))
        file_lst.append(doc)
        for i in range(0, len(file_lst)):
            lab = ct.CTkLabel(root, text=file_lst[i],bg_color='#FFB6C1',text_color='black').place(x=550, y=(i * 20) + 450)

    def combine():
        f = ''
        f = combinefile.get()
        if f == '':
            messagebox.showerror('showerror', 'enter the file name')
        elif len(file_lst) < 2:
            messagebox.showerror('showerror', 'select more than one document')
        else:
            merged_document = Document()
            for path in file_lst:
                doc = Document(path)
                for element in doc.element.body:
                    merged_document.element.body.append(element)
            merged_document.save("{}.docx".format(f))
            messagebox.showinfo("showinfo", "Success")
    root.config(background='#FFB6C1')
    combinefile = StringVar()
    btn_select = ct.CTkButton(root, text="Select Document",bg_color='#FFB6C1', fg_color="#800080", hover_color="green", width=200,
                              height=50,
                              corner_radius=40, command=selectdoc).place(x=650, y=250)
    filename_entry = ct.CTkEntry(root, textvariable=combinefile,text_color='black', bg_color='#FFB6C1',corner_radius=20,fg_color='white', width=400).place(x=550, y=335)
    btn_print = ct.CTkButton(root, text="combine",bg_color='#FFB6C1', fg_color="#800080", hover_color="green", command=combine,
                             width=200, height=50, corner_radius=40).place(x=650, y=400)

    label = ct.CTkLabel(root, text='MULTIPLEXER', fg_color='#FFB6C1', text_color='black', font=('times', 50)).place(x=580,
                                                                                                                 y=100)
    root.mainloop()
def excel():
    # Load the Excel files
    global excel_files
    global dataframes
    dataframes = []
    excel_files = []
    root = ct.CTkToplevel()
    root.grab_set()
    root.title('MULTIPLEXER')
    root.config(background='#FFB6C1')
    width = root.winfo_screenwidth()
    height = root.winfo_screenheight()
    root.geometry("{}x{}".format(width, height))

    def file_open():

        filename = filedialog.askopenfilename(
            title='Select a file',
            filetypes=(('EXCEL', '*.xlsx'),))
        excel_files.append(filename)
        for i in range(0, len(excel_files)):
            filename_lbl = ct.CTkLabel(root,text=excel_files[i],bg_color='#FFB6C1',text_color='black').place(x=550, y=(i * 20) + 450)

    def file_print():
        print(excel_files)
        f = ''
        f = combine_filename.get()
        if f == "":
            messagebox.showerror("showerror", "Enter file name !!!")
        elif len(excel_files) < 2:
            messagebox.showerror("showerror", "Select 2 file to combine !!!")
        else:
            merged_data = pd.DataFrame()
            for file in excel_files:
                # Load the Excel file into a pandas DataFrame
                excel_data = pd.read_excel(file)
                dataframes.append(excel_data)
            merged_data = pd.concat(dataframes, ignore_index=True)
            merged_data.to_excel(f + '.xlsx', index=False)
            messagebox.showinfo("showinfo", "Success")

    combine_filename = StringVar()
    btn_select = ct.CTkButton(root, text="Select file",bg_color='#FFB6C1', fg_color="#800080", hover_color="green", width=200, height=50,
                              corner_radius=40, command=file_open).place(x=650, y=250)
    filename_entry = ct.CTkEntry(root, textvariable=combine_filename,text_color='black', bg_color='#FFB6C1',corner_radius=20,fg_color='white', width=400).place(x=550, y=335)
    btn_print = ct.CTkButton(root, text="combine",bg_color='#FFB6C1', fg_color="#800080", hover_color="green", command=file_print,
                             width=200, height=50, corner_radius=40).place(x=650, y=400)
    label = ct.CTkLabel(root, text='MULTIPLEXER', fg_color='#FFB6C1', text_color='black', font=('times', 50)).place(
        x=580,
        y=100)
    root.mainloop()

from tkinter import Tk, Button, Label, filedialog, messagebox, Entry
from pikepdf import Pdf
import customtkinter as ct
def spl():
    global input_path
    global output_dir
    def split_pdf():
        # Open a file dialog to select the input PDF file
        global input_path
        input_path= filedialog.askopenfilename(filetypes=[('PDF Files', '*.pdf')])
        if not input_path:
            return

    # Create a directory dialog to select the output directory
        global output_dir
        output_dir = filedialog.askdirectory()
        if not output_dir:
            return
    def comb():
        try:
            # Get the pages to split from the entry field
            pages_to_split = page_entry.get()
            page_numbers = [int(p) for p in pages_to_split.split(',')]

            with Pdf.open(input_path.strip()) as pdf:
                total_pages = len(pdf.pages)
                for page_number, page in enumerate(pdf.pages, 1):
                    if page_number in page_numbers:
                        output_path = f"{output_dir}/page_{page_number}.pdf"
                        output_pdf = Pdf.new()
                        output_pdf.pages.append(page)
                        output_pdf.save(output_path)

            messagebox.showinfo("Success", f"PDF pages {pages_to_split} have been split.")
        except Exception as e:
            messagebox.showerror("Error", str(e))


# Create the main Tkinter window
    window = ct.CTkToplevel()
    h=window.winfo_screenheight()
    w=window.winfo_screenwidth()
    window.grab_set()
    window.geometry('{}x{}'.format(w,h))
    window.title("MULTIPLEXER")
    window.config(background='#FFB6C1')
    # Create a button to select and split the PDF file
    split_button = ct.CTkButton(window, text="Split PDF",fg_color="#800080",bg_color='#FFB6C1', hover_color="green", width=200, height=50,
                              corner_radius=40, command=split_pdf)
    split_button.place(x=650, y=250)
    label = ct.CTkLabel(window, text='MULTIPLEXER', fg_color='#FFB6C1', text_color='black', font=('times', 50)).place(
        x=580,
        y=100)

# Create a label and entry field to input the page numbers to split
    page_label = ct.CTkLabel(window, text="Pages to Split (e.g., 1,3,5):",fg_color='#FFB6C1',text_color='black')
    page_label.place(x=680, y=310)
    btn_print = ct.CTkButton(window, text="split",bg_color='#FFB6C1', fg_color="#800080", hover_color="green",
                         width=200, height=50, corner_radius=40,command=comb).place(x=650, y=400)
    page_entry = ct.CTkEntry(window,width=400,bg_color='#FFB6C1',fg_color='white',text_color='black',corner_radius=20)
    page_entry.place(x=550, y=335)

# Run the Tkinter event loop
    window.mainloop()

s = ct.CTk()
s.title('MULTIPLEXER')
height = s.winfo_screenheight()
width = s.winfo_screenwidth()
s.geometry("{}x{}".format(width, height))
s.config(background="#FFB6C1")
c=StringVar()
btn_select = ct.CTkButton(s, text="PDF MERGER", bg_color='#FFB6C1',fg_color="#800080",hover_color="green", width=300, height=50,corner_radius=40 ,command=pdf).place(x=600, y=300)
btn_select2 = ct.CTkButton(s, text="DOCUMENT MERGER", bg_color='#FFB6C1', fg_color="#800080",hover_color="green", width=300, height=50,corner_radius=40 ,command=doc).place(x=300 * 2, y=400)
btn_select3 = ct.CTkButton(s, text="EXCEL MERGER", bg_color='#FFB6C1', fg_color="#800080",hover_color="green", width=300, height=50,corner_radius=40 ,command=excel).place(x=300*2, y=500)
btn_select4 = ct.CTkButton(s, text="PDF SPLITER", bg_color='#FFB6C1', fg_color="#800080",hover_color="green", width=300, height=50,corner_radius=40 ,command=spl).place(x=300*2, y=600)

label=ct.CTkLabel(s,text='MULTIPLEXER', fg_color='#FFB6C1',text_color='black',font=('times',50)).place(x=580,y=100)
s.mainloop()
