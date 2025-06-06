# Import Module
import tkinter as tk
from tkinter import *
from tkinter import ttk, messagebox
from fpdf import FPDF
from PIL import Image, ImageWin
from pdf2image import convert_from_path
from PyPDF2 import PdfWriter, PdfReader
import qrcode
import os
import csv
import win32print
import win32ui


tag_dict = {
        "Container":"C",
        "Shelf":"ST",
        "Assembly": "ASMLY-DSK",
        "Booklet Maker": "BKLT",
        "Cutter": "CTR",
        "Flatbed Printer":"FBP",
        "Folder": "FLDR",
        "Jetpress": "JP",
        "Manifest Desk":"MFD",
        "Rimage": "RMG",
        "Magazine":""
}

global label_name
global label_size
global entry_name
global cb2
global name_var
global size_choice

cb2 = None
entry_name = None
label_size = None
label_name = None

def update_options(event=None):
    global label_name
    global label_size
    global entry_name
    global cb2
    global size_choice
    global name_var

    if dropdown_var.get() == "Shelf":
        if label_name == None:
            label_size_opt = ["Large label (5x9cm)", "Small Label (5x4cm)"]
            label_name = tk.Label(root_3, text="Name of Location", fg = 'white', bg = 'black', font=("Verdana", 14), width = 15)
            label_name.pack(padx=5, pady=5, side = tk.LEFT)
            name_var = tk.StringVar()
            entry_name = tk.Entry(root_4,
                            textvariable = name_var,
                            font=("Verdana", 14),
                            width =15
                            )
            entry_name.pack(padx = 5, pady = 5, side = tk.LEFT)

            label_size = tk.Label(root_3, text="Label size", fg = 'white', bg = 'black', font=("Verdana", 14), width = 25)
            label_size.pack(padx=5, pady=5, side = tk.LEFT)
            size_choice = tk.StringVar()
            cb2 = ttk.Combobox(root_4, 
                            values=label_size_opt,
                            font=("Arial", 14),
                            textvariable= size_choice,
                            width = 25
                            )
            cb2.set("Choose a size")
            cb2.option_add("*TCombobox*Listbox.font", ("Verdana", 14))
            cb2.pack(padx = 5, pady = 5, side = tk.LEFT)
            

    else:
        if label_name != None:
            label_name.destroy()
            entry_name.destroy()
            label_size.destroy()
            cb2.destroy()

            label_name = None
            entry_name = None
            label_size = None
            cb2 = None
        

def resize_png(input_path, output_path, width):
    with Image.open(input_path) as img:
        resized_img = img.resize((width, int(width*(5/11))), Image.Resampling.LANCZOS)
        resized_img.save(output_path, format='PNG')

def create_label():
    global size_choice
    global name_var
    start_ID_number = start_ID_variable.get()

    if dropdown_var.get()=="Select a location":
        messagebox.showinfo("ERROR", "Select a location from the dropdown menu")
        return

    else:
        if dropdown_var.get() == "Magazine":
            ID_list = []
            #clear pdf
            label_path = 'labels.pdf'
            writer = PdfWriter()
            if os.path.exists(label_path):
                # Overwrite with an empty PDF
                os.remove(label_path)
            
            #check if csv file is present
            csv_present, csv_file = check_csv()
            if csv_present == True:
                #collect data from first column of .csv in list format
                f = open(csv_file, 'r', newline="", encoding='utf-8')
                for row in f:
                    ID_list.append(row.split()[0])

                #generate labels with these as the ID
                for mag_id in ID_list:
                    img_qr = qrcode.make(mag_id)
                    img_qr.save("QR.png")
                    qr_code = "QR.png"
                    mag_id = mag_id.replace('\ufeff', '')
                    generate_pdf(qr_code, mag_id, writer, label_path)

            with open(label_path, "wb") as f:
                writer.write(f)
            return

        elif dropdown_var.get() == "Shelf":
            #add the options to name for 'location' and small label or large label
            #remove optional text
            name = name_var.get()
            print(name)
            if start_ID_number == '':
                messagebox.showinfo("ERROR", "Choose a start ID number")
                return
            else:
                start_ID_number = int(start_ID_variable.get())
                if end_ID_variable.get() == '':
                    end_ID_number = start_ID_number
                else:
                    end_ID_number = int(end_ID_variable.get())

                #clear pdf
                label_path = 'labels.pdf'
                writer = PdfWriter()
                if os.path.exists(label_path):
                    # Overwrite with an empty PDF
                    os.remove(label_path)

                if size_choice.get() == "Choose a size":
                    messagebox.showinfo("ERROR", "Choose a label size")
                    return
                else:
                    for id_num in range((start_ID_number),(end_ID_number+1),1):
                        #generate new serial number and set to class
                        ID_st = create_st_ID(id_num, name)

                        #GENERATE QR CODES AS IMAGES TO STORE IN DATABASE
                        img_qr = qrcode.make(ID_st)
                        img_qr.save("QR.png")
                        qr_code = "QR.png" 

                        if size_choice.get() == "Small Label (5x4cm)":
                            #generate a png file of label
                            generate_pdf_small(qr_code, ID_st,writer, label_path)

                        elif size_choice.get() == "Large label (5x9cm)":
                            generate_pdf(qr_code, ID_st,writer, label_path)

                        else:
                            print("error")
            with open(label_path, "wb") as f:
                writer.write(f)
            return
        else:
            #collect system type from dropdown + set to class
            system_type_l = dropdown_var.get()


            if start_ID_number == '':
                messagebox.showinfo("ERROR", "Choose a start ID number")
                return
            else:
                start_ID_number = int(start_ID_variable.get())
                if end_ID_variable.get() == '':
                    end_ID_number = start_ID_number
                else:
                    end_ID_number = int(end_ID_variable.get())

                #clear pdf
                label_path = 'labels.pdf'
                writer = PdfWriter()
                if os.path.exists(label_path):
                    # Overwrite with an empty PDF
                    os.remove(label_path)


                for id_num in range((start_ID_number),(end_ID_number+1),1):
                    #generate new serial number and set to class
                    ID_l = create_ID(id_num, system_type_l)

                    #GENERATE QR CODES AS IMAGES TO STORE IN DATABASE
                    img_qr = qrcode.make(ID_l)
                    img_qr.save("QR.png")
                    qr_code = "QR.png" 

                    #generate a png file of label
                    generate_pdf(qr_code, ID_l,writer, label_path)

            with open(label_path, "wb") as f:
                writer.write(f)
            return
        
def generate_pdf(qr_t, ID_t, writer, label_path):
        
    logo = "logo_black.png"
    temp_pdf = 'temp_pdf.pdf'
    height_logo = 10 #mm
    height_qr = 40 #mm
    width_logo = get_scaled_dimensions(logo, height_logo)
    width_qr = get_scaled_dimensions(qr_t, height_qr)
    page_width = 90
    page_height = 50
    pdf = FPDF(unit="mm", format=(page_width, page_height))
    pdf.add_page()

    #add logo on left
    pdf.image(logo, x=1, y=((page_height/2)-(height_logo/2)), w=width_logo, h=height_logo)

    #add qr in the middle
    pdf.image(qr_t, x=11, y=((page_height/2)-(height_qr/2)), w=width_qr, h=height_qr)

    #add ID next to qr
    optional_text = text_option.get()
    pdf.set_xy(50, 15)
    pdf.set_font("Arial", size=15)

    text = ID_t + '\n' + optional_text  # manual line break
    pdf.multi_cell(w=35,h = 6, txt=text, align='J')
    pdf.output(temp_pdf)

    # If original PDF exists, copy its pages
    if os.path.exists(label_path):
        reader = PdfReader(label_path)
        for page in reader.pages:
            writer.add_page(page)

    # Add new page
    new_page_reader = PdfReader(temp_pdf)
    writer.add_page(new_page_reader.pages[0])
    if multiple_var.get() == "":
        return
    else:
        copies = int(multiple_var.get()) -1
        for copy in range(copies):
            writer.add_page(new_page_reader.pages[0])
        return
    

    
def generate_pdf_small(qr_t, ID_t, writer, label_path):
        
    temp_pdf_S = 'temp_pdf_S.pdf'
    height_qr_S = 20 #mm
    width_qr_S = get_scaled_dimensions(qr_t, height_qr_S)
    page_width_S = 40
    page_height_S = 50
    pdf_S = FPDF(unit="mm", format=(page_width_S, page_height_S))
    pdf_S.add_page()

    #add qr in the middle
    pdf_S.image(qr_t, x=10, y=((page_height_S/2)-(height_qr_S/2)), w=width_qr_S, h=height_qr_S)

    #add ID above qr
    pdf_S.set_xy(5,10)
    pdf_S.set_font("Arial", size=10)

    text = ID_t # manual line break
    pdf_S.multi_cell(w=30,h = 6, txt=text, align='C')
    pdf_S.output(temp_pdf_S)

    # If original PDF exists, copy its pages
    if os.path.exists(label_path):
        reader = PdfReader(label_path)
        for page in reader.pages:
            writer.add_page(page)

    # Add new page
    new_page_reader = PdfReader(temp_pdf_S)
    writer.add_page(new_page_reader.pages[0])
    if multiple_var.get() == "":
        return
    else:
        copies = int(multiple_var.get()) -1
        for copy in range(copies):
            writer.add_page(new_page_reader.pages[0])
        return
    



def get_scaled_dimensions(path, target_height_mm, dpi=96):
    img = Image.open(path)
    width_px, height_px = img.size
    width_mm = (width_px / height_px) * target_height_mm
    return width_mm


def send_to_print_new():
    pdf_to_print = 'labels.pdf'

    ##############################################################
    printer_name = 'Brother QL-810W (Copy 1)'
    ###############################################################

    # Set default printer (optional, but good practice)
    win32print.SetDefaultPrinter(printer_name)

    # Create printer device context
    hDC = win32ui.CreateDC()
    hDC.CreatePrinterDC(printer_name)

    # Convert PDF pages to images
    pages = convert_from_path(pdf_to_print, dpi=300)  # Increase dpi for higher resolution

    for i, page in enumerate(pages):
        png_file = f'png_to_print_{i}.png'  # Unique filename per page
        page.save(png_file, 'PNG')
        print(f'Saved {png_file}')

        # Open image for direct printing
        img = Image.open(png_file)

        # Start print job
        hDC.StartDoc(f"Label {i+1}")
        hDC.StartPage()

        dib = ImageWin.Dib(img)

        # Draw image at full size - adjust if needed to fit label dimensions
        dib.draw(hDC.GetHandleOutput(), (0, 0, img.width, img.height))

        # End print job
        hDC.EndPage()
        hDC.EndDoc()

    hDC.DeleteDC()
    return


def create_ID(serial, system):
    #fetch system tag from dictionary to make label
    tag = tag_dict[system]
    new_ID = f"{tag}-{serial}"
    return new_ID

def create_st_ID(serial, name):
    #fetch system tag from dictionary to make label
    tag = 'ST'
    new_ID = f"{tag}-{name}-{serial}"
    return new_ID

def resize_png(input_path, output_path, width):
    with Image.open(input_path) as img:
        resized_img = img.resize((width, int(width*(5/11))), Image.Resampling.LANCZOS)
        resized_img.save(output_path, format='PNG')

def check_csv():
    folder = "."
    file_empty = ""
    count = 0
    for file in os.listdir(folder):
        if file.endswith(".csv"):
            file_csv = file
            count += 1
    if count == 1:
        return True, file_csv
    elif count >= 1:
        messagebox.showinfo("ERROR", "Multiple csv files detected")
        return False, file_empty
    else:
        messagebox.showinfo("ERROR", "Cannot detect .csv file")
        return False, file_empty


# create root window
window = Tk()
# root window title and dimension
window.title("Label Generator")
# Set geometry (widthxheight)
wwidth = window.winfo_screenwidth()
wheight = window.winfo_screenheight()
window.geometry(f"{wwidth}x{wheight}")

root_top = tk.Frame(window, width = wwidth, height=(0.1*wheight), bg = 'black')
root_top.pack(fill=tk.BOTH)

logo_file_gui = 'logo_black.png'
logo_img = tk.PhotoImage(file=logo_file_gui)
logo_img_= tk.Label(root_top,
                image=logo_img)
logo_img_.pack(padx = 5, pady = 5, side = tk.LEFT)

lbltop = tk.Label(root_top,
    font=("Verdana", 24),
    text = "Label Generator",
    bg = 'black',
    fg = 'white'
    )
lbltop.pack(pady = 20, padx =5, anchor = tk.CENTER)

root = tk.Frame(window, width = wwidth, height = (0.5*wheight), bg = 'black')
root.pack(fill = tk.BOTH, padx = 5, pady =5)

label_instructions_2 = tk.Label(root, text="If generating a magazine labels ensure .csv is uploaded and all others are deleted", fg = 'white', bg = 'black', font=("Verdana", 14), width = 100)
label_instructions_2.pack(padx=5, pady=5)


root_2 = tk.Frame(root, width = wwidth, height = (0.2*wheight), bg = 'black')
root_2.pack(fill = tk.X, padx = 5, pady =5, side = tk.TOP)

root_1 = tk.Frame(root, width = wwidth, height = (0.2*wheight), bg = 'black')
root_1.pack(fill = tk.X, padx = 5, pady =5)

root_3 = tk.Frame(root, width = wwidth, height = (0.2*wheight), bg = 'black')
root_3.pack(fill = tk.X, padx = 5, pady =5)

root_4 = tk.Frame(root, width = wwidth, height = (0.2*wheight), bg = 'black')
root_4.pack(fill = tk.X, padx = 5, pady =5, side = tk.BOTTOM)


# Dropdown options  
types = list(tag_dict.keys())
dropdown_var = tk.StringVar()

# Combobox  
label_start = tk.Label(root_2, text="Select location", fg = 'white', bg = 'black', font=("Verdana", 14), width = 15)
label_start.pack(padx=5, pady=5, side = tk.LEFT)
cb = ttk.Combobox(root_1, 
                values=types,
                font=("Arial", 14),
                textvariable=dropdown_var,
                width = 15
                )
cb.set("Select a location")
cb.option_add("*TCombobox*Listbox.font", ("Verdana", 14))
cb.bind("<<ComboboxSelected>>", update_options)
cb.pack(padx = 5, pady = 5, side = tk.LEFT)

##FRAME 2
label_start = tk.Label(root_2, text="Start ID number: ", fg = 'white', bg = 'black', font=("Verdana", 14), width = 15)
label_start.pack(padx=5, pady=5, side = tk.LEFT)
start_ID_variable = tk.StringVar()
entry_1 = tk.Entry(root_1,
                textvariable = start_ID_variable,
                font=("Verdana", 14),
                width = 15
                )
entry_1.pack(padx = 5, pady = 5, side = tk.LEFT)

label_end = tk.Label(root_2, text="End ID number: ", fg = 'white', bg = 'black', font=("Verdana", 14), width = 15)
label_end.pack(padx=5, pady=5, side = tk.LEFT)
end_ID_variable = tk.StringVar()
entry_2 = tk.Entry(root_1,
                textvariable = end_ID_variable,
                font=("Verdana", 14),
                width =15
                )
entry_2.pack(padx = 5, pady = 5, side = tk.LEFT)

label = tk.Label(root_2, text="Optional text entry: ", fg = 'white', bg = 'black', width = 20, font=("Verdana", 14),)
label.pack(padx=5, pady=5, side = tk.LEFT)
text_option = tk.StringVar()
entry_3 = tk.Entry(root_1,
                textvariable = text_option,
                font=("Verdana", 14),
                width =20
                )
entry_3.pack(padx = 5, pady = 5, side = tk.LEFT)

label_4 = tk.Label(root_2, text="No. of each label (default = 1): ", fg = 'white', bg = 'black', font=("Verdana", 14), width = 30)
label_4.pack(padx=5, pady=5, side = tk.LEFT)
multiple_var = tk.StringVar()
entry_4 = tk.Entry(root_1,
                textvariable = multiple_var,
                font=("Verdana", 14),
                width = 30
                )
entry_4.pack(padx = 5, pady = 5, side = tk.LEFT)

root_bottom = tk.Frame(window, width = wwidth, height = (0.4*wheight), bg = 'black')
root_bottom.pack(fill = tk.Y, padx = 5, pady =5)

btn1 = tk.Button(root_bottom, 
            text = "Generate Labels",
            font=("Verdana", 14),
            bg="#2E2E2E",        # dark grey background
            fg="white",          # white text
            activebackground="#444444",  # hover effect
            activeforeground="white",
            relief="flat",       # removes raised border
            #padx=10,
            #pady=6,
            width=20,            # consistent button width
            command=create_label)
btn1.pack(padx = 5, pady = 5, side = tk.LEFT)


print_btn = tk.Button(root_bottom, 
            text = "Print",
            font=("Verdana", 14),
            bg="#2E2E2E",        # dark grey background
            fg="white",          # white text
            activebackground="#444444",  # hover effect
            activeforeground="white",
            relief="flat",       # removes raised border
            #padx=10,
            #pady=6,
            width=20,            # consistent button width
            command=send_to_print_new)
print_btn.pack(padx = 5, pady = 5, side = tk.RIGHT)

# Execute Tkinter
window.mainloop()