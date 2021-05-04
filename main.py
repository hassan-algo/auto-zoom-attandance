import os
import pytesseract as tess
import tkinter as tk
from tkinter import filedialog, StringVar, IntVar
import pandas as pd
from PIL import Image, ImageTk, ImageOps

desktop = os.path.join(os.path.join(os.path.expanduser('~')), 'Desktop')
tess.pytesseract.tesseract_cmd = r'C:\Program Files\Tesseract-OCR\tesseract.exe'
root = tk.Tk()

canvas = tk.Canvas(root, width=600, height=200)
canvas.grid(column=3, row=6)
Names = ""
data = ""
Attandees = []


def checkAttandance():
    present = []
    for index, row in Names.iterrows():
        yes = False
        name = row['NAME']
        name = name.lower()
        for i in Attandees:
            if i in name or name in i:
                present.append("P")
                yes = True
                break
        if not yes:
            present.append("A")

    return present


def convertImagesToText(files):
    files = root.tk.splitlist(files)
    text = []
    for i in files:
        img = Image.open(i)
        gray_image = ImageOps.grayscale(img)
        # gray_image.show()
        text.append(tess.image_to_string(gray_image).splitlines())
    return text


def SelectExcelCallBack():
    global Names, data
    file = filedialog.askopenfilename(initialdir="", title="Select Excel File...",
                                      filetypes=[
                                          ("xlsx files", "*.xlsx"), ("xls files", "*.xls"), ("All files", "*.*")])
    data = pd.read_excel(file, engine='openpyxl')
    Names = pd.DataFrame(data, columns=['NAME'])
    print(Names)


def photosCallBack():
    global files, Attandees, data
    files = filedialog.askopenfilenames(initialdir="", title="Select Images...",
                                        filetypes=[
                                            ("jpeg files", "*.jpeg"), ("png files", "*.png"), ("jpg files", "*.jpg"),
                                            ("All files", "*.*")])

    convertImages = convertImagesToText(files)

    for i in convertImages:
        for j in i:
            if j:
                Attandees.append(''.join(k.replace('-', '').lower() for k in j if not k.isdigit()))
    print(Attandees)
    present = checkAttandance()
    data['Status'] = present

    data.to_excel(desktop+'/Attendees.xlsx')
    print(data)


B = tk.Button(root, text="Select Photos", command=photosCallBack, width=20, height=7, bg='white')
label = tk.Label(root, textvariable="Select Pics")
B1 = tk.Button(root, text="Select Excel Sheet", command=SelectExcelCallBack, width=20, height=7, bg='blue', fg='white')
exit = tk.Button(root, text="QUIT", command=root.destroy, width=10, height=2, bg='red', fg='white')
B1.grid(column=3, row=2)
B.grid(column=3, row=6)
exit.grid(column=5, row=6)

root.mainloop()