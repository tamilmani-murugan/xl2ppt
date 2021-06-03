from os import path
import tkinter as tk
from app import xl_to_ppt, xtract_ppt
from tkinter import filedialog
from PIL import Image, ImageTk
from tkinter import Tk

root = tk.Tk()
root.title('  XL to PPT')
root.iconbitmap('images/app_logo.ico')
root.geometry('1600x1200')

canvas = tk.Canvas(root, width=600, height=650)
canvas.pack()
canvas.place(relx=0.25, rely=0)


class inputs:
    def open_file(self):
        global sample


def open_sample_layout():
    global sample_layout
    sample_layout = filedialog.askopenfilename(title="Select Powerpoint File",
                                               filetypes=(('Powerpoint File', '*.ppt*'), ("all files", "*.*")))
    if sample_layout != '':
        done_img = tk.Label(image=done)
        done_img.image = done
        canvas.create_window(400, 190, window=done_img)


def open_sample_data():
    global sample_data
    sample_data = filedialog.askopenfilename(title="Select Excel File",
                                             filetypes=(('Excel File', '*.xls*'), ("all files", "*.*")))
    if sample_data != '':
        done_img = tk.Label(image=done)
        done_img.image = done
        canvas.create_window(400, 265, window=done_img)


def open_inf_data():
    global inf_data
    inf_data = filedialog.askopenfilename(title="Select Inf File",
                                          filetypes=(('.Inf File', '*.inf'), ("all files", "*.*")))
    if inf_data != '':
        done_img = tk.Label(image=done)
        done_img.image = done
        canvas.create_window(400, 340, window=done_img)


def get_output_name():
    btn_text.set('Processing..')
    global output_file
    opt_dir = path.dirname(path.abspath(sample_layout))
    output_file = (opt_dir+"\\"+txt_box.get()+path.splitext(sample_layout)[1])
    if sample_layout is not None and sample_data is not None and inf_data is not None:
        try:
            xl_to_ppt(sample_layout, sample_data, inf_data, output_file)
        except:
            btn_text.set('Failed')
        else:
            btn_text.set('Completed')


def generate_inf():
    btn_text1.set('Processing..')
    global output_file
    opt_dir = path.dirname(path.abspath(sample_layout))
    if txt_box.get() == '':
        output_file = (opt_dir+"\\"+'xl2ppt.inf')
    else:
        output_file = (opt_dir+"\\"+txt_box.get()+'.inf')

    if sample_layout is not None:
        try:
            xtract_ppt(sample_layout, output_file)
        except:
            btn_text1.set('Failed')
        else:
            btn_text1.set('Completed')


xlppt_img = Image.open('images/xltoppt.png')
xlppt_img = ImageTk.PhotoImage(xlppt_img)
xlppt_img_lbl = tk.Label(image=xlppt_img)
xlppt_img_lbl.image = xlppt_img

done = Image.open('images/done.png')
done = ImageTk.PhotoImage(done)
done_img = tk.Label(image=done)
done_img.image = done

browse = Image.open('images/browse.jpg')
browse = ImageTk.PhotoImage(browse)

btn_text = tk.StringVar()
btn_text.set('Create PPT')

btn_text1 = tk.StringVar()
btn_text1.set('Run INF')

label1 = tk.Label(canvas, text='select a PPT file',
                  font='Raleway', bg='lightgrey')
button1 = tk.Button(image=browse, command=lambda: open_sample_layout())

label2 = tk.Label(canvas, text='select a XL file',
                  font='Raleway', bg='lightgrey')
button2 = tk.Button(image=browse, command=lambda: open_sample_data())

label3 = tk.Label(canvas, text='select a INF file',
                  font='Raleway', bg='lightgrey')
button3 = tk.Button(image=browse, command=lambda: open_inf_data())

label4 = tk.Label(canvas, text='Enter output name', font='Raleway')
txt_box = tk.Entry(canvas)

button4 = tk.Button(canvas, text='Quit', font='Raleway',
                    command=canvas.quit, bg='red', fg='white', height=2, width=10)
button5 = tk.Button(canvas,  textvariable=btn_text, font='Raleway',
                    command=get_output_name, bg='black', fg='white', height=2, width=15)
button6 = tk.Button(canvas,  textvariable=btn_text1, font='Raleway',
                    command=generate_inf, bg='black', fg='white', height=2, width=15)


canvas.create_window(300, 50, window=xlppt_img_lbl)
canvas.create_window(300, 175, window=label1)
canvas.create_window(300, 200, window=button1)
canvas.create_window(300, 250, window=label2)
canvas.create_window(300, 275, window=button2)
canvas.create_window(300, 325, window=label3)
canvas.create_window(300, 350, window=button3)
canvas.create_window(300, 400, window=label4)
canvas.create_window(300, 425, window=txt_box)
canvas.create_window(200, 500, window=button6)
canvas.create_window(400, 500, window=button5)
canvas.create_window(300, 600, window=button4)


root.mainloop()
