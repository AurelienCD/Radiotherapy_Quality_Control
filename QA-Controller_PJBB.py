from tkinter import *
import tkinter as tk
from tkinter import ttk
from PIL import ImageTk ,Image

class MyApp:

    def __init__(self):

        self.window = Tk()
        self.window.title("QA-Controller_PJBB")
        self.window.geometry("1800x1000")
        self.window.minsize(480, 360)
        #self.window.iconbitmap('/home/aureliencd/Documents/Baclesse_ACD/QA-Controller_PJBB/icon.ico')
        self.window.config(background='#2086dc')

        label_window = Label(self.window, text="Bienvenue sur l'application QA-Controller_RT", font=("Calibri", 20),fg='gray')
        #label_window.grid(row=0, column=1)
        label_window.pack()

        logo=ImageTk.PhotoImage(Image.open("/home/aureliencd/Documents/Baclesse_ACD/QA-Controller_PJBB/logo.jpeg"))

        label_logo=Label(image=logo)
        label_logo.pack()

        self.notebook = ttk.Notebook(self.window)
        self.notebook.pack(pady=10, expand=True)

        # create frames
        self.frame1 = ttk.Frame(self.notebook, width=1500, height=800)
        self.frame2 = ttk.Frame(self.notebook, width=1500, height=800)

        self.frame1.pack(fill='both', expand=True)
        self.frame2.pack(fill='both', expand=True)

        # add frames to notebook

        self.notebook.add(self.frame1, text='2021')
        self.notebook.add(self.frame2, text='2020')


        self.notebook2 = ttk.Notebook(self.frame1)
        self.notebook2.grid(row=0, column=0)

        # create frames
        self.frame3 = ttk.Frame(self.notebook2, width=1500, height=800)
        self.frame4 = ttk.Frame(self.notebook2, width=1500, height=800)

        self.frame3.pack(fill='both', expand=True)
        self.frame4.pack(fill='both', expand=True)

        # add frames to notebook

        self.notebook2.add(self.frame3, text='Rapid_Arc_1')
        self.notebook2.add(self.frame4, text='Rapid_Arc_2')



        label_CQ_Quotidien = Label(self.frame3, text="CQ_Quotidien", font=("Courrier", 20),fg='black')
        label_CQ_Quotidien.grid(pady=10, sticky = W, row=0, column=0)
        label_CQ_Quotidien2 = Label(self.frame3, text="                          ", font=("Courrier", 20),fg='black')
        label_CQ_Quotidien2.grid(pady=10, sticky = W, row=0, column=1)

        label_CQ_Hebdo = Label(self.frame3, text="CQ_Hebdo", font=("Courrier", 20),fg='black')
        label_CQ_Hebdo.grid(pady=10, sticky = W, row=0, column=2)

        label_CQ_Mensuel = Label(self.frame3, text="CQ_Mensuel", font=("Courrier", 20),fg='black')
        label_CQ_Mensuel.grid(pady=10, sticky = W, row=1, column=0)
        label_CQ_Mensuel2 = Label(self.frame3, text="                          ", font=("Courrier", 20),fg='black')
        label_CQ_Mensuel2.grid(pady=10, sticky = W, row=1, column=1)


        label_CQ_Semestriel = Label(self.frame3, text="CQ_Semestriel", font=("Courrier", 20),fg='black')
        label_CQ_Semestriel.grid(pady=10, sticky = W, row=1, column=2)

        label_CQ_Annuel = Label(self.frame3, text="CQ_Annuel", font=("Courrier", 20),fg='black')
        label_CQ_Annuel.grid(pady=10, sticky = W, row=9, column=0)

# afficher
app = MyApp()
app.window.mainloop()
