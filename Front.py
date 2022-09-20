import tkinter as tk
import os
from ReportsDownloader import ReportsExporter


class UI:
    def __init__(self):
        self.root = tk.Tk()
        self.settings()

    def settings(self):
        ws = self.root.winfo_screenwidth()
        hs = self.root.winfo_screenheight()
        x = ws - 750
        y = hs - 750
        self.root.geometry('%dx%d+%d+%d' % (400, 400, x, y))
        self.root.iconbitmap('DATA/SMSbot.ico')
        self.root.title(' Reports Downlowder V3.1')
        self.generate_widgets()

    def OptionMenu_SelectionEvent(self, event):
        if event == "Reporte Total Bancos":
            self.stores["state"] = "disabled"
            self.entryyy2.delete(0,tk.END)
            self.entryyy2.insert(0,'detalle_banco')
        else:
            self.stores["state"] = "normal"
            self.entryyy2.delete(0,tk.END)
            self.entryyy2.insert(0,'reporte_ingresos')
            
        
    def generate_widgets(self):
        self.canvas = tk.Canvas(self.root)
        self.canvas.place(relwidth=1, relheight=1)

        self.labelframe = tk.LabelFrame(self.canvas, text="Input Box", bd=2)
        self.labelframe.place(relx=0.05, rely=0.05, relwidth=0.9, relheight=0.8)
        
        reportbl = tk.Label(self.labelframe, text='Select Report: ')
        reportbl.place(relx=0.05, rely=0.02, relwidth=0.22, relheight=0.1)
        
        options = ["Reporte Pos vs Ingresos", "Reporte Total Bancos"]
        self.clicked = tk.StringVar()
        self.clicked.set("Reporte Pos vs Ingresos")
        rprtspinr = tk.OptionMenu(self.labelframe, self.clicked, *options, command=self.OptionMenu_SelectionEvent)
        rprtspinr.place(relx=0.28, rely=0.02, relwidth=0.5, relheight=0.1)
        
        lblb = tk.Label(self.labelframe, text='Select Store')
        lblb.place(relx=0.05, rely=0.15, relwidth=0.18, relheight=0.05)

        self.stores = tk.Listbox(self.labelframe, selectmode="multiple")
        self.stores.place(relx=0.05, rely=0.21, relwidth=0.3, relheight=0.35)
        stores_code = ['SAMS', 'Supercenter', 'bodega Aurerra', 'Superama', 'Mi bodega', 'Bodega Express']
        for store in stores_code:
            self.stores.insert(tk.END, store)

        lb22 = tk.Label(self.labelframe, text='date from :')
        lb22.place(relx=0.4, rely=0.2, relwidth=0.2, relheight=0.1)

        self.entryy = tk.Entry(self.labelframe)
        self.entryy.place(relx=0.6, rely=0.2, relwidth=0.3, relheight=0.1)
        self.entryy.insert(tk.INSERT, '01-01-2021')

        lb222 = tk.Label(self.labelframe, text='date to :')
        lb222.place(relx=0.4, rely=0.32, relwidth=0.2, relheight=0.1)

        self.entryyy = tk.Entry(self.labelframe)
        self.entryyy.place(relx=0.6, rely=0.32, relwidth=0.3, relheight=0.1)
        self.entryyy.insert(tk.INSERT, '01-01-2021')

        lb222 = tk.Label(self.labelframe, text='default fn :')
        lb222.place(relx=0.4, rely=0.44, relwidth=0.2, relheight=0.1)

        self.entryyy2 = tk.Entry(self.labelframe)
        self.entryyy2.place(relx=0.6, rely=0.44, relwidth=0.3, relheight=0.1)
        self.entryyy2.insert(tk.INSERT, 'reporte_ingresos')

        framed = tk.LabelFrame(self.labelframe, text=" Delay (s) ", bd=2)
        framed.place(relx=0.05, rely=0.6, relwidth=0.9, relheight=0.18)

        lb2 = tk.Label(framed, text='Timeout:')
        lb2.place(relx=0.05, rely=0.05, relwidth=0.15, relheight=0.9)

        self.entry = tk.Entry(framed)
        self.entry.place(relx=0.2, rely=0.05, relwidth=0.1, relheight=0.9)
        self.entry.insert(tk.INSERT, '100')

        lb5 = tk.Label(framed, text='Delay:')
        lb5.place(relx=0.4, rely=0.05, relwidth=0.15, relheight=0.9)

        self.entry5 = tk.Entry(framed)
        self.entry5.place(relx=0.55, rely=0.05, relwidth=0.1, relheight=0.9)
        self.entry5.insert(tk.INSERT, '5')
        
        lb6 = tk.Label(framed, text='waitFile:')
        lb6.place(relx=0.7, rely=0.05, relwidth=0.15, relheight=0.9)

        self.entry6 = tk.Entry(framed)
        self.entry6.place(relx=0.85, rely=0.05, relwidth=0.1, relheight=0.9)
        self.entry6.insert(tk.INSERT, '1000')

        #####################################

        framedexcl = tk.LabelFrame( self.labelframe, text=" Excel Formating ", bd=2)
        framedexcl.place(relx=0.05, rely=0.8, relwidth=0.9, relheight=0.18)

        self.chbcolvar = tk.BooleanVar()
        fchcol = tk.Checkbutton(framedexcl, text="Fecha Col", variable=self.chbcolvar, onvalue=True, offvalue=False)
        fchcol.place(relx=0.05, rely=0.05, relwidth=0.22, relheight=0.7)

        self.obj = ReportsExporter(svpth=os.path.join(os.getcwd(), 'Reports'), isHide=False, timeout=self.entry, time_delay=self.entry5, wait_file=self.entry6)

        button222 = tk.Button(framedexcl, bg='#80b3ff', text="Select Folder", command=lambda: self.obj.uplaod_file())
        button222.place(relx=0.4, rely=0.05, relwidth=0.25, relheight=0.7)

        button333 = tk.Button(framedexcl, bg='#80b3ff', text="Formating", command=lambda: self.obj.formating(self.chbcolvar.get(), self.clicked))
        button333.place(relx=0.7, rely=0.05, relwidth=0.25, relheight=0.7)

        #################################

        button2 = tk.Button(self.canvas, bg='#80b3ff', text="Open Browser", command=lambda: self.obj.process(self.stores, self.entryy, self.entryyy, self.entryyy2, self.clicked))
        button2.place(relx=0.5, rely=0.88, relwidth=0.2, relheight=0.08)

        button3 = tk.Button(self.canvas, bg='#80b3ff', text="Start", command=lambda: self.obj.strt())
        button3.place(relx=0.75, rely=0.88, relwidth=0.2, relheight=0.08)

        # button4 = tk.Button(self.canvas, bg='#80b3ff', text="Rename Files", command=lambda: self.obj.rename_files())
        # button4.place(relx=0.75, rely=0.88, relwidth=0.2, relheight=0.08)
        
        self.root.mainloop()


if __name__ == '__main__':
    UI()
    
