import os
import time
import pandas as pd
import datetime
from bs4 import BeautifulSoup
from threading import Thread
from selenium import webdriver
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.remote.command import Command
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.support.wait import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.common.by import By
from tkinter import messagebox 
import easygui

HOST  = os.getenv('WHOST')

class ReportsExporter:
    def __init__(self, svpth, isHide, timeout, time_delay, wait_file):
        self.time_delay = time_delay
        self.wait_file = wait_file
        self.oldurl = True
        self.timeout = timeout
        self.isHide = isHide
        self.path = None
        self.svpth = svpth
        self.browser = None
        self.start = False
        self.wait = None
        self.default_fn = ''
        self.stores_code = {
            'SAMS': 0,
            'Supercenter': 3,
            'bodega Aurerra': 5,
            'Superama': 6,
            'Mi bodega': 17,
            'Bodega Express': 37
        }

    def initialize(self):
        self.start = False
        chrome_options = Options()
        chrome_options.add_experimental_option('prefs', {
            "download.default_directory": self.svpth,
            "download.prompt_for_download": False,
            "download.directory_upgrade": True
        })
        self.browser = webdriver.Chrome(
            executable_path = 'chromedriver.exe', options=chrome_options)
        self.wait = WebDriverWait(self.browser, int(self.timeout.get().strip()))
        
    def is_browser_open(self, driver):
        try:
            driver.execute(Command.STATUS)
            return True
        except:
            return False
    
    def rename_file(self, old, new):
        try:
            os.rename(old, new)
        except Exception as e:
            messagebox.showinfo('info', str(e))
        
    def file_is_exist(self, timeout, file_path):
        tm = 0
        while tm < timeout:
            if os.path.isfile(file_path):
                time.sleep(3)
                return True
            time.sleep(1)
            tm += 1
        return False
        
    def download_excel(self, stores, date_from, date_to):
        try:
            now = datetime.datetime.now()
            ts = now.strftime("%m%d%Y_%H%M%S")
            init_path = os.path.join(self.svpth, f'Reports_{ts}') 
            os.mkdir(init_path)
            d1 = datetime.datetime.strptime(date_from, "%d-%m-%Y")
            d2 = datetime.datetime.strptime(date_to, "%d-%m-%Y")
            rng = abs((d2 - d1).days)
            dfltnm = os.path.join(self.svpth, self.default_fn.get())
            tmdly = int(self.time_delay.get().strip())
            oldurl = True if self.oldurl.get() == "Reporte Pos vs Ingresos" else False            
            while not self.start:
                time.sleep(3)
            for d in range(rng + 1):
                try:
                    date = d1 + datetime.timedelta(days=d)
                    current_date = date.date().strftime("%d-%m-%Y").replace('-', '%2F')
                    cdate = current_date.replace('%2F', '')
                    if not oldurl:
                        stores = ['']
                    for store in stores:
                        try:
                            if oldurl:
                                url = f'{HOST}/Dolares/reporte_ingresos.asp?negocio={self.stores_code[store]}&determinante=&banco=&FechaIni={current_date}&FechaFin={current_date}&transaccion=0&reporte=4'
                                self.browser.get(url)
                                excel_btn = self.wait.until(EC.element_to_be_clickable((By.XPATH, '//*[@id="ImagenExcel"]')))   
                            else:
                                url = f'{HOST}/Dolares/reporte_bancos.asp?negocio=&determinante=&banco=1&FechaIni={current_date}&FechaFin={current_date}&transaccion=0&reporte=3'
                                self.browser.get(url)
                                excel_btn = self.wait.until(EC.element_to_be_clickable((By.XPATH, '//*[@id="Linea1"]/b/a/b')))  
                            excel_btn.click()
                            file_exist = self.file_is_exist(int(self.wait_file.get().strip()), f'{dfltnm}.xls')
                            if file_exist:
                                if oldurl:
                                    self.rename_file(f'{dfltnm}.xls', os.path.join(init_path, f'{store}_{cdate}.xls'))
                                else:
                                    cdt = cdate[0:2]
                                    self.rename_file(f'{dfltnm}.xls', os.path.join(init_path, f'{cdt}.xls'))
                        except Exception as e:
                            messagebox.showinfo('info', str(e))
                        finally:
                            time.sleep(tmdly)
                except Exception as e:
                    messagebox.showinfo('info', str(e))
        except Exception as e:
            messagebox.showinfo('info', str(e))
        finally:
            time.sleep(8)
            messagebox.showinfo('Done', 'Dwonloading Files secussfully done.')
    
    def strt(self):
        self.start = True
        Thread(target=self.start_process).start()
            
    def start_process(self):
        if not self.is_browser_open(self.browser):
            messagebox.showwarning('warning', 'Please click "open browser" button first and logged in.')
        else:
            values = [self.values.get(idx) for idx in self.values.curselection()]
            self.download_excel(values, self.from_date.get().strip(), self.to_date.get().strip())
        
    
    def process(self, stores, from_date, to_date, default_fn, oldurl):
        Thread(target=self.open_browser, args=(stores, from_date, to_date, default_fn, oldurl)).start()
            
    def open_browser(self, stores, from_date, to_date, default_fn, oldurl):
        try:
            self.initialize()
            self.oldurl, self.default_fn, self.values = oldurl, default_fn, stores
            self.from_date, self.to_date = from_date, to_date
            self.browser.get(HOST)
        except Exception as e:
            messagebox.showinfo('info', str(e))
        

    def uplaod_file(self):
        path = easygui.diropenbox(title='Select Your Excel File')
        if path is not None:
            self.path = path

    def formating(self, with_slashes, tpe):
        tpe = tpe.get()
        if tpe == "Reporte Pos vs Ingresos":
            Thread(target=self.formating_excel, args=(with_slashes,)).start()
        else:
            Thread(target=self.formating_excel_2).start()
        
        
    def formating_excel_2(self):
        files = os.listdir(self.path)
        data = []
        for i, f in enumerate(files):
            if f.endswith('.xls'):
                soup = BeautifulSoup(open(os.path.join(self.path, f)), 'html.parser')
                if i == 0:
                    HTML_data = soup.find_all("table")[:-1]
                else:
                    HTML_data = soup.find_all("table")[1:-1]
                for element in HTML_data:
                    element_td = element.find_all("td")
                    row = [elm.get_text().strip() for elm in element_td]
                    data.append(row)
        dataFrame = pd.DataFrame(data=data)
        dataFrame.to_excel('CONSOLIDADO Total bancos.xlsx', header=False, index=False)
        messagebox.showinfo('info', 'CONSOLIDADO Excel Seccessfully Done')


    def htmh2excel(self, fl):
        data = []
        if fl.endswith('.xls'):
            soup = BeautifulSoup(open(fl), 'html.parser')
            HTML_data = soup.find_all("table")[4:-1]
            for element in HTML_data:
                element_td = element.find_all("td")
                row = [elm.get_text().strip() for elm in element_td]
                if row[-1] != '0':
                    data.append(row)
        return data


    def formating_excel(self, with_slashes):
        files = os.listdir(self.path)
        formato = {'AURRERA': "BODEGA AURRERA", "MIBODEGA": "MI BODEGA"}
        result = []
        for f in files:
            try:
                filename = f.split('.')[0]
                comp, date = filename.split('_')
                data = self.htmh2excel(os.path.join(self.path, f))
                for dataframe in data:
                    try:
                        comp = formato[comp]
                    except:
                        pass
                    if with_slashes:
                        slshdt = date[0:2] + '/' + date[2:4] + '/' + date[4:]
                        result.append(dataframe + ['', f"Reproceso {date}", date, comp, slshdt])
                    else:
                        result.append(dataframe + ['', f"Reproceso {date}", date, comp])
            except:
                pass
        if with_slashes:
            df = pd.DataFrame(result, columns=['Tienda', 'Dolares', 'Tienda', 'Dolares', 'Diferencia', 'F6', 'FileName', 'TERMINO', 'Formato', 'Fecha']) 
        else:
            df = pd.DataFrame(result, columns=['Tienda', 'Dolares', 'Tienda', 'Dolares', 'Diferencia', 'F6', 'FileName', 'TERMINO', 'Formato'])
        try:
            df.to_excel('CONSOLIDADO TOTAL FECHA.xlsx', index=False)
            messagebox.showinfo('info', 'Formating Excel Seccessfully Done')
        except:
            messagebox.showinfo('info', 'Permission denied: "CONSOLIDADO TOTAL FECHA.xlsx"\nThe File is open with other program\nPlease close it and try again.')
            

