
from tkinter import ttk
import tkinter as tk
import tkinter.filedialog as fd
import pandas as pd
import threading
import hashlib
import os
import pandas.io.formats.excel
import logging
from logging.handlers import RotatingFileHandler
import pem
pandas.io.formats.excel.header_style = None


class App(tk.Tk):

    def __init__(self):
        super().__init__()
        self.resizable(False, False)
        self.geometry("500x250")
        self.title("Simple Pseudonymiser")
        self.welcomeLabel = tk.Label(self, text="Welcome to the Simple Pseudonymiser")
        self.welcomeLabel.pack(padx=60, pady=10)

        self.logger = logging.getLogger()
        handler = RotatingFileHandler("pseudo_log.log", maxBytes=10*1024*1024, backupCount=5)
        formatter = logging.Formatter('%(asctime)s %(name)-12s %(levelname)-8s %(message)s')
        handler.setFormatter(formatter)
        self.logger.addHandler(handler)
        self.logger.setLevel(logging.DEBUG)

        style = ttk.Style()
        style.configure("foreGreen.Label", foreground="green")
        style.configure("foreRed.Label", foreground="red")
        style.configure("foreOrange.Label", foreground="coral4")
        style.configure("button.flat", relief="flat")

        self._fileName = tk.StringVar()
        self._result = tk.StringVar()
        self._salt = tk.StringVar()
        self._saltOutput = tk.StringVar()
        self._pseudoOutput =tk.StringVar()

        self._inputFileName = tk.StringVar()
        self._resultOutput = tk.StringVar()

        self._pseudoOutput.set("Pseudonymise the file")
        self.btn_salt = ttk.Button(self, text="Choose a cert/pem file to generate your salt",
                                   command=self.choose_pem_file, width=100)

        self.btn_salt.pack(padx=60, pady=10)

        self.btn_file = ttk.Button(self, text="Choose excel file with 'identifier' column",
                                   command=self.choose_file, state="disabled", width = 100)
        self.btn_file.pack(padx=60, pady=10)

        self.btn_pseudo = ttk.Button(self, textvariable=self._pseudoOutput,
                                     command=self.pseudonymize_file, state="disabled", width=100)
        self.btn_pseudo.pack(padx=60, pady=10)

        self.resultLabel = ttk.Label(self, textvariable=self._resultOutput,
                                     width = 400, wraplength=390, font=('Helvetica', 9, 'bold'))
        self.resultLabel.configure(style="foreGreen.Label",anchor="center")
        self.resultLabel.pack(padx=60, pady=10)

        self.processing_bar = ttk.Progressbar(self, orient='horizontal', mode='determinate', length=300)

    def report_callback_exception(self, exc, val, tb):
        self.logger.error('Error!', val)
        self.destroy_unmapped_children(self)


    def destroy_unmapped_children(self, parent):
        """
        Destroys unmapped windows (empty gray ones which got an error during initialization)
        recursively from bottom (root window) to top (last opened window).
        """
        children = parent.children.copy()
        for index, child in children.items():
            if not child.winfo_ismapped():
                parent.children.pop(index).destroy()
            else:
                self.destroy_unmapped_children(child)

    def choose_salt_file(self):
        self.btn_file['state'] = 'disabled'
        self._salt.set("")
        file_types = (("Text File", "*.txt"),)
        filepath = fd.askopenfilename(title="Open PEM file", filetypes=file_types)
        exists = os.path.isfile(filepath)
        if exists:
            self._salt.set(filepath)
            with open(self._salt.get()) as f:
                self._salt.set(f.readline())
            self._saltOutput.set("Your salt term is " + self._salt.get()[4:].rjust(len(self._salt.get()), "*"))
            self.btn_file['state'] = 'normal'
            self.logger.info('Salt Loaded')

    def choose_pem_file(self):
        self.btn_file['state'] = 'disabled'
        self._salt.set("")
        file_types = (("pem file", "*.pem"),("cert file", "*.cert"),("crt file", "*.crt"))
        filepath = fd.askopenfilename(title="Open pem or cert file", filetypes=file_types)
        exists = os.path.isfile(filepath)
        if exists:
            certs = pem.parse_file(filepath)
            self._salt.set(filepath)
            self._salt.set(certs[0].sha1_hexdigest)
            self._saltOutput.set("Your salt term is " + self._salt.get()[4:].rjust(len(self._salt.get()), "*"))
            self.btn_file['state'] = 'normal'
            self.logger.info('Salt Loaded')

    def choose_file(self):
        self.btn_pseudo['state'] = 'disabled'
        self._fileName.set("")
        file_types = (("xlsx", "*.xlsx"),)
        filepath = fd.askopenfilename(title="Open file", filetypes=file_types)
        exists = os.path.isfile(filepath)
        if exists:
            self._fileName.set(filepath)
            self._inputFileName.set(os.path.basename(self._fileName.get()))
            self.btn_pseudo['state'] = 'normal'
            self._resultOutput.set("")
            self.logger.info('Data File Loaded '+self._fileName.get())

            temp_name = self.get_file_display_name(self._fileName.get())

            self._pseudoOutput.set("Pseudonymise the file "+temp_name)

    def pseudo(self, x):
        sentence = str(x) + self._salt.get()
        return str(hashlib.blake2s(sentence.encode('utf-8')).hexdigest())

    def pseudonymize_file(self):
        self.logger.info('Starting Pseudo: ' + self._fileName.get())
        self.processing_bar.pack(padx=60, pady=10)
        self.processing_bar.start(1000)
        t = threading.Thread(target=self.pseudonymize_file_callback)
        t.start()

    def kill_progress(self):
        self.processing_bar.stop()
        self.processing_bar.pack_forget()

    def get_extension(self, filename):
        filename, file_extension = os.path.splitext(filename)
        return file_extension if file_extension else None

    def get_file_display_name(self, filename):
        temp_name = os.path.basename(filename);
        return temp_name[:15] + ('..' + self.get_extension(temp_name) if len(temp_name) > 15 else '')

    def pseudonymize_file_callback(self):
        try:
            self.btn_pseudo['state'] = 'disabled'
            self.btn_file['state'] = 'disabled'
            self.btn_salt['state'] = 'disabled'
            temp_name = self.get_file_display_name(self._fileName.get())

            self.resultLabel.config(style="foreOrange.Label")
            self._resultOutput.set(temp_name + " is being loaded")
            self.update()

            df = pd.read_excel(self._fileName.get(), dtype='str', encoding='utf-8')
            df.columns = df.columns.str.lower()
            if 'identifier' not in df.columns:
                self.resultLabel.config(style="foreRed.Label")
                self._resultOutput.set("No 'identifier' column exists in file!")
                self.btn_pseudo['state'] = 'normal'
                self.btn_file['state'] = 'normal'
                self.btn_salt['state'] = 'normal'
                self.kill_progress()

            else:
                temp_name = str(self._fileName.get())
                temp_name = temp_name.replace(".xlsx", "_psuedo.xlsx")
                new_name = temp_name
                self.btn_pseudo['state'] = 'disabled'
                self.resultLabel.config(style="foreOrange.Label")
                temp_name = self.get_file_display_name(self._fileName.get())
                self._resultOutput.set(temp_name + " is being pseudonymised")
                self.config(cursor="wait")
                self.update()

                df['DIGEST'] = df.identifier.apply(self.pseudo)
                del df['identifier']
                self._result.set(os.path.basename(temp_name))
                if os.path.exists(new_name):
                    os.remove(new_name)
                df.to_excel(new_name, index=False)
                self._resultOutput.set(os.path.basename(str(self._fileName.get())) + " has been pseudonymised")
                self.resultLabel.config(style="foreGreen.Label")
                self.btn_pseudo['state'] = 'disabled'
                self.btn_file['state'] = 'normal'
                self.btn_salt['state'] = 'normal'
                self.config(cursor="")
                self.logger.info('Completing Pseudo: ' + self._fileName.get())
                self.kill_progress()

        except BaseException as error:
            self.resultLabel.config(style="foreRed.Label")
            self._resultOutput.set('An exception occurred: details in log file')
            self.btn_pseudo['state'] = 'normal'
            self.btn_file['state'] = 'normal'
            self.btn_salt['state'] = 'normal'
            self.logger.error('An exception occurred: {}'.format(error))
            self.kill_progress()


if __name__ == "__main__":
    app = App()
    app.mainloop()
