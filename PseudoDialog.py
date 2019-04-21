
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
pandas.io.formats.excel.header_style = None


class App(tk.Tk):

    def __init__(self):
        super().__init__()
        self.resizable(False, False)
        self.geometry("500x350")
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
        style.configure("foreOrange.Label", foreground="orange")

        self._fileName = tk.StringVar()
        self._result = tk.StringVar()
        self._salt = tk.StringVar()
        self._saltOutput = tk.StringVar()

        self._inputFileName = tk.StringVar()
        self._resultOutput = tk.StringVar()

        self.btn_salt = ttk.Button(self, text="Choose a file with your salt string", command=self.choose_salt_file, width=100)

        self.btn_salt.pack(padx=60, pady=10)
        self.saltLabel = tk.Label(self, textvariable=self._saltOutput, justify="left",  width=100)

        self.saltLabel.pack(padx=60, pady=10)
        self.btn_file = ttk.Button(self, text="Choose excel file with an identifier column to pseudo",
                                   command=self.choose_file, state="disabled", width = 100)
        self.btn_file.pack(padx=60, pady=10)
        self.fileLabel = tk.Label(self, textvariable=self._inputFileName,justify="left", width=100)

        self.fileLabel.pack(padx=60, pady=10)
        self.btn_pseudo = ttk.Button(self, text="Pseudonymise the file",
                                     command=self.pseudonymize_file, state="disabled", width=100)
        self.btn_pseudo.pack(padx=60, pady=10)

        self.resultLabel = ttk.Label(self, textvariable=self._resultOutput, justify="center", width = 400)
        self.resultLabel.configure(style="foreGreen.Label")
        self.resultLabel.pack(padx=60, pady=10)

    def report_callback_exception(self, exc, val, tb):
        self.destroy_unmapped_children(self)
        self.logger.error('Error!', val)

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
        file_types = (("Text File", "*.txt"),)
        self._salt.set(fd.askopenfilename(title="Open salt file", filetypes=file_types))

        with open(self._salt.get()) as f:
            self._salt.set(f.readline())
        self._saltOutput.set("Your salt term is " + self._salt.get()[4:].rjust(len(self._salt.get()), "*"))
        self.btn_file['state'] = 'normal'
        self.logger.info('Salt Loaded')

    def choose_file(self):
        file_types = (("xlsx", "*.xlsx"),("All files", "*"))
        self._fileName.set(fd.askopenfilename(title="Open file", filetypes=file_types))
        self._inputFileName.set(os.path.basename(self._fileName.get()))
        self.btn_pseudo['state'] = 'normal'
        self._resultOutput.set("")
        self.logger.info('Data File Loaded '+self._fileName.get())

    def pseudo(self, x):
        sentence = str(x) + self._salt.get()
        return str(hashlib.blake2s(sentence.encode('utf-8')).hexdigest())

    def pseudonymize_file(self):
        self.logger.info('Starting Pseudo: ' + self._fileName.get())

        t = threading.Thread(target=self.callback)
        t.start()

    def callback(self):
        try:
            self.btn_pseudo['state'] = 'disabled'
            self.btn_file['state'] = 'disabled'
            self.btn_salt['state'] = 'disabled'
            temp_name = str(self._fileName.get())
            self.resultLabel.config(style="foreOrange.Label")
            self._resultOutput.set(os.path.basename(temp_name) + " is being loaded")
            self.update()

            df = pd.read_excel(self._fileName.get(), dtype='category')
            if 'identifier' not in df.columns:
                self.resultLabel.config(style="foreRed.Label")
                self._resultOutput.set("No identifier column exists in file that you have selected!")
                self.btn_pseudo['state'] = 'normal'
                self.btn_file['state'] = 'normal'
                self.btn_salt['state'] = 'normal'

            else:
                temp_name = str(self._fileName.get())
                temp_name = temp_name.replace(".xlsx", "_psuedo.xlsx")
                self.btn_pseudo['state'] = 'disabled'
                self.resultLabel.config(style="foreOrange.Label")
                self._resultOutput.set("patientid column in " + os.path.basename(temp_name) + " is being pseudonymised")
                self.config(cursor="wait")
                self.update()

                df['DIGEST'] = df.identifier.apply(self.pseudo)
                del df['identifier']
                self._result.set(os.path.basename(temp_name))
                if os.path.exists(temp_name):
                    os.remove(temp_name)
                df.to_excel(temp_name, index=False)
                self._resultOutput.set("identifier column in " + os.path.basename(temp_name) + " has been pseudonymised")
                self.resultLabel.config(style="foreGreen.Label")
                self.btn_pseudo['state'] = 'disabled'
                self.btn_file['state'] = 'normal'
                self.btn_salt['state'] = 'normal'
                self.config(cursor="")
                self.logger.info('Completing Pseudo: ' + self._fileName.get())
        except BaseException as error:
            self.resultLabel.config(style="foreRed.Label")
            self._resultOutput.set('An exception occurred: {}'.format(error))
            self.btn_pseudo['state'] = 'normal'
            self.btn_file['state'] = 'normal'
            self.btn_salt['state'] = 'normal'
            self.logger.error('An exception occurred: {}'.format(error))


if __name__ == "__main__":
    app = App()
    app.mainloop()
