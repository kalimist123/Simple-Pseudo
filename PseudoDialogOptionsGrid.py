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
import gc
import sys

pandas.io.formats.excel.header_style = None


class App(tk.Tk):

    def __init__(self):
        super().__init__()
        self.resizable(False, False)
        self.geometry("500x320")
        self.title("Simple Pseudonymiser")
        self.welcomeLabel = tk.Label(self, text="Welcome to the Simple Pseudonymiser")
        self.welcomeLabel.pack(padx=60, pady=10)

        self.logger = logging.getLogger()
        handler = RotatingFileHandler("pseudo_log.log", maxBytes=10 * 1024 * 1024, backupCount=5)
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
        self._pseudoOutput = tk.StringVar()

        self._inputFileName = tk.StringVar()
        self._resultOutput = tk.StringVar()

        self._pseudoOutput.set("Pseudonymise the file")
        self.btn_salt = ttk.Button(self, text="Choose a cert/pem file to generate your salt",
                                   command=self.choose_pem_file, width=100)

        self.btn_salt.pack(padx=60, pady=10)

        self.btn_file = ttk.Button(self, text="Choose excel file and the select column to pseudo",
                                   command=self.choose_file, state="disabled", width=100)
        self.btn_file.pack(padx=60, pady=10)

        self.menu_label_text = tk.StringVar()
        self.menu_label_text.set("Choose the excel column that you would like to have pseudonymised")
        self.menu_label = tk.Label(self, textvariable=self.menu_label_text)
        self.options = ['']
        self.om_variable = tk.StringVar(self)
        self.om = ttk.OptionMenu(self, self.om_variable, *self.options)
        self.om.configure(width=60)
        self.alwaysActiveStyle(self.om)

        self.om['state'] = 'disabled'
        self.om_variable.trace("w", self.option_menu_selection_event)

        self.btn_pseudo = ttk.Button(self, textvariable=self._pseudoOutput,
                                     command=self.pseudonymize_file, state="disabled", width=100)

        self.resultLabel = ttk.Label(self, textvariable=self._resultOutput,
                                     width=400, wraplength=390, font=('Helvetica', 9, 'bold'))
        self.resultLabel.configure(style="foreGreen.Label", anchor="center")
        self.processing_bar = ttk.Progressbar(self, orient='horizontal', mode='determinate', length=400)

    def report_callback_exception(self, exc, val, tb):
        exc_type, exc_value, exc_traceback = sys.exc_info()
        self.logger.error('exception line: ' + str(exc_traceback.tb_lineno) + ' error: ' + str(exc_value))
        self.destroy_unmapped_children(self)
        self.btn_pseudo.pack(padx=60, pady=10)

    def alwaysActiveStyle(self, widget):
        widget.config(state="active")
        widget.bind("<Leave>", lambda e: "break")

    def show_pickers(self):
        self.menu_label.pack(padx=60, pady=0)
        self.om.pack(padx=60, pady=10)
        self.alwaysActiveStyle(self.om)
        self.btn_pseudo.pack(padx=60, pady=10)

    def hide_pickers(self):
        self.menu_label.pack_forget()
        self.om.pack_forget()
        self.btn_pseudo.pack_forget()

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
        if self.resultLabel.winfo_ismapped():
            self.resultLabel.pack_forget()
        self.btn_file['state'] = 'disabled'
        self._salt.set("")
        file_types = (("crt file", "*.crt"), ("cert file", "*.cert"), ("pem file", "*.pem"))
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
        if self.resultLabel.winfo_ismapped():
            self.resultLabel.pack_forget()
        self.btn_pseudo['state'] = 'disabled'
        self._fileName.set("")
        file_types = (("xlsx", "*.xlsx"),)
        filepath = fd.askopenfilename(title="Open file", filetypes=file_types)
        exists = os.path.isfile(filepath)
        self.hide_pickers()
        if exists:
            self.btn_salt['state'] = 'disabled'
            self.btn_file['state'] = 'disabled'

            self._fileName.set(filepath)
            self._inputFileName.set(os.path.basename(self._fileName.get()))
            self.btn_pseudo['state'] = 'normal'
            self._resultOutput.set("")
            self.logger.info('Data File Loaded ' + self._fileName.get())
            first_row = pd.read_excel(self._fileName.get(), dtype='str', encoding='utf-8', nrows=1)
            first_row.columns = first_row.columns.str.strip().str.lower().str.replace(' ', '_').str.replace('(',
                                                                                                            '').str.replace(
                ')', '')

            self.options = list(first_row)
            self.update_option_menu()
            self.om['state'] = 'normal'
            self.om_variable.set(self.options[0])
            self._pseudoOutput.set("Pseudonymise the column " + self.om_variable.get())
            self.show_pickers()
            self.btn_salt['state'] = 'normal'
            self.btn_file['state'] = 'normal'
            del first_row
            gc.collect()

    def update_option_menu(self):
        menu = self.om["menu"]
        menu.delete(0, "end")
        for string in self.options:
            menu.add_command(label=string,
                             command=lambda value=string: self.om_variable.set(value))

    def option_menu_selection_event(self, *args):
        self._pseudoOutput.set("Pseudonymise the column " + self.om_variable.get())
        pass

    def pseudo(self, x):
        sentence = str(x) + self._salt.get()
        return str(hashlib.blake2s(sentence.encode('utf-8')).hexdigest())

    def pseudonymize_file(self):
        self.logger.info('Starting Pseudo: ' + self._fileName.get())
        self.processing_bar.pack(padx=60, pady=10)
        if not self.resultLabel.winfo_ismapped():
            self.resultLabel.pack(padx=60, pady=10)
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
            df.columns = df.columns.str.strip().str.lower()\
                .str.replace(' ', '_').str.replace('(', '').str.replace(')', '')

            temp_name = str(self._fileName.get())
            temp_name = temp_name.replace(".xlsx", "_psuedo.xlsx")
            new_name = temp_name
            self.btn_pseudo['state'] = 'disabled'
            self.resultLabel.config(style="foreOrange.Label")
            temp_name = self.get_file_display_name(self._fileName.get())
            self._resultOutput.set(temp_name + " is being pseudonymised")
            self.config(cursor="wait")
            self.update()

            df['DIGEST'] = df[self.om_variable.get()].apply(self.pseudo)
            del df[self.om_variable.get()]

            self._result.set(os.path.basename(temp_name))
            if os.path.exists(new_name):
                os.remove(new_name)
            df.to_excel(new_name, index=False)
            del df
            gc.collect()
            self._resultOutput.set(str(self._fileName.get()) + " has been pseudonymised")
            self.resultLabel.config(style="foreGreen.Label")
            self.btn_pseudo['state'] = 'disabled'
            self.btn_file['state'] = 'normal'
            self.btn_salt['state'] = 'normal'
            self.config(cursor="")
            self.logger.info('Completing Pseudo: ' + self._fileName.get())
            self.kill_progress()
            self.hide_pickers()

        except BaseException as error:
            exc_type, exc_value, exc_traceback = sys.exc_info()
            self.resultLabel.config(style="foreRed.Label")
            self._resultOutput.set('An exception occurred: details in log file')
            self.btn_pseudo['state'] = 'normal'
            self.btn_file['state'] = 'normal'
            self.btn_salt['state'] = 'normal'
            self.logger.error('An exception occurred: {}'.format(error))
            self.logger.error('exception line: ' + str(exc_traceback.tb_lineno) + ' error: ' + str(exc_value))
            self.kill_progress()
            self.hide_pickers()


if __name__ == "__main__":
    app = App()
    app.mainloop()
