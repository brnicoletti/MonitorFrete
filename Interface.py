from tkinter import filedialog

class Janela():
    def __init__(self, master):
        self.master = master


def openFile():
    file_path = filedialog.askopenfilename(
        initialdir="C:\\Users\\rafar\\Downloads",
        title="Selecione o arquivo",
        filetypes= (("planilha", ("*.xls", "*.xlsx", "*.csv")), ("all files", "*.*"))
    )
    return file_path