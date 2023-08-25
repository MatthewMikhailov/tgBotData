from tkinter import Tk
from tkinter import filedialog


def send_way_to_file():
    root = Tk()
    ftypes = [('Excel файл', '*.xls')]
    dlg = filedialog.Open(root, filetypes=ftypes)
    fl = dlg.show()
    root.destroy()
    return fl


