from tkinter import Tk
from tkinter import filedialog


def send_way_to_file():
    root = Tk()
    ftypes = [('Excel файл', '*.xlsx')]
    dlg = filedialog.Open(root, filetypes=ftypes)
    fl = dlg.show()
    root.destroy()
    return fl