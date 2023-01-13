import eel
from tkinter import filedialog, messagebox

@eel.expose
def askForLocFile():
  filepath = filedialog.askopenfilename(initialdir = './', title = 'Выбрать файл', filetypes = (('Текстовые файлы', ('*.txt')),('all files', '*.*')))
  with open(filepath, encoding = 'utf-8', mode = 'r') as file:
    text = file.read()
    return text
  
@eel.expose
def askForExcelFile():
  filepath = filedialog.askopenfilename(initialdir = "./", title = "Выбрать файл", filetypes = (("Excel файлы", ("*.xlsx*", '*.xlsm', '*.xltx', '*.xltm')),("all files", "*.*")))
  if filepath.endswith(".xlsx") or filepath.endswith(".xlsm") or filepath.endswith(".xltx") or filepath.endswith(".xltm"):
    return filepath
  else: 
    return 'error'

@eel.expose
def askForDriverFile():
  filepath = filedialog.askopenfilename(initialdir = "./", title = "Выбрать файл", filetypes = (("Приложения", "*.exe*"),("all files", "*.*")))
  if filepath.endswith(".exe"):
    return filepath
  else: 
    return 'error'

if __name__ == '__main__':
  eel.init('./ui')
  eel.start('./main.html', size = (700, 500))