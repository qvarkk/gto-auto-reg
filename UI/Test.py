import tkinter as tk
from tkinter import filedialog
from tkinter import messagebox
import threading
import random
from mailtm import Email
from openpyxl import load_workbook
from openpyxl.styles import Font
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys



def generateNewAccount(row, excelPath, drivePath):
  email = Email()
  email.register()
  address = str(email.address)

  driver = webdriver.Chrome(drivePath[0:-4])
  driver.implicitly_wait(10)

  book = load_workbook(excelPath)
  sheet = book.active

  firstname = sheet[f'A{row}'].value
  surname = sheet[f'B{row}'].value
  dogname = sheet[f'C{row}'].value
  bdate = sheet[f'D{row}'].value
  day = bdate.day
  month = bdate.month
  year = bdate.year
  if day < 10:  day = "0" + str(day)
  if month < 10: month = "0" + str(month)


  driver.get('https://user.gto.ru/user/register') 

  email_input_field = driver.find_element(By.CSS_SELECTOR, '#gto_user_register_form_email')
  driver.execute_script(f"arguments[0].value = '{address}';", email_input_field)
  email_input_field.send_keys(' ')

  send_code_button = driver.find_element(By.CSS_SELECTOR, 'button[style="opacity: 1;"]')
  driver.execute_script("arguments[0].click()", send_code_button)

  def input_person_data():
    number = random.randint(100000000, 999999999)
    driver.execute_script('''
      let a; (a = (string) => {
        document.querySelector("#gto_user_register_form_surname").value = string.split(" ")[0]
        document.querySelector("#gto_user_register_form_firstname").value = string.split(" ")[1]
        document.querySelector("#gto_user_register_form_patronymic").value = string.split(" ")[2]
    
        if (string.split(" ")[2][string.split(" ")[2].length - 1] === "ч")
          document.querySelector("#gto_user_register_form_gender_0").checked = true
        else
          document.querySelector("#gto_user_register_form_gender_1").checked = true

        document.querySelector("#gto_user_register_form_address_actual_full_address").value = "Краснодарский край, Лабинский р-н, ст-ца Владимирская "
        document.querySelector("#gto_user_register_form_phone").value = "+79%s"
        document.querySelector("#gto_user_register_form_password").value = "qwerty123"
        document.querySelector("#gto_user_register_form_agree_personal_data").click()
      })("%s %s %s")
    ''' % (str(number), firstname, surname, dogname))
    year_input = driver.find_element(By.CSS_SELECTOR, "#gto_user_register_form_birthday")
    driver.execute_script("arguments[0].select", year_input)
    for i in range(20):
      year_input.send_keys(Keys.DELETE)
    year_input.send_keys(f"{day}.{month}.{year}")
    address_input = driver.find_element(By.CSS_SELECTOR, "#gto_user_register_form_address_actual_full_address")
    driver.execute_script("arguments[0].select()", address_input)
    driver.execute_script("arguments[0].blur()", address_input)

    

  def listener(message):
    subject = message['subject']
    text = message['text']

    if subject == 'Ваш код активации для сайта gto.ru':
      index = text.find('gto.ru: ')
      confirmation_code = text[index + 7 : index + 14]

      code_input = driver.find_element(By.CSS_SELECTOR, '#gto_user_register_form_activateCode')
      confirm_button = driver.find_element(By.CSS_SELECTOR, 'button[name="next"]')

      driver.execute_script(f"arguments[0].value = '{confirmation_code}'", code_input)
      driver.execute_script("arguments[0].click()", confirm_button)
      driver.find_element(By.CSS_SELECTOR, "#gto_user_register_form_birthday")
      input_person_data()
    elif subject == 'ВФСК ГТО: Ваш аккаунт активирован':
      index = text.find('Ваш УИН: ')
      uin = text[index + 8 : index + 22]
      sheet[f'E{row}'].value = uin
      sheet[f'E{row}'].font = Font(name='Arial')
      book.save(excelPath)
      driver.close()
      return

  email.start(listener, interval = 1)




class App(threading.Thread):
  def __init__(self):
    threading.Thread.__init__(self)

  def browseFilesForExcel(self):
    filename = filedialog.askopenfilename(initialdir = "./", title = "Выбрать файл", filetypes = (("Excel файлы", "*.xlsx*"),("all files", "*.*")))
    self.sel_ex_file_text.configure(text = filename)

  def browseFilesForWebdrive(self):
    filename = filedialog.askopenfilename(initialdir = "./", title = "Выбрать файл", filetypes = (("Приложения", "*.exe*"),("all files", "*.*")))
    self.sel_wd_file_text.configure(text = filename)


  def startAutoreg(self, rowStart, rowEnd, pauseTime, excelPath, drivePath):
    if not(rowStart > 0 and rowEnd > 0 and pauseTime > 0 and rowStart < rowEnd):
      messagebox.showerror('Ошибка в вводе переменных', 'Где-то ты напортачил со строками')
      return

    self.start_button['state'] = 'disabled'
    self.explore_button['state'] = 'disabled'
    self.explore_wd_button['state'] = 'disabled'
    self.pause_entry['state'] = 'disabled'
    self.end_rows_entry['state'] = 'disabled'
    self.start_rows_entry['state'] = 'disabled'

    for i in range(rowStart, rowEnd + 1):
      self.root.after(pauseTime * 1000, generateNewAccount(i, excelPath, drivePath))

    self.start_button['state'] = 'active'
    self.explore_button['state'] = 'active'
    self.explore_wd_button['state'] = 'active'
    self.pause_entry['state'] = 'normal'
    self.end_rows_entry['state'] = 'normal'
    self.start_rows_entry['state'] = 'normal'

  def initializeAutoreg(self):
    self.rows_start = int(float(self.start_rows_entry.get()))
    self.rows_end = int(float(self.end_rows_entry.get()))
    self.pause_time = int(float(self.pause_entry.get()))
    self.excel_path = self.sel_ex_file_text.cget("text")
    self.drive_path = self.sel_wd_file_text.cget("text")
    self.startAutoreg(self.rows_start, self.rows_end, self.pause_time, self.excel_path, self.drive_path)


  def quit(self):
    if messagebox.askokcancel("Выйти", "Вы действительно хотите выйти?"):
        self.root.destroy()

  def run(self):
    self.root = tk.Tk()
    self.root.protocol('WM_DELETE_WINDOW', self.quit)

    self.root.geometry('700x350')
    self.root.title('Авторег ГТО')

    self.label = tk.Label(self.root, text="Hello World!", font=('Arial', 18))
    self.label.pack()

    def limitEntry(*args):
        self.valueSR = self.sr_value.get()
        if len(self.valueSR) > 4: self.sr_value.set(self.valueSR[:4])
        self.valueER = self.er_value.get()
        if len(self.valueER) > 4: self.er_value.set(self.valueER[:4])
        self.valueSEC = self.sec_value.get()
        if len(self.valueSEC) > 2: self.sec_value.set(self.valueSEC[:2])

    self.sr_value = tk.StringVar()
    self.sr_value.trace('w', limitEntry)
    self.er_value = tk.StringVar()
    self.er_value.trace('w', limitEntry)
    self.sec_value = tk.StringVar()
    self.sec_value.trace('w', limitEntry)

    self.main_frame = tk.Frame(self.root)
    self.main_frame.columnconfigure(0, weight = 1)
    self.main_frame.columnconfigure(1, weight = 2)
    self.main_frame.rowconfigure(0, weight = 1)
    self.main_frame.rowconfigure(1, weight = 1)
    self.main_frame.rowconfigure(2, weight = 1)
    self.main_frame.rowconfigure(3, weight = 1)
    self.main_frame.rowconfigure(4, weight = 1)
    self.main_frame.rowconfigure(5, weight = 1)

    self.start_rows_text = tk.Label(self.main_frame, text = 'С какого ряда', font = ('Arial', 14))
    self.start_rows_text.grid(row = 0, column = 0, sticky = tk.W, padx = 10)

    self.start_rows_entry = tk.Entry(self.main_frame, text = '1', width = 4, font = ('Arial', 12), textvariable = self.sr_value)
    self.start_rows_entry.grid(row = 0, column = 1, sticky = tk.E, padx = 10)

    self.start_rows_text = tk.Label(self.main_frame, text = 'До какого ряда', font = ('Arial', 14))
    self.start_rows_text.grid(row = 1, column = 0, sticky = tk.W, padx = 10)

    self.end_rows_entry = tk.Entry(self.main_frame, text = '1', width = 4, font = ('Arial', 12), textvariable = self.er_value)
    self.end_rows_entry.grid(row = 1, column = 1, sticky = tk.E, padx = 10)

    self.pause_text = tk.Label(self.main_frame, text = 'Пауза между итерациями', font = ('Arial', 14))
    self.pause_text.grid(row = 2, column = 0, sticky = tk.W, padx = 10)

    self.pause_entry = tk.Entry(self.main_frame, text = '1', width = 4, font = ('Arial', 12), textvariable = self.sec_value)
    self.pause_entry.grid(row = 2, column = 1, sticky = tk.E, padx = 10)

    self.files_frame = tk.Frame(self.main_frame)
    self.files_frame.columnconfigure(0, weight = 1)
    self.files_frame.columnconfigure(1, weight = 3)
    self.files_frame.columnconfigure(2, weight = 1)
    self.files_frame.rowconfigure(0, weight = 1)
    self.files_frame.rowconfigure(1, weight = 1)

    self.excel_file_text = tk.Label(self.files_frame, text = 'Путь к Excel', font = ('Arial', 12))
    self.excel_file_text.grid(row = 0, column = 0, sticky = tk.W, padx = 10)

    self.sel_ex_file_text = tk.Label(self.files_frame, text = 'Не выбрано', font = ('Arial', 12))
    self.sel_ex_file_text.grid(row = 0, column = 1, sticky = tk.W + tk.E)

    self.explore_button = tk.Button(self.files_frame, text = 'Выбрать путь', font = ('Arial', 12), command = self.browseFilesForExcel)
    self.explore_button.grid(row = 0, column = 2, sticky = tk.E, padx = 10)

    self.wb_file_text = tk.Label(self.files_frame, text = 'Путь к Webdrive', font = ('Arial', 12))
    self.wb_file_text.grid(row = 1, column = 0, sticky = tk.W, padx = 10, pady = 5)

    self.sel_wd_file_text = tk.Label(self.files_frame, text = 'Не выбрано', font = ('Arial', 12))
    self.sel_wd_file_text.grid(row = 1, column = 1, sticky = tk.W + tk.E, pady = 5)

    self.explore_wd_button = tk.Button(self.files_frame, text = 'Выбрать путь', font = ('Arial', 12), command = self.browseFilesForWebdrive)
    self.explore_wd_button.grid(row = 1, column = 2, sticky = tk.E, padx = 10, pady = 5)

    self.files_frame.grid(columnspan = 2, rowspan = 2, sticky  = tk.W + tk.E, pady = 20)
    self.main_frame.pack(fill='x')

    self.start_button = tk.Button(self.root, text='Начать', font=('Arial', 18), command = self.initializeAutoreg)
    self.start_button.pack(pady = 15)

    self.root.mainloop()

app = App()
app.run()