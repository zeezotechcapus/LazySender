from fileinput import filename
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.common.keys import Keys
import pyperclip
import time
import tkinter
from tkinter import Tk, ttk
from tkinter import filedialog
import openpyxl
import os
import sys 
from tkinter import scrolledtext

top = Tk()
top.geometry("600x400")
top.title("LazySender")


#btn =(top, text = 'Click me !', bd = '5', command = top.destroy)
#inputtxt = tkinter.Text(top, height = 20, width = 20)
  
num_var=tkinter.StringVar()
masseg_var=tkinter.StringVar()
 
# creating a label for name and Numbers 
num_label = tkinter.Label(top, text = 'Numbers', font = ('calibre',10,'bold'))
num_label.place(x = 100, y = 80)
#num_entry = tkinter.Entry(top,textvariable = num_var, font=('calibre',10,'normal'), justify= 'left')
#num_entry.place( x=180,y=80 , width=300 , height=100)
  
# creating a label for password
massage_label = tkinter.Label(top, text = 'Message', font = ('calibre',10,'bold'))
massage_label.place(x=100 ,y=120)
programe_name = tkinter.Label(top, text = 'LazySender', font = ('calibre',28,'bold'))
programe_name.place(x=200 ,y=10)

#massage_entry=tkinter.Entry(top, textvariable = masseg_var, font = ('calibre',10,'normal'),justify= 'center')
#massage_entry.place( x=180,y=120 , width=300 , height=180)
text_area = scrolledtext.ScrolledText(top, x =180, y= 120,width=65, height = 12,font=("Times New Roman", 10))
  
text_area.grid(column=0, row=2, pady=120, padx=180)
def retrieve_input():
    global inputValue
    inputValue=text_area.get("1.0","end-1c")
    print(inputValue)
    return inputValue


# placing cursor in text area
text_area.focus()

#num_label.grid(row=1000,column=1)

def upload_excel():
    global filename
    filename = filedialog.askopenfilename(title="Select a File", filetype=(("Excel", "*.xlsx"), ("Excel", "*.xls")))
    # uploading the excel file 
    wb_obj = openpyxl.load_workbook(filename)
    ws = wb_obj.active
    global nums 
    nums = []
    for col in ws['A']:
        nums.append(col.value)
    
up_bt = ttk.Button(top, text="Upload", command= upload_excel).place(x=180,y=80)



def sender():
    #num_value  = num_var.get()
    masseg_var= retrieve_input()

    #path genaration func
    def resource_path(relative_path):
        try:
            base_path = sys._MEIPASS
        except Exception:
            base_path = os.path.dirname(__file__)
        return os.path.join(base_path, relative_path)
    
    browser = webdriver.Chrome(
    resource_path('./driver/chromedriver.exe') )

    browser.maximize_window()

    browser.get('https://web.whatsapp.com/')
    groups = []
    groups = nums
    #print(groups)
    mass = masseg_var
    #print(mass)
    for group in groups:
            search_xpath = '//div[@contenteditable="true"][@data-tab="3"]'

            search_box = WebDriverWait(browser, 500).until(
                EC.presence_of_element_located((By.XPATH, search_xpath))
            )

            search_box.clear()

            time.sleep(1)

            pyperclip.copy(group)

            search_box.send_keys(Keys.SHIFT, Keys.INSERT)  # Keys.CONTROL + "v"

            time.sleep(1)
            group_xpath = f'//span[@title="{group}"]'
            group_title = browser.find_element_by_xpath(group_xpath)

            group_title.click()

            time.sleep(2)


            #input_xpath = '//div[@contenteditable="true"][@data-tab="1"]'
            input_xpath = '//*[@id="main"]/footer/div[1]/div/span[2]/div/div[2]/div[1]'
            input_box = browser.find_element_by_xpath(input_xpath)

            #pyperclip.copy(msg)
            #input_box.send_keys(Keys.SHIFT, Keys.INSERT)  # Keys.CONTROL + "v"

            input_box.send_keys(mass)
            #btt = '//*[@id="main"]/footer/div[1]/div/span[2]/div/div[2]/div[2]/button'
            #btt_box =  browser.find_element_by_xpath(btt)
            time.sleep(1)

            input_box.click()
            input_box.send_keys(Keys.ENTER)


            time.sleep(5)
btn = ttk.Button(top, text="send", command= sender).place(x=280,y=320)
top.mainloop()
