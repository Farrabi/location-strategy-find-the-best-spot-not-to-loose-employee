import pandas as pd
import re
from time import sleep
import random
import openpyxl
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys
from selenium import webdriver
from selenium import webdriver
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.common.action_chains import ActionChains
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.common.by import By
import tkinter
from tkinter import filedialog
from tkinter import messagebox
from PIL import Image, ImageTk


input_name=['Office address']
Entries=[]
root_2=tkinter.Tk()
root_2.geometry("720x406")
root_2.title('gmappy')
image = Image.open("broly.jpg")
background_image=ImageTk.PhotoImage(image)
background_label = tkinter.Label(root_2, image=background_image)
background_label.pack()
y_=70
i=70

for input_ in input_name:
    label =  tkinter.Label(root_2, text = input_).place(x = 20,y = y_)
    y_=y_+20
    entry=tkinter.Entry(root_2)
    entry.place(x=600,y=i)
    i=i+20
    Entries.append(entry)
    
def get1():
    global address
    address=[]
    for entry in Entries:
        address.append(entry.get())
    
label1 =tkinter.Label(root_2, text = 'Input office address')
label1.place(x =180,y =20)
label1.config(font=('Calibri', 22))
sbmitbtn = tkinter.Button(root_2, text = "Submit",activebackground = "pink", activeforeground = "blue", command=lambda: [get1(), root_2.destroy()])
sbmitbtn.pack()

sbmitbtn.place(x = 30, y = 220)


root_2.mainloop()

root = tkinter.Tk()
root.withdraw()

file = filedialog.askopenfile(parent=root,mode='rb',title='Choose a file')
if file != None:
    data = file.read()
    file.close()
    print("I got %d bytes from this file." % len(data))

    

df=pd.read_excel(file.name, sheet_name='employees_info', encoding = 'utf-8')
employee_first_name=df.loc[:,'Employee first name']
employee_first_name_clean=employee_first_name.fillna('no name')
employee_last_name= df.loc[:,'Employee last name']
employee_last_name_clean=employee_last_name.fillna('no name')
employee_id=df.loc[:,'Employee number']
employee_id_clean=employee_id.fillna('no id')
employee_address=df.loc[:,'Address ']
employee_address_clean = employee_address.fillna('no address')
old_commutation_time= df.loc[:,'Commutation time (min)']
old_commutation_time_clean=old_commutation_time.fillna('no time')


address_comm_time=dict(zip(employee_address_clean, old_commutation_time_clean))


def scrape(dep_address, dest_address):
    details_text=[]
    try:
        chrome_options = Options()
        chrome_options.add_argument('--no-sandbox')
        chrome_options.add_argument('start-maximized')
        chrome_options.add_argument('disable-infobars')
        chrome_options.add_argument("--disable-extensions")
        chrome_options.add_argument("--disable-gpu")
        chrome_options.add_argument("--headless")
        chrome_options.add_argument("user-agent=Mozilla/5.0 (Windows NT 6.1) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/76.0.3809.100 Safari/537.36")
        driver = webdriver.Chrome(options=chrome_options)
        driver.get('https://www.google.com/maps/dir///@35.7203438,139.7831109,14z/data=!4m2!4m1!3e3')
        sleep(3)
        departure= driver.find_element_by_xpath("/html/body/jsl/div[3]/div[8]/div[3]/div[1]/div[2]/div/div[3]/div[1]/div[1]/div[2]/div/div/input")
        departure.send_keys(dep_address)
        destination = driver.find_element_by_xpath("/html/body/jsl/div[3]/div[8]/div[3]/div[1]/div[2]/div/div[3]/div[1]/div[2]/div[2]/div/div/input")
        destination.send_keys(dest_address)
        button=driver.find_element_by_xpath('/html/body/jsl/div[3]/div[8]/div[3]/div[1]/div[2]/div/div[3]/div[1]/div[2]/div[2]/button[1]')
        ActionChains(driver).move_to_element(button).click(button).perform()
        element = WebDriverWait(driver, 4)
        element.until(EC.element_to_be_clickable((By.XPATH, '/html/body/jsl/div[3]/div[8]/div[9]/div/div[1]/div/div/div[5]/div[1]/div[2]'))).click()
        element2= WebDriverWait(driver, 10)
        time=element2.until(EC.element_to_be_clickable((By.CSS_SELECTOR, '#pane > div > div.widget-pane-content.scrollable-y > div > div > div.section-trip-summary.noprint > div.section-trip-summary-header > h1 > span:nth-child(2) > span.section-trip-summary-subtitle > span'))).text
        sleep(10)
        elements=driver.find_elements_by_class_name('transit-stop-name')
        stops=[]
        for i in elements:
            stops.append(i.text)
        driver.quit()
    except Exception as e:
        print(e)
    return (time, stops)                  


new_commute={k:scrape(k,address[0])for k , _ in address_comm_time.items()}

df2=pd.DataFrame(data=list(new_commute.values()), columns=['commutation time', 'transit'])

D=pd.concat([employee_first_name_clean, employee_last_name_clean, employee_address_clean, df2], axis=1)

D.to_excel('new_commutation_time.xlsx', encoding='utf_8_sig', index=False)







