import undetected_chromedriver as uc
#from lib2to3.pgen2 import driver
from selenium.webdriver.common.by import By
import tkinter
import random
import operator
import time
import os
from selenium.webdriver.common.keys import Keys
import tkinter as tk
from tkinter.constants import CENTER
from selenium.webdriver.support.ui import Select
from selenium.webdriver.common.action_chains import ActionChains
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.chrome.options import Options
from pynput.keyboard import Controller


#初始變數建立--------------------------
Comment_List= []

#爬蟲---------------------------------


def scrape(driver, account, password, keyword):

    Keyword= keyword.split(',')

    time.sleep(3)
    try:
        Email= driver.find_element(By.ID,"email")
        Password= driver.find_element(By.ID,"pass")
        Email.send_keys(account)
        Password.send_keys(password)
        Login= driver.find_element(By.ID,"loginbutton")
        driver.execute_script("arguments[0].click();", Login)
    except:
        print("already login")

    time.sleep(5)

    # 抓取"相關留言"物件，並點選
    ActionChains(driver).move_by_offset(600,500).click().perform()
    WebDriverWait(driver,30,0.5).until(EC.presence_of_element_located((By.CSS_SELECTOR,"div.x1jx94hy.x12nagc")))
    form= driver.find_elements(By.CSS_SELECTOR,"div.x1jx94hy.x12nagc")

    span= form[-1].find_elements(By.CSS_SELECTOR,"span.x193iq5w")
    #form_click= driver.find_element(By.CSS_SELECTOR,".x78zum5> span")
    driver.execute_script("arguments[0].click();", span[0])

    #點選"所有留言"
    WebDriverWait(driver,30,0.5).until(EC.presence_of_element_located((By.CSS_SELECTOR,"div.xb57i2i.x1q594ok.x5lxg6s.x6ikm8r.x1ja2u2z > div")))
    List= driver.find_elements(By.CSS_SELECTOR,"div.xb57i2i.x1q594ok.x5lxg6s.x6ikm8r.x1ja2u2z > div")
    for i in range(0,len(List)):
       if "留言" in List[i].text:
           find_all= List[i].find_elements(By.CSS_SELECTOR,"span")
        
    driver.execute_script("arguments[0].click();",find_all[-1])


    #抓取第一則貼文
    WebDriverWait(driver,30,0.5).until(EC.presence_of_element_located((By.CLASS_NAME,"x78zum5.x1n2onr6.xh8yej3")))
    new_comment= driver.find_elements(By.CLASS_NAME,"x78zum5.x1n2onr6.xh8yej3")

    for i in range(0,len(new_comment)):
        if "留言" in new_comment[i].text:
            Web= new_comment[i]
            break


    #展開留言內容
    flag= 0
    start= 0
    WebDriverWait(driver,30,0.5).until(EC.presence_of_element_located((By.CLASS_NAME,"x1i10hfl.xjbqb8w")))
    while flag!= 1:
        flag= 1
        time.sleep(5)
        more= Web.find_elements(By.CLASS_NAME,"x1i10hfl.xjbqb8w")
        for i in range(start,len(more)):
            print(more[i].text)
            if "檢視另" in more[i].text or (operator.not_("隱藏" in more[i].text) and "則回覆" in more[i].text)or "顯示更多" in more[i].text or "查看更多" in more[i].text:
                driver.execute_script("arguments[0].click();",more[i])
                time.sleep(5)
                flag= 0
                print("yes")
        start= len(more)-1
        #more.clear


    #抓取留言內容
    WebDriverWait(driver,30,0.5).until(EC.presence_of_element_located((By.CSS_SELECTOR,"div.x1y1aw1k.xn6708d > div")))
    comment= Web.find_elements(By.CSS_SELECTOR,"div.x1y1aw1k.xn6708d > div")

    for i in range(0,len(comment)):
        if comment[i].text!= "作者"and operator.not_("留言……" in comment[i].text) and operator.not_("回覆" in comment[i].text and "......" in comment[i].text):
            for k in range(0,len(Keyword)):
                if Keyword[k] in comment[i].text:
                    Comment_List.append(comment[i].text)
                    print(comment[i].text)
                    break

#  貼入目標網站
def data_login(driver):
     
    WebDriverWait(driver,20,0.5).until(EC.presence_of_element_located((By.ID,"inputUsername")))
    account= driver.find_element(By.ID,"inputUsername")
    password= driver.find_element(By.ID,"inputPassword")
    button= driver.find_element(By.ID,"signinSubmit")

    account.send_keys("cam0034")
    password.send_keys("&k6rmcm9")
    driver.execute_script("arguments[0].click();",button)

def data_scan(alt,driver):
    WebDriverWait(driver,20,0.5).until(EC.presence_of_element_located((By.ID,"inputLink")))
    alt.link= driver.find_element(By.ID,"inputLink").get_attribute("value")
    alt.select_att= driver.find_element(By.ID,"attribute").get_attribute("value")
    alt.select_vic= driver.find_element(By.ID,"viceattribute").get_attribute("value")
    '''select_vic.select_by_value("民生經濟")
    print("select_vic=",select_vic)'''

    #driver.execute_script("arguments[0].value = '你猜一下';", search_button)
    #js = 'arguments[0].removeAttribute("readonly");'
    alt.media_date= driver.find_element(By.ID,"found_date").get_attribute("value")
    '''js = 'arguments[0].removeAttribute("readonly");'
    driver.execute_script(js,media_date)
    driver.execute_script("arguments[0].value = '2021-02-20';", media_date)'''
    alt.china_media= driver.find_element(By.ID,"found_type").get_attribute("value")
    alt.media_name= driver.find_element(By.ID,"found_group").get_attribute("value")
    alt.media_account= driver.find_element(By.ID,"found_account").get_attribute("value")


    alt.source_date= driver.find_element(By.ID, "source_date").get_attribute("value")
    '''js = 'arguments[0].removeAttribute("readonly");'
    driver.execute_script(js,source_date)
    driver.execute_script("arguments[0].value = '2021-02-20';", source_date)'''
    alt.source_media= driver.find_element(By.ID,"source_type").get_attribute("value")
    alt.source_name= driver.find_element(By.ID,"source_group").get_attribute("value")
    alt.source_account= driver.find_element(By.ID,"source_account").get_attribute("value")
    alt.source_link= driver.find_element(By.ID,"source_link").get_attribute("value")

    alt.company_main= driver.find_element(By.ID,"firstgovernment").get_attribute("value")
    alt.company_vice= driver.find_element(By.ID,"secondgovernment").get_attribute("value")

def data_write(alt,driver):
    js = 'arguments[0].removeAttribute("readonly");'
    WebDriverWait(driver,20,0.5).until(EC.presence_of_element_located((By.ID,"inputTitle")))

    limit= 5
    title_string= ""
    if len(Comment_List) < 5:
        limit= len(Comment_List)

    title= driver.find_element(By.ID,"inputTitle")

    for i in range(0,limit):
        num= random.randint(0,len(Comment_List)-1)
        title_string+= Comment_List[num]
        Comment_List.pop(num)
    
    driver.execute_script("arguments[0].value = '"+ title_string+ "';", title)

    link= driver.find_element(By.ID,"inputLink")
    driver.execute_script("arguments[0].value = '"+ alt.link+ "';", link)

    select_att= Select(driver.find_element(By.ID,"attribute"))
    select_att.select_by_value(alt.select_att)
    select_vic= Select(driver.find_element(By.ID,"viceattribute"))
    select_vic.select_by_value(alt.select_vic)

    #driver.execute_script("arguments[0].value = '你猜一下';", search_button)
    #js = 'arguments[0].removeAttribute("readonly");'
    media_date= driver.find_element(By.ID,"found_date")
    driver.execute_script(js,media_date)
    driver.execute_script("arguments[0].value = '"+ alt.media_date+ "';", media_date)

    china_media= Select(driver.find_element(By.ID,"found_type"))
    china_media.select_by_value(alt.china_media)

    media_name= driver.find_element(By.ID,"found_group")
    driver.execute_script("arguments[0].value = '"+ alt.media_name+ "';", media_name)
    media_account= driver.find_element(By.ID,"found_account")
    driver.execute_script("arguments[0].value = '"+ alt.media_account+ "';", media_account)


    source_date= driver.find_element(By.ID, "source_date")
    driver.execute_script(js,source_date)
    driver.execute_script("arguments[0].value = '"+ alt.source_date+ "';", source_date)

    source_media= Select(driver.find_element(By.ID,"source_type"))
    source_media.select_by_value(alt.source_media)

    source_name= driver.find_element(By.ID,"source_group")
    driver.execute_script("arguments[0].value = '"+ alt.source_name+ "';", source_name)
    source_account= driver.find_element(By.ID,"source_account")
    driver.execute_script("arguments[0].value = '"+ alt.source_account+ "';", source_account)
    source_link= driver.find_element(By.ID,"source_link")
    driver.execute_script("arguments[0].value = '"+ alt.source_link+ "';", source_link)


    company_main= Select(driver.find_element(By.ID,"firstgovernment"))
    company_main.select_by_value(alt.company_main)
    company_vice= Select(driver.find_element(By.ID,"secondgovernment"))
    company_vice.select_by_value(alt.company_vice)


#介面---------------------------------

class Window(object):

    def __init__(self):

        self.root = tk.Tk()
        self.root.title("爬蟲")
        self.root.geometry('450x500')
        # Entry
        self.input_account = tk.Entry(self.root, width=40, font=('Courier',9))
        self.input_password = tk.Entry(self.root, width=40, font=('Courier',9))
        self.input_url = tk.Entry(self.root, width=40, font=('Courier',9))
        self.input_keyword = tk.Entry(self.root, width=40, font=('Courier',9))
        # Label
        self.label_account = tk.Label(self.root, text="帳號: ", font=('Courier',9))  
        self.label_password = tk.Label(self.root, text="密碼: ", font=('Courier',9))  
        self.label_url = tk.Label(self.root, text="請輸入網址: ", font=('Courier',9))    
        self.label_keyword = tk.Label(self.root, text="請輸入關鍵字: ", font=('Courier',9))  
        # Button
        self.start_botton = tk.Button(text = "開始",  command=self.start_event, width=30)
        self.clear_botton = tk.Button(text = "清除",  command=self.clear_event, width=30)
        self.scan_botton = tk.Button(text = "掃描",  command=self.scan_event, width=30)
        self.return_botton = tk.Button(text = "重新",  command=self.return_event, width=30)
    def gui_arrange(self):
        # Entry
        self.input_account.place(x=110, y=110, height=25)
        self.input_password.place(x=110, y=140, height=25)
        self.input_url.place(x=110, y=170, height=25)  
        self.input_keyword.place(x=110, y=200, height=25)
        # Label
        self.label_account.place(x=60, y=110)
        self.label_password.place(x=60, y=140)
        self.label_url.place(x=25, y= 170) 
        self.label_keyword.place(x=15, y= 200)
        # Botton
        self.start_botton.place(x=170, y=250, height=30,  width = 150)
        self.clear_botton.place(x=170, y=290, height=30,  width = 150)
        self.scan_botton.place(x=170, y=330, height=30,  width = 150)
        self.return_botton.place(x=170, y=370, height=30,  width = 150)
        # Web Element
        options = Options()
        self.driver=  uc.Chrome(use_subprocess= True, options=options)

        self.link= ''
        self.select_att= '請選擇'
        self.select_vic= '請選擇'
        self.media_date= ''
        self.media_name= ''
        self.media_account= ''
        self.china_media= '臉書個人專頁'
        self.source_media= '臉書個人專頁'
        self.source_date= ''
        self.source_name= ''
        self.source_account= ''
        self.source_link= ''
        self.company_main= '請選擇'
        self.company_vice= '無'


    def start_event(self):
        url = self.input_url.get()
        keyword = self.input_keyword.get()
        account = self.input_account.get()
        password = self.input_password.get()

        self.driver.get(url)
        scrape(self.driver, account, password, keyword)
        time.sleep(10)
        self.driver.get("https://btsp.servehttp.com/report/report")
        WebDriverWait(self.driver,60,0.5).until(EC.presence_of_element_located((By.ID,"inputUsername")))
        try:
            data_login(self.driver)
        except:
            print("error")
        
        data_write(self,self.driver)
        # 接口

    def clear_event(self):
        self.input_url.delete(0, 'end')
        self.input_keyword.delete(0, 'end')
        self.input_account.delete(0, 'end')
        self.input_password.delete(0, 'end')

    def scan_event(self):
        data_scan(self,self.driver)

    def return_event(self):
        self.driver.get("https://btsp.servehttp.com/report/report")
        data_write(self,self.driver)

def main():
    window = Window()
    window.gui_arrange()
    tk.mainloop()

if __name__ == '__main__':  
    main()



#--------------------------------



