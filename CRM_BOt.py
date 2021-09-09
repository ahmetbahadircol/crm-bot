from selenium import webdriver
from selenium.webdriver.common import keys
import time
import xlwt 
from xlwt import Workbook
from datetime import date
import datetime
import user_info
import version_chromedriver #personel class import
class CRM:
    def __init__(self,username,password):
        try: #check for chromedriver version or existing the chromedriver.exe file
            self.Browser = webdriver.Chrome("C:/Users/x/Desktop/Python/Atiker_CRM/chromedriver.exe")
        except: #if version is not the lastest or not existing the chromedriver.exe file, download it
            version_chromedriver.main()
            self.Browser = webdriver.Chrome("C:/Users/x/Desktop/Python/Atiker_CRM/chromedriver.exe")
        self.Browser.maximize_window()
        self.username=username
        self.password=password
    def signin(self):
        self.Browser.get('http://crm.atikeryazilim.com.tr/Login')
        print('-------------------------------')
        print('Log-in')
        print('-------------------------------')
        usernameInput = self.Browser.find_element_by_xpath("//*[@id='KULLANICI_ADI']")
        passwordInput = self.Browser.find_element_by_xpath("//*[@id='PAROLA']")
        print('Enter info')
        usernameInput.send_keys(self.username)
        passwordInput.send_keys(self.password)
        btnSubmit = self.Browser.find_element_by_xpath("//*[@id='BtnGiris']")
        btnSubmit.click()
        print('Log-in succesful...')
        #time.sleep(2)

    def OpenReport(self):
        print('Open Left Menu')
        btnSubmit = self.Browser.find_element_by_xpath("/html/body/div[1]/div/div[1]/ul/li[4]/a")
        print(btnSubmit)
        btnSubmit.click()
        #time.sleep(1)
        print('-------------------------------')
        print('Menu is expanding')
        print('-------------------------------')
        btnSubmit = self.Browser.find_element_by_xpath("//*[@id='SolMenu_ulMenu']/li[4]/ul/li[1]/a")
        btnSubmit.click()
        wb = Workbook() 
        sheet1 = wb.add_sheet('Sheet 1')
        sheet1.write(0, 0, "Tarih") 
        sheet1.write(0, 1, "Görüşen Kişi")
        sheet1.write(0, 2, "Cari Adı")
        sheet1.write(0, 3, "Konu")
        sheet1.write(0, 4, "Başlangıç Saat")
        sheet1.write(0, 5, "Bitiş Saat")
        row = 1
        page = 1
        onluk = 1
        until_date = user_info.until_date
        flag = True
        if_flag = True
        while flag:
            self.Browser.execute_script("window.scrollTo(0,0);")
            for tr in range(2,52):
                print('Enter for loop')
                btnSubmit = self.Browser.find_element_by_xpath("/html/body/div[3]/div/form/div[3]/div[2]/div/div/div[2]/div/span/div[3]/div/div/div/div/table/tbody/tr["+str(tr)+"]/td[1]/img") # + butonuna basılır
                btnSubmit.click()
                tarih = self.Browser.find_element_by_xpath("//*[@id='GrdSurecRapor']/tbody/tr["+str(tr)+"]/td[3]").text
                print(tarih)
                gorusen = self.Browser.find_element_by_xpath("//*[@id='GrdSurecRapor']/tbody/tr["+str(tr)+"]/td[4]").text
                print(gorusen)
                cari = self.Browser.find_element_by_xpath("//*[@id='GrdSurecRapor']/tbody/tr["+str(tr)+"]/td[5]").text
                print(cari)
                konu = self.Browser.find_element_by_xpath("/html/body/div[3]/div/form/div[3]/div[2]/div/div/div[2]/div/span/div[3]/div/div/div/div/table/tbody/tr["+str(tr+1)+"]/td[2]/div/table/tbody/tr[2]/td[2]").text
                print(konu)
                bas_saat = self.Browser.find_element_by_xpath("/html/body/div[3]/div/form/div[3]/div[2]/div/div/div[2]/div/span/div[3]/div/div/div/div/table/tbody/tr["+str(tr+1)+"]/td[2]/div/table/tbody/tr[2]/td[3]").text
                print(bas_saat)
                bit_saat = self.Browser.find_element_by_xpath("/html/body/div[3]/div/form/div[3]/div[2]/div/div/div[2]/div/span/div[3]/div/div/div/div/table/tbody/tr["+str(tr+1)+"]/td[2]/div/table/tbody/tr[2]/td[4]").text
                print(bit_saat)
                btnSubmit = self.Browser.find_element_by_xpath("/html/body/div[3]/div/form/div[3]/div[2]/div/div/div[2]/div/span/div[3]/div/div/div/div/table/tbody/tr["+str(tr)+"]/td[1]/img") # - butonuna basılır
                btnSubmit.click()
                print(onluk,' ',page,' ',tarih,' ',konu)
                cari_list = [
                         'ALFA - BETA MAK. İML. İNŞ. MÜH. TUR. İTH. İHR. SAN. VE TİC. LTD. ŞTİ.'
                        ,'APEKS HAVACILIK TASARIM SAV. MÜH. DANŞ. LTD. ŞTİ.'
                        ,'HEZARFEN SAVUNMA SAN. HAVACILIK MAK. İNŞ. OTO. SAN. VE TİC. LTD. ŞTİ.'
                        ,'ERVE SAVUNM. MAK. İMLT. SAN. VE TİC. LTD. ŞTİ.'
                        ,'KMT SAVUNMA MAK. METAL İTH. İHR. SAN. VE TİC. A.Ş.'
                        ,'TOLGA PLASTİK SAN. VE TİC. LTD. ŞTİ.'
                        ,'ORİON HAVACILIK MAK. METAL SAN. VE TİC. LTD. ŞTİ.'
                ]
                if datetime.datetime.strptime(tarih.replace(".", "-"), '%d-%m-%Y') <= datetime.datetime.strptime(until_date.replace(".", "-"), '%d-%m-%Y'):#until_date:
                    print('Reached due date...')
                    flag = False
                    if_flag = False
                    break
                if cari in cari_list: #(gorusen == 'bahadir.col') or (gorusen == 'ozkan.keskin'): 
                    sheet1.write(row, 0, tarih) 
                    sheet1.write(row, 1, gorusen)
                    sheet1.write(row, 2, cari)
                    sheet1.write(row, 3, konu)
                    sheet1.write(row, 4, bas_saat)
                    sheet1.write(row, 5, bit_saat)
                    row += 1
            if page != 10 and if_flag:
                if onluk == 1:
                    x = 1
                else:
                    x = 2
                btnSubmit = self.Browser.find_element_by_xpath("/html/body/div[3]/div/form/div[3]/div[2]/div/div/div[2]/div/span/div[3]/div/div/div/div/table/tbody/tr[52]/td/table/tbody/tr/td["+str(page+x)+"]/a") # Sayfa atlanır
                btnSubmit.click()
                time.sleep(25)
                page += 1
            elif page == 10 and if_flag:
                if onluk == 1:
                    x = 11
                else:
                    x = 12
                btnSubmit = self.Browser.find_element_by_xpath("/html/body/div[3]/div/form/div[3]/div[2]/div/div/div[2]/div/span/div[3]/div/div/div/div/table/tbody/tr[52]/td/table/tbody/tr/td["+str(x)+"]/a") # 10luk Sayfa atlanır
                btnSubmit.click()
                time.sleep(25)
                page = 1
                onluk += 1
        today = date.today()
        d1 = today.strftime("%Y%m%d")
        try:
            wb.save('C:/Users/x/Desktop/Python/Atiker_CRM/'+str(d1)+'.xls')
        except:
            wb.save('abc.xls')
        self.Browser.quit()
             




username = user_info.username
password = user_info.password
CRM = CRM(username,password)
CRM.signin()
CRM.OpenReport()
print('Done...')