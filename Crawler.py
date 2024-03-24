from selenium import webdriver
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.common.by import By
import time
from openpyxl import Workbook
from bs4 import BeautifulSoup
import socket


class Crawler:
    def CheckInternet(self):
        try:
            socket.create_connection(("www.google.com", 80))
            return True
        except OSError:
            pass
        return False

    def GrabCompanyInfo(self):
        if self.CheckInternet():
            driver = webdriver.Chrome()
            response = driver.get("https://findbiz.nat.gov.tw/fts/query/QueryBar/queryInit.do")
            search_box = driver.find_element(By.ID, "qryCond")
            search_box.send_keys("悟智股份有限公司")
            search_box.send_keys(Keys.RETURN)
            time.sleep(3)
            search_link = driver.find_element(By.LINK_TEXT, "悟智股份有限公司")
            search_link.click()
            time.sleep(3)
            html_content = driver.page_source
            driver.close()
            # 儲存整個網頁
            # with open('saved_page.html', 'w', encoding='utf-8') as file:
            #     file.write(html_content)
            soup = BeautifulSoup(html_content, 'html.parser')
            # 印出整個網頁
            #print(soup.prettify())
            # 印出網頁部分資訊
            company_name_tag = soup.find("td", string="公司名稱")
            capital_tag = soup.find("td", string="資本總額(元)")
            company_name = company_name_tag.find_next_sibling("td").get_text(strip=True)
            capital = capital_tag.find_next_sibling("td").get_text(strip=True)
            print("公司名稱:", company_name)
            print("資本總額(元):", capital)
        else:
            print("請確認網路連線!")




if __name__ == '__main__':
    Object = Crawler()
    Object.GrabCompanyInfo()

