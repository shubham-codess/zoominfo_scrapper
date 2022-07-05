import pandas as pd
# import urllib
# import requests
# from bs4 import BeautifulSoup
from selenium import webdriver
from time import sleep
import pyautogui
import pyperclip
import re
from webdriver_manager.chrome import ChromeDriverManager

df = pd.read_excel("Data.xlsx")
d1 = list(df['Domains or Company Name'])
emp_size= []
comp_rev = []
url = []
domain = []
revenue_size = ["Million", "Billion"]


driver = webdriver.Chrome(ChromeDriverManager().install())
for i in d1:
        if type(i) == float:
            break
        driver.get("https://www.google.com/search?q=" + str(i) + " zoominfo")
        sleep(1)

        # Tripple Dot Click

        if len(driver.find_elements_by_class_name("D6lY4c"))>0:
            driver.find_elements_by_class_name("D6lY4c")[0].click()
        
            sleep(1)

        # Cache Click

            if len(driver.find_elements_by_xpath("/html/body/div[13]/div/div/div[5]/div/span/span/a/div/span"))>0:
                driver.find_elements_by_xpath("/html/body/div[13]/div/div/div[5]/div/span/span/a/div/span")[0].click()
            else:
                emp_size.append("Not Found")
                comp_rev.append("Not Found")
                url.append("Not Found")
                domain.append("Not Found")
                continue
            sleep(1)

            # Current Url

            zf_link = driver.current_url
            try:
                linkregex = re.compile(r'[\w\S\.]+:(https://[\w\.]+.com[\w/\-]+)[\S\s]*')
                url.append(linkregex.search(zf_link).group(1))
            except AttributeError:
                linkregex = re.compile(r'[\w\S\.]+:(https://[\w\.]+.com[\w/\-]+)[\S\s]*')
                url.append(linkregex.search(zf_link).group(1))
                

            # Domain 

            if len(driver.find_elements_by_xpath("//div[@class='vertical-icons']/app-icon-text/div/a"))>0:
                domain.append(driver.find_element_by_xpath("//div[@class='vertical-icons']/app-icon-text/div/a").text)
            else:
                domain.append("Not Found")

            # EMP Size

            if len(driver.find_elements_by_class_name("company-header-subtitle"))>0:
                size = driver.find_element_by_class_name("company-header-subtitle")            
                emp_size.append(size.text.split()[-2])

            else:
                emp_size.append("Not Found")

            # Revenue
            last = []
            second_last = []
            if len(driver.find_elements_by_xpath("//div[@class='vertical-icons']//app-icon-text/div/div/span"))>0:
                revenue = driver.find_elements_by_xpath("//div[@class='vertical-icons']//app-icon-text/div/div/span")
                if len(revenue)>1:
                    if ("Million" in revenue[-1].text or "Billion" in revenue[-1].text):
                        last.append(revenue[-1].text)
                    elif ("Million" in revenue[-2].text or "Billion" in revenue[-2].text):
                        second_last.append(revenue[-2].text)
                elif len(revenue) == 1:
                    if ("Million" in revenue[-1].text or "Billion" in revenue[-1].text):
                        last.append(revenue[-1].text)
                else:
                    comp_rev.append("Not Found")
                
                if len(last)>0:
                    comp_rev.append(last[0])
                elif len(second_last)>0:
                    comp_rev.append(second_last[0])
                else:
                    comp_rev.append("Not Found")
                        
            else:
                comp_rev.append("Not Found")
        else:
            emp_size.append("Not Found")
            comp_rev.append("Not Found")
            url.append("Not Found")
            domain.append("Not Found")
print(len(emp_size))

print(len(comp_rev))

print(len(url))

print(len(domain))

dict = {"Domains": domain, "Company Size": emp_size, "Company Revenue": comp_rev, "Zoominfo URL": url}

df = pd.DataFrame.from_dict(dict)
    


df.to_csv('data_extracted.csv')

driver.close()
