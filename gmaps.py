from selenium import webdriver
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.common.by import By
import time
import openpyxl as xl
from selenium.webdriver.chrome.service import Service

service = Service()
options = webdriver.ChromeOptions()
driver = webdriver.Chrome(service=service, options=options)
time.sleep(2)
# driver = webdriver.Chrome("D:\\OLD DATA\\Selenium Drivers\\chromedriver_win32\\chromedriver.exe")
# time.sleep(2)
driver.get("https://www.google.com/maps/dir///@27.2255871,77.9455585,12z?entry=ttu")
time.sleep(5)

elem=driver.find_element(By.XPATH,'//*[@id="sb_ifc50"]/input') # to enter our location and destination
time.sleep(2)
# elem.click()
elem.send_keys("Yamunotri, Foundry Nagar, Agra, Uttar Pradesh 282006")
time.sleep(2)
elem2=driver.find_element(By.XPATH,'//*[@id="sb_ifc51"]/input')  # start point
time.sleep(2)
elem2.send_keys("Sarai rohilla railway station power office, Delhi Sarai Rohilla Station, Railway Station, Railway Officers Colony, Sarai Rohilla, Delhi, 110007")
# elem4 = driver.find_element(By.XPATH,'//*[@id="omnibox-directions"]/div/div[4]') #  enter our location
# time.sleep(1)
# elem4.click()
time.sleep(2)
# elem3=driver.find_element(By.XPATH,'//*[@id="sb_ifc52"]/input') # destination
# time.sleep(2)
# elem3.send_keys("Delhi Sarai Rohilla, Guru Gobind Singh Marg, Railway Officers Colony, Sarai Rohilla, New Delhi, Delhi 110005")
# time.sleep(1)
elem5 = driver.find_element(By.XPATH,'//*[@id="directions-searchbox-1"]/button[1]') #  click search option
elem5.click()
time.sleep(3)
elem=driver.find_element(By.XPATH,('//*[@id="section-directions-trip-0"]/div[1]/div/div[1]/div[2]/div'))
innertext = elem.get_attribute('innerHTML')
print(innertext.strip())
# elem2 = driver.find_element(By.XPATH,'//*[@id="sb_ifc51"]/input')
# elem2.click()
# elem.send_keys("Delhi")

# elem=Select(driver.find_element_by_id('sb_ifc50')).select_by_placeholder('Choose starting point, or click on the map...')
# elem.send_keys("Dehli")
# elem = find_element(By.XPATH,'/html/body/div[3]/div[9]/div[3]/div[1]/div[2]/div/div[3]/div[1]/div[1]/div[2]/div[1]/div/input')
# elem.send_keys("dehli")
# elem2 = find_element_by_xpath('//*[@id="sbsg50"]/div')
# print(elem2.is_displayed())