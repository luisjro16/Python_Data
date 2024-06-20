from selenium import webdriver
from selenium.webdriver.common.by import By
import time

url = 'https://www.instagram.com/'

driver = webdriver.Chrome()
driver.get(url)

pesquisa = driver.find_element(By.XPATH, '//*[@id="loginForm"]/div/div[1]/div/label/input')
pesquisa.send_keys("luis_jro")
pesquisa = driver.find_element(By.XPATH, '//*[@id="loginForm"]/div/div[2]/div/label/input')
pesquisa.send_keys("digital246")
pesquisa = driver.find_element(By.XPATH, '//*[@id="loginForm"]/div/div[3]/button/div')
pesquisa.click()
time.sleep(20)
