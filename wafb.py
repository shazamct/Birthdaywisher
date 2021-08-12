

from selenium import webdriver
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.common.by import By
import time
import sys

import fbchat 
from getpass import getpass
def whatsapp(name):
    
    # add path of chromedriver
    driver = webdriver.Chrome('D:\softwares\chromedriver')

    driver.get("https://web.whatsapp.com/")
    wait = WebDriverWait(driver, 600)

    # target is a person to whom message to be send
    target = name

    # birthday message
    string = "Heartly Felicitation on your Birthday "+name

    x_arg = '//span[contains(@title,' + target + ')]'
    group_title = wait.until(EC.presence_of_element_located((By.XPATH, x_arg)))
    group_title.click()


    message = driver.find_elements_by_xpath('//*[@id="main"]/footer/div[1]/div[2]/div/div[2]')[0]


    message.send_keys(string)

    sendbutton = driver.find_elements_by_xpath('//*[@id="main"]/footer/div[1]/div[3]/button')[0]
    sendbutton.click()

    driver.close()

def facebook(name):
    
    username = str("saurabh.deshmukh.9279") #enter login ID
    client = fbchat.Client(username, getpass()) 
    name = str(name) #friend's name who is having birthday
    friends = client.searchForUsers(name)  # return a list of names 
    friend = friends[0] 
    msg = str("Heartly felicitation on your birthday"+name) 
    sent = client.send(friend.uid, msg) 
    if sent: 
        print("Message sent successfully!")     
    
