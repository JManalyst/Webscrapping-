# -*- coding: utf-8 -*-
"""
Created on Thu Sep  8 12:03:31 2022

@author: john.wambugu
"""
#Import Relevant Libraries
#Make sure to have selenium installed 
#You can use pip install selenium on the terminal
#Also make sure to have downloaded the chrome websdriver & add it to a folder preferably in the prorgram files > chrome >application as specified laater on this script
#Important to note that the chrome webdriver and the chrome version to be used are compatible, otherwise the browser won't launch. To be sure, just download the latest versions of both.
from selenium import webdriver 
import time
from selenium.webdriver.common.by import By
from selenium.webdriver.common.action_chains import ScrollOrigin
from selenium.webdriver.support.select import Select
import datetime
from datetime import date, timedelta, datetime
import pandas as pd
import os
import smtplib
from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart
from selenium import webdriver 
import time
from selenium.webdriver.common.by import By
from selenium.webdriver.support.select import Select
import os

#define a function to send email
def send_email(email_recipient, email_subject, email_message):
    email_sender = "-Your email adress-"

    msg = MIMEMultipart()
    msg['From'] = email_sender
    msg['To'] = email_recipient
    msg['Subject'] = email_subject

    msg.attach(MIMEText(email_message, 'plain'))

    try:
        server = smtplib.SMTP('smtp.office365.com', 587)
        server.starttls()
        server.login("-Your email address-", "-Your password-")
        text = msg.as_string()
        server.sendmail(email_sender, email_recipient, text)
        print('email sent')
        server.quit()
    except:
        print("SMTP server connection error! Process may have been interupted")
    return
try:
    # we can also define the path that we want to keep our downloads (if we have to):
    options = webdriver.ChromeOptions()
    options.add_experimental_option ("prefs", {
    "download.default_directory": "-Your preffered donwload folder",
    "download.prompt_for_download": False,
    "download.directory_upgrade": True,
    "safebrowsing.enabled": False,
    "safebrowsing.ebabled": "false"})
    
    #Open chrome browser
    driver = webdriver.Chrome(executable_path =  "C:\Program Files\Google\Chrome\Application\chromedriver.exe", options= options); 

    ##Navigate Login using the link
    driver.get("https://registry.verra.org/app/search/VCS/All%20Projects")

    #maximize the browser window
    driver.maximize_window()

    #wait for 5 seconds. (This is just to be sure. Remember driver.get() waits untill the page is loaded)
    time.sleep(5)

    #Click on the vcu's button & wait for the page to load
    driver.find_element(By.XPATH , "//a[normalize-space()='VCUs']").click()
    time.sleep(4)

    #select the option from the issuance status dropdown
    driver.find_element(By.XPATH, "//select[@id='search_sel_issuance_status']").click()
    ddelement = Select(driver.find_element(By.XPATH, "//select[@id='search_sel_issuance_status']"))
    ddelement.select_by_visible_text("Active")

    #click on the submit button
    driver.find_element(By.XPATH, "//button[@type='submit']").click()
    print("Searching...")

    #get all the contents of the webpage
    body_text = driver.find_element(By.XPATH, "/html/body").text

    # wait untill the contents are loaded & proceed to the next step
    while "0 of 0 items" in body_text:
        time.sleep(1)
        body_text = driver.find_element(By.XPATH, "/html/body").text
    time.sleep(2)

    #back to the top of the page
    driver.execute_script("window.scrollTo(0, 220)")
    time.sleep(2)

    #click on the button to download .xlsx file
    driver.find_element(By.CSS_SELECTOR, ".fas.fa-file-excel.fa-lg.pr-2").click()
    print("Page loaded! Now waiting to download...")

    #get the downloaded file and rename it to the name of your choice
    downloaded_file = "C:\\Users\john.wambugu\\Burn Manufacturing\\Marketing - Business Intelligence\\BI Automations 2020\\John\\Carbon Webscrapping\\vcus.xlsx"
    dest_path = "C:\\Users\john.wambugu\\Burn Manufacturing\\Marketing - Business Intelligence\\BI Automations 2020\\John\\Carbon Webscrapping\Verra Active.xlsx"
    try:
        os.remove(dest_path)
    except:
        print("No previous file found")
    while os.path.isfile(downloaded_file) == False:
        time.sleep(1)
    os.rename(downloaded_file, dest_path)
    print("Operaton succeful!")
    driver.close()
    send_email("-recipient email-", "Webscrapping Succesful", "The task has been completed sucessfully")
except:
    send_email("-recpient email-", "Webscrapping Failed", "The task failed. Please re-check the code for errors.")