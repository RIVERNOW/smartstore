import time
from selenium import webdriver
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.common.by import By
import tkinter as tk
from tkinter import filedialog
import openpyxl
import requests
import bcrypt
import pybase64
import urllib.request
import urllib.parse
from datetime import datetime, timedelta
from selenium.webdriver.common.keys import Keys
import pyperclip
import pyautogui
import smtplib
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText

# Get webdriver
co = Options()
co.add_experimental_option('debuggerAddress', '127.0.0.1:9222')
driver = webdriver.Chrome(options=co)
driver.implicitly_wait(10)

driver.execute_script("""
   var elms = document.getElementsByClassName("modal fade seller-layer-modal modal-transparent has-close-check-box in");
Array.from(elms).forEach(function(element) {
    element.parentNode.removeChild(element);
});
""")

time.sleep(1)

driver.execute_script("""
   var elms = document.getElementsByClassName("modal-backdrop fade in");
Array.from(elms).forEach(function(element) {
    element.parentNode.removeChild(element);
});
""")

time.sleep(1)

driver.execute_script("""
   var elms = document.getElementsByClassName("modal-content");
Array.from(elms).forEach(function(element) {
    element.parentNode.removeChild(element);
});
""")

time.sleep(1)

driver.execute_script("""
   var elms = document.getElementsByClassName("modal-content");
Array.from(elms).forEach(function(element) {
    element.parentNode.removeChild(element);
});
""")

