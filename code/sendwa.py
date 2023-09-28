import pandas as pd
from selenium import webdriver
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.common.by import By
from webdriver_manager.chrome import ChromeDriverManager
from time import sleep
from urllib.parse import quote
import os
import pickle




#Styling
os.system("")
os.environ["WDM_LOG_LEVEL"] = "0"
class style():
    BLACK = '\033[30m'
    RED = '\033[31m'
    GREEN = '\033[32m'
    YELLOW = '\033[33m'
    BLUE = '\033[34m'
    MAGENTA = '\033[35m'
    CYAN = '\033[36m'
    WHITE = '\033[37m'
    UNDERLINE = '\033[4m'
    RESET = '\033[0m'

print(style.BLUE)
print("**********************************************************")
print("**********************************************************")
print("*****                                               ******")
print("*****  THANK YOU FOR USING WHATSAPP BULK MESSENGER  ******")
print("*****      This tool Sends Bulk Whatsapp Msg        ******")
print("**** https://github.com/eztiaapps/bulkclientmessage ******")
print("*****                                               ******")
print("**********************************************************")
print("**********************************************************")
print(style.RESET)

def prep(filename):
    df = pd.read_excel(filename)
    
    #Columns name shouldn't have space
    df.columns = df.columns.str.replace(' ','')

    #Rows that has values
    df = df[df['MOBILENO'].notna()]

    #First name respect
    df['FN'] = df['NAME'].apply(lambda x: x.split(' ')[0])

    #Create sent message title
    df['MESSAGE_TITLE'] = df['FN'].apply(lambda x: 'Dear ' + x.title() + ',' )

    #save the processed data
    df.to_excel("../data/AundhList.xlsx", index = False)

    print(df.head())
    return df


def send():
    #Read the contact list
    #send msg batch 200

    f1 = pd.read_excel('../data/AundhList.xlsx').loc[0:199]

    #add a delay between page load and sending the message
    delay = 20
    print(f1)

    return f1


if __name__ == "__main__":
    #prep('../data/AUNDH.xlsx')
    send()