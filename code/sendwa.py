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

def prep(fn):
    df = pd.read_excel('../data/testdata.xlsx')
    
    #Columns name shouldn't have space
    df.columns = df.columns.str.replace(' ','')

    #Rows that has values
    df = df[df['MOBILENO'].notna()]

    #First name respect
    df['FN'] = df['NAME'].apply(lambda x: x.split(' ')[0])

    #Create sent message title
    df['MESSAGE_TITLE'] = df['FN'].apply(lambda x: 'Dear ' + x.title() + ',' )

    print(df.head())
    return df


if __name__ == "__main__":
    prep('../data/testdata.xlsx')