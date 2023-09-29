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

#Promotional Message
f = open("../data/message.txt", "r", encoding="utf8")
message = f.read()
f.close()


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

    #Create sent message title
    df['MESSAGE_TITLE'] = df['NAME'].apply(lambda x: 'Dear ' + x.title() + ',' )

    #save the processed data
    df.to_excel("../data/testlist.xlsx", index = False)

    print(df.head())
    return df



def send(listname):
    #Read the contact list
    #send msg batch 200

    st = 0
    end = 2

    f1 = pd.read_excel(listname).loc[st:end]

    #add a delay between page load and sending the message
    delay = 20

    #Load chrome driver
    driver = webdriver.Chrome()

    #Load Whatsapp web and login
    print('Once your browser opens up sign in to web whatsapp')
    driver.get('https://web.whatsapp.com')
    input(style.MAGENTA + "AFTER logging into Whatsapp Web is complete and your chats are visible, press ENTER..." + style.RESET)
    sleep(2)

    
    for index, row in f1.iterrows():
        try:
            phoneNumber = str(row["MOBILENO"])
            text = row["MESSAGE_TITLE"] + message
            url = 'https://web.whatsapp.com/send?phone=' + phoneNumber + '&text=' + text
            sent = False
            for i in range(1):
                if not sent:
                    #Get the web.whatsapp.com page to compose and send the message.
                    driver.get(url)
                    #If the current window is not in focus, it is throwing error and sending message.
                    #Bring the current window into focus.
                    driver.switch_to.window(driver.window_handles[-1]); 
                    
                    try:
                        print('Sending message to {}'.format(phoneNumber))
                        click_btn = WebDriverWait(driver, delay).until(EC.visibility_of_element_located((By.XPATH, '//*[@id="main"]/footer/div[1]/div/span[2]/div/div[2]/div[2]/button/span')))
                        click_btn.click()
                        sleep(3)
                        
                    except Exception as e:
                        print(style.RED + f"\nFailed to send message to: {phoneNumber}, retry ({i+1}/3)")
                        print("Make sure your phone and computer is connected to the internet.")
                        print("If there is an alert, please dismiss it." + style.RESET)
                else:
                    sleep(3)
                    click_btn.click()
                    sent=True
                    sleep(3)
                    print(style.GREEN + 'Message sent to: ' + phoneNumber + style.RESET)
		
        except Exception as e:
            print(style.RED + 'Failed to send message to ' + phoneNumber + str(e) + style.RESET)

    driver.quit()
    return None


if __name__ == "__main__":
    prep('../data/testdata.xlsx')
    #sleep(10)
    send('../data/testlist.xlsx')