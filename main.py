import urllib.parse
from selenium import webdriver
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys
import pandas as pd
import time
import urllib


WEB_DRIVER_PATH = "./chromedriver"
WEB_PAGE_TO_OPEN = "https://web.whatsapp.com/"
walletOfClients = pd.read_excel("ContatosClientes.xlsx", engine='openpyxl')
print(walletOfClients)

def prepareAndSendMessage(navigationBrowser): 

     # Logged in whats
    for i, mensagem in enumerate(walletOfClients['mensagem']):
        client = walletOfClients.loc[i, "cliente"]
        number = walletOfClients.loc[i, "telefone"]
        parsedMessage = urllib.parse.quote(f"Oi {client}! {mensagem}")
        link = f"https://web.whatsapp.com/send?phone={number}&text={parsedMessage}"
        navigationBrowser.get(link)  
        
        isLoadedWhatsPage(navigationBrowser)
        
        sendMessage(navigationBrowser)


def sendMessage(navigationBrowser):
        #Wainting load conversation
        time.sleep(5)   
        #send message
        messageBox = navigationBrowser.find_element(By.XPATH, '//*[@id="main"]/footer/div[1]/div/span/div/div[2]/div[1]/div[2]/div[1]/p')
        messageBox.send_keys(Keys.ENTER)
 
        time.sleep(2) 

def startChromeBrowser():
    driverNavigationBrowser = webdriver.Chrome(service=browserServicePath(), options=browserOptions())
    driverNavigationBrowser.get(WEB_PAGE_TO_OPEN)   

    isLoadedWhatsPage(driverNavigationBrowser) 

    return driverNavigationBrowser
        
def browserServicePath():
   return Service(executable_path=WEB_DRIVER_PATH)

def isLoadedWhatsPage(navigationBrowser): 
    while len(navigationBrowser.find_elements(By.ID, "side")) < 1:
        time.sleep(2)

    return True

def browserOptions():
    options = webdriver.ChromeOptions()
    options.add_experimental_option("detach", True)
    options.add_argument("--start-maximized")

    return options

if __name__ == "__main__":
    mainNavigationBrowser = startChromeBrowser()
    prepareAndSendMessage(mainNavigationBrowser)
    