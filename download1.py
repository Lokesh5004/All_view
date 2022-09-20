import shutil

from selenium import webdriver
import time
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.common.by import By
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.chrome.options import Options
from datetime import datetime, timedelta
import win32com.client
from dateutil import tz

current_time = datetime.today().astimezone(tz.gettz('Asia/Kolkata'))
# other libraries to be used in this script
import os
#aom_production_support@groups.nokia.com

email_config = {
    'to': 'vaishnavi.duddula.ext@nokia.com',
    'cc': 'lokesh.vudugu.ext@nokia.com',
}

print(os.path.join(os.path.dirname(os.path.realpath(__file__)),"autoreport"))
file_path = os.path.dirname(os.path.realpath(__file__))
attachmentPath = os.path.join(file_path, "autoreport/")
imagePath = os.path.join(file_path, "amberss/")
driverPath =  os.path.join(file_path, "chromedriver_win32/chromedriver.exe")

#attachmentPath = r"C:/Users/vudugu/PycharmProjects/allview/autoreport/"
#imagePath = r"C:/Users/vudugu/PycharmProjects/allview/amberss/"
outlook = win32com.client.Dispatch('outlook.application')

chrome_options = Options()
chrome_options.add_argument("--headless")

zipAndEmailtime = ""
internalExternal = "internal"
browser = webdriver.Chrome(executable_path=driverPath,chrome_options=chrome_options)
#browser = webdriver.Chrome(executable_path="C:\\Users\\vudugu\\Downloads\\chromedriver_win32\\chromedriver.exe",chrome_options=chrome_options)


def elementSelector(xpath):
    return WebDriverWait(browser, 3).until(EC.element_to_be_clickable((By.XPATH, xpath)))


def whole():
    # time.sleep(5)

    # get element
    # element = browser.find_element_by_xpath("//*[@id='contentBox']/div/div/form/div[1]/div/div/div/div[2]/input")

    elementSelector("//*[@id='ctl01']/div[4]/div[1]/div[3]/div[1]/h3/span").click()
    elementSelector("//*[@id='ctl01']/div[4]/div[1]/div[3]/div[1]/ul/li[2]/a").click()
    elementSelector("//*[@id='MainContent_txtNetworkOrder']").send_keys("12345678")
    elementSelector("//*[@id='MainContent_btnSearchOrder']").click()
    elementSelector("//*[@id='MainContent_grdSAPOrders']/tbody/tr[2]/td[1]/a").click()
    elementSelector("//*[@id='MainContent_btnWholeOrderDownload']").click()
    browser.execute_script("divexpandcollapse('div_5889702')")
    browser.execute_script(
        'document.querySelector("#div_5889702 > table > tbody > tr:nth-child(1) > td > input[type=checkbox]").click()')
    browser.execute_script(
        'document.querySelector("#div_5889702 > table > tbody > tr:nth-child(2) > td > input[type=checkbox]").click()')
    browser.execute_script('document.querySelector("#MainContent_txtEmail").value="lokesh.vudugu.ext@nokia.com"')
    elementSelector("//*[@id='MainContent_btnZipandEmail']").click()
    browser.save_screenshot(os.path.join(imagePath, internalExternal +".png"))
    global zipAndEmailtime
    # ZipandEmailTime=datetime.now()
    zipAndEmailtime = (datetime.now() - timedelta(minutes=1)).strftime('%Y-%m-%d %H:%M')
    elementSelector("//*[@id='MainContent_btnAlert']").click()
    browser.delete_all_cookies()

    # print(element)
    # options=chrome_options)
    # browser.execute_script("divexpandcollapse('div_5889702')")



def external(url):
    browser.get(url)
    elementSelector("//*[@id='contentBox']/div/div/form/div[1]/div/div/div/div[2]/input").send_keys("Arpa")
    # WebDriverWait(browser, 9).until(EC.element_to_be_clickable(
    # (By.XPATH, "//*[@id='contentBox']/div/div/form/div[1]/div/div/div/div[2]/input"))).send_keys("Arpa")
    elementSelector("//*[@id='contentBox']/div/div/form/div[2]/div/div/div/div[2]/input").send_keys("Atbtmation@2022")
    elementSelector("//*[@id='logInbutton']/input").click()
    print("loged in")
    elementSelector("//*[@id='MainContent_btncookie']").click()
    global internalExternal
    internalExternal = "external"


def internal(url):
    browser.get(url)
    elementSelector('//*[@id="txtUser"]').send_keys("Arpa")
    elementSelector('//*[@id="txtPassword"]').send_keys("Atbtmation@2022")
    elementSelector('//*[@id="submit"]').click()
    browser.execute_script("window.document.body.scrollTo(0, document.body.scrollHeight);")
    elementSelector('//*[@id="MainContent_btnOk"]').click()
    browser.execute_script("window.document.body.scrollTo(0, document.body.scrollHeight);")
    #browser.execute_script("window.scrollTo(0, document.body.scrollHeight);")
    elementSelector('//*[@id="MainContent_btnTerms"]').click()
    elementSelector('//*[@id="MainContent_btncookie"]').click()
    global internalExternal
    internalExternal = "internal"


def save_attachments(messages):
    # for message in messages:

    for attachment in messages[0].Attachments:
        print(attachment.FileName)
        attachment.SaveAsFile(os.path.join(attachmentPath, internalExternal + str(attachment)))
        #attachment.SaveAsFile(r"C:\\Users\\vudugu\\Desktop\\autoreport" + attachment.FileName)


def removeAttachments(path):
    for filename in os.listdir(path):
        file_path = os.path.join(path, filename)
        try:
            if os.path.isfile(file_path) or os.path.islink(file_path):
                os.unlink(file_path)
            elif os.path.isdir(file_path):
                shutil.rmtree(file_path)
        except Exception as e:
            print('Failed to delete %s. Reason: %s' % (file_path, e))


def sendmail(to, cc, subject, body, attachmentPath,amberAlert):
    mail = outlook.CreateItem(0)
    mail.To = to
    mail.CC = cc
    mail.Subject = subject
    mail.Body = body
    for file in os.listdir(attachmentPath):
        mail.Attachments.Add(os.path.join(attachmentPath, file))
    if amberAlert:
        for file in os.listdir(imagePath):
            mail.Attachments.Add(os.path.join(imagePath, file))
    mail.Send()


def sentAlert(isInternalMailReceived , isExternalMailReceived):
    sub = ""
    body=""
    amberAlert= False
    if isInternalMailReceived:
        if isExternalMailReceived:
            print("green alert")
            sub = "GREEN ALERT"
            body = " PFA, mails received successfully."
        else:
            print("red alert")
            sub = "RED ALERT"
            body = "External server is not working and Internal is working fine. Do the needful thing for internal"

    else:
        if isExternalMailReceived:
            print("red alert")
            sub = "RED ALERT"
            body = "Internal server is not working and External is working fine. Do the needful thing for internal"
        else:
            print("amber alert")
            sub = "AMBER ALERT"
            amberAlert = True
            body = "we are not receiving mails from any of the server. Please look into the issue & do the needful."
    sub = "AllView- Whole Order Download Functionality Monitoring | " + datetime.now().strftime(
        '%Y/%m/%d %H:%M') + " | " + sub

    body = '''Hi All, 
We have completed AllView- Whole Order Functionality monitoring on both URLs.
''' + body + '''
1. https://allview.nokia.com
2.https://allview.int.net.nokia.com'''
    sendmail(email_config['to'], email_config['cc'], sub, body, attachmentPath, amberAlert)


def checkMail():
    isMailReceived = False
    for i in range(6):
        time.sleep(30)
        messages = outlook.GetNamespace("MAPI").GetDefaultFolder(6).Items
        #print(messages)
        a = "[ReceivedTime] >= '" + zipAndEmailtime + "'"
        print(a)
        messages.sort("ReceivedTime", True)
        messages = messages.Restrict(a)
        messages = messages.Restrict("[SenderEmailAddress] = 'allview-support@list.nokia.com'")
        if len(messages) != 0:
            isMailReceived = True
            save_attachments(messages)
            # for message in messages:
            # print(message.ReceivedTime)
            break
    return isMailReceived


external("https://allview.nokia.com/")
whole()
isExternalMailReceived = checkMail()
time.sleep(60)

internal("https://allview.int.net.nokia.com/")
whole()
isInternalMailReceived = checkMail()
#isInternalMailReceived = False
#isExternalMailReceived = False
sentAlert(isInternalMailReceived , isExternalMailReceived)
removeAttachments(attachmentPath)
removeAttachments(imagePath)
browser.close()