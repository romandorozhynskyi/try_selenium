#import
import xlrd
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.support.wait import WebDriverWait
import time
import keyboard

#qty of specifications = a
a = 22
b = a + 1

#exel
loc = (r"C:\RD_projects\Python\QMS\QMS.xls")
wb = xlrd.open_workbook(loc) 
sheet = wb.sheet_by_index(0) 

#edge
browser = webdriver.Edge(r"C:\RD_projects\Python\QMS\msedgedriver.exe")
time.sleep(10)
browser.get("https://qms.steripackgroup.com/")
time.sleep(10)
button_login = browser.find_element_by_xpath("/html/body/app-root/ng-component/div/div/div/div/div/div/div/ng-component/div/div[2]/div/div/div/img")
button_login.click()

#select page documents
time.sleep(10)
qms_option = browser.find_element_by_xpath("/html/body/app-root/ng-component/div/default-layout/div[1]/div/button[1]/span")
qms_option.click()
time.sleep(5)
dok_mod_click = browser.find_element_by_xpath("/html/body/app-root/ng-component/div/default-layout/div[2]/div/div[1]/div/side-bar-menu/div/div/ul/li[2]/div/ul/li[1]/a/span[3]")
dok_mod_click.click()

i = 1
while i < b:

    #select new document
    time.sleep(5)
    button_klick_new_document = browser.find_element_by_xpath("/html/body/app-root/ng-component/div/default-layout/div[2]/div/div[2]")
    button_klick_new_document.click()
    button_new_document = browser.find_element_by_xpath("/html/body/app-root/ng-component/div/default-layout/div[2]/div/div[2]/div[2]/ng-component/div/div/sub-header/div/div/div[2]/div/button[2]")
    button_new_document.click()

    #filing first_tab
    dt1 = str(sheet.cell_value(i, 0))
    text_Document_Title = browser.find_element_by_xpath("/html/body/app-root/ng-component/div/default-layout/div[2]/div/div[2]/div[2]/view-or-edit-document-record-component/div/div/div/div/div/div[1]/div/document-core-fields/form/div[1]/input")
    text_Document_Title.send_keys("" + dt1)
    category = browser.find_element_by_xpath("/html/body/app-root/ng-component/div/default-layout/div[2]/div/div[2]/div[2]/view-or-edit-document-record-component/div/div/div/div/div/div[1]/div/document-core-fields/form/div[4]/storm-dropdown/div/kendo-dropdownlist/span/span[1]")
    category.click()
    keyboard.write("Lab\n")
    dr1 = str(sheet.cell_value(i, 1))
    text_Document_Ref = browser.find_element_by_xpath("/html/body/app-root/ng-component/div/default-layout/div[2]/div/div[2]/div[2]/view-or-edit-document-record-component/div/div/div/div/div/div[1]/div/document-core-fields/form/div[2]/input")
    text_Document_Ref.send_keys("" + dr1)
    traning = browser.find_element_by_xpath("/html/body/app-root/ng-component/div/default-layout/div[2]/div/div[2]/div[2]/view-or-edit-document-record-component/div/div/div/div/div/div[2]/kendo-tabstrip/div[1]/document-training-details/div/div[1]/storm-dropdown[1]/div/kendo-dropdownlist/span/span[1]")
    traning.click()
    keyboard.write("N\n")
    s1 = str(sheet.cell_value(i, 2))
    text_Summary = browser.find_element_by_xpath("/html/body/app-root/ng-component/div/default-layout/div[2]/div/div[2]/div[2]/view-or-edit-document-record-component/div/div/div/div/div/div[1]/div/document-core-fields/form/div[5]/storm-editor/div/textarea")
    text_Summary.send_keys("" + s1)
   
    #Approvers_tab
    time.sleep(5)
    approv_tab = browser.find_element_by_xpath("/html/body/app-root/ng-component/div/default-layout/div[2]/div/div[2]/div[2]/view-or-edit-document-record-component/div/div/div/div/div/div[2]/kendo-tabstrip/ul/li[2]/span")
    approv_tab.click()
    approv_field = browser.find_element_by_xpath("/html/body/app-root/ng-component/div/default-layout/div[2]/div/div[2]/div[2]/view-or-edit-document-record-component/div/div/div/div/div/div[2]/kendo-tabstrip/div[2]/div/storm-selector/div/kendo-autocomplete/kendo-searchbar/input")
    approv_field.send_keys("Bartosz" + Keys.ENTER)

    #wait
    time.sleep(5)
    approv_field2 = browser.find_element_by_xpath("/html/body/app-root/ng-component/div/default-layout/div[2]/div/div[2]/div[2]/view-or-edit-document-record-component/div/div/div/div/div/div[2]/kendo-tabstrip/div[2]/div/storm-selector/div/kendo-autocomplete/kendo-searchbar/input")
    approv_field2.send_keys("" + Keys.BACKSPACE)
    time.sleep(5)
    approv_field3 = browser.find_element_by_xpath("/html/body/app-root/ng-component/div/default-layout/div[2]/div/div[2]/div[2]/view-or-edit-document-record-component/div/div/div/div/div/div[2]/kendo-tabstrip/div[2]/div/storm-selector/div/kendo-autocomplete/kendo-searchbar/input")
    approv_field3.send_keys("" + Keys.ENTER)

    #upload
    time.sleep(5)
    upload_button = browser.find_element_by_xpath("/html/body/app-root/ng-component/div/default-layout/div[2]/div/div[2]/div[2]/view-or-edit-document-record-component/div/div/sub-header/div/div/div[2]/div/document-toolbar/div/div/div[2]/button[2]")
    upload_button.click()

    #input file
    time.sleep(5)
    keyboard.write("" + dr1 + ".pdf \n")

    #save and send
    time.sleep(10)
    save_button = browser.find_element_by_xpath("/html/body/app-root/ng-component/div/default-layout/div[2]/div/div[2]/div[2]/view-or-edit-document-record-component/div/div/sub-header/div/div/div[2]/div/document-toolbar/div/div/div[2]/button[1]")
    save_button.click()
    time.sleep(30)
    summit_button = browser.find_element_by_xpath("/html/body/app-root/ng-component/div/default-layout/div[2]/div/div[2]/div[2]/view-or-edit-document-record-component/div/div/sub-header/div/div/div[2]/div/document-toolbar/div/div/div[2]/button")
    summit_button.click()
    time.sleep(10)
    ok_button = browser.find_element_by_xpath("/html/body/app-root/ng-component/div/default-layout/div[2]/div/div[2]/div[2]/view-or-edit-document-record-component/div/div/div/send-for-approval-confirmation-modal/div/div/div/div[3]/button[2]")
    ok_button.click()

    #to start menu
    time.sleep(10)
    qms_option = browser.find_element_by_xpath("/html/body/app-root/ng-component/div/default-layout/div[1]/div/button[1]/span")
    qms_option.click()
    time.sleep(5)
    dok_mod_click = browser.find_element_by_xpath("/html/body/app-root/ng-component/div/default-layout/div[2]/div/div[1]/div/side-bar-menu/div/div/ul/li[2]/div/ul/li[1]/a/span[3]")
    dok_mod_click.click()
    i +=1

browser.quit()