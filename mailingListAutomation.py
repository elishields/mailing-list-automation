#################
## mailwave.py ##
#################

# This script copies email addresses from a spreadsheet and pastes them into dynamic form

# TODO ignore unrecognized emails and print list at end

from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.support.select import Select
from openpyxl import load_workbook
import time

###################
## SET VARIABLES ##
###################

# real file path has been removed
excel_file = 'C:\\Users\\...\\file.xlsx'
spreadsheet_sheet_name = 'Sheet4'
# real file path has been removed
chromedriver = "C:\\Users\\...\\chromedriver.exe"
# real url has been removed
form_url = "https://realformurlremoved.com
authenticate_email = "email@form.com"

#######################
## CREATE EMAIL LIST ##
#######################

workbook = load_workbook(excel_file, data_only=True)
sheet_name = spreadsheet_sheet_name
spreadsheet = workbook[sheet_name]

email_list_dirty = []

for cell in spreadsheet['C']:
    email_list_dirty.append(cell.value)

######################
## CLEAN EMAIL LIST ##
######################

email_list_clean = [email.strip(' ') for email in email_list_dirty]
email_list_clean = [email.strip(',') for email in email_list_dirty]
email_list_clean = [email.strip(';') for email in email_list_dirty]

############################
## AUTHENTICATE INTO FORM ##
############################

browser = webdriver.Chrome(chromedriver)
browser.get(mailwave_url)

login_element = browser.find_element_by_name('loginfmt')
login_element.send_keys(authenticate_email)
login_element.send_keys(Keys.ENTER)

time.sleep(20)

############################
## ENTER EMAILS INTO FORM ##
############################

input_element = browser.find_element_by_css_selector('input.form-control.ng-untouched.ng-pristine.ng-valid')
input_element.click()

for email in email_list:
	input_element.send_keys(email)
    time.sleep(5)
	confirm_element = browser.find_element_by_css_selector('button.list-group-item.list-group-item-action')
    if confirm_element == none:
        print email, " not found\n"
    else:
        confirm_element.click()
	input_element = browser.find_element_by_css_selector('input.form-control.ng-valid.ng-dirty.ng-touched')
    time.sleep(1)
