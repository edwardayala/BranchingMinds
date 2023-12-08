# %%
import time, sys, os
from selenium import webdriver
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
import pandas as pd

def resource_path(relative_path: str) -> str:
    try:
        base_path = sys._MEIPASS
    except Exception:
        base_path = os.path.dirname(__file__)

    return os.path.join(base_path, relative_path)

# Read XLSX file to get Group & Notes using pandas
df = pd.read_excel(resource_path('OpportunityReports.xlsx'), sheet_name='Sheet1')

# Save the Group & Notes columns to a dictionary where the key is the Group and the value is the Notes
groupNotes = df.set_index('Group')['Notes'].to_dict()

# Start the Chrome browser
# browser = webdriver.Chrome(resource_path('./chromedriver'))

# Create an instance of Options
chrome_options = Options()

# Set the path to the chromedriver
webdriver_service = Service(resource_path('chromedriver'))

# Pass the Options instance and Service instance into the webdriver
browser = webdriver.Chrome(service=webdriver_service, options=chrome_options)

# Navigate to the mapnewa website
browser.get('https://fwisd.branchingminds.com/#/login')

# Wait until the page fully loads by waiting for the "Sign in with Classlink" link to appear
# Wait for the "Sign in with Classlink" link to appear
wait = WebDriverWait(browser, 10)
linkElem = wait.until(EC.presence_of_element_located((By.LINK_TEXT, "Sign in with Classlink")))

# Wait for the page to load
time.sleep(3.75)
### Authentication Start ###
# Click the link with the text "Sign in with Classlink"
linkElem = browser.find_element(By.LINK_TEXT, "Sign in with Classlink")
# Click the link
linkElem.click()

# Wait 3/4 a second
time.sleep(0.75)
loginHereButton = wait.until(EC.presence_of_element_located((By.XPATH, "/html/body/div[2]/div[1]/div[2]/div[2]/div[3]/button")))
# Find "LOGIN HERE" button - full xpath: /html/body/div[2]/div[1]/div[2]/div[2]/div[3]/button
loginHereButton = browser.find_element(By.XPATH, "/html/body/div[2]/div[1]/div[2]/div[2]/div[3]/button")
# Click "LOGIN HERE" button
loginHereButton.click()

# Wait 3/4 a second
time.sleep(0.75)

signInWithWindowsIcon = wait.until(EC.presence_of_element_located((By.XPATH, "/html/body/div[2]/div/div[2]/div[3]/a[1]")))
# Find and click Sign in with Windows icon - full xpath: /html/body/div[2]/div/div[2]/div[3]/a[1]
signInWithWindowsIcon = browser.find_element(By.XPATH, "/html/body/div[2]/div/div[2]/div[3]/a[1]")
# Click Sign in with Windows icon
signInWithWindowsIcon.click()
# Wait 3/4 a second
time.sleep(0.75)

# Wait for user to login, continue automation when URL changes to https://fwisd.branchingminds.com/#/students/home
while browser.current_url != 'https://fwisd.branchingminds.com/#/students/home':
    # Wait 1/4 a second
    time.sleep(0.25)
    # print('Waiting for user to login...')

# emailElem = wait.until(EC.presence_of_element_located((By.ID, 'i0116')))
# # Find email input (input tag and id = email)
# emailElem = browser.find_element(By.ID, 'i0116')
# emailElem.send_keys('enriqueta.hernandez@fwisd.org')

# # Wait 3/4 a second
# time.sleep(0.75)
# submitElem = wait.until(EC.presence_of_element_located((By.ID, 'idSIButton9')))
# # Find submit button (input tag and id = idSIButton9)
# submitElem = browser.find_element(By.ID, 'idSIButton9')
# submitElem.click()

# # Wait 3/4 a second
# time.sleep(0.75)
# passwordElem = wait.until(EC.presence_of_element_located((By.NAME, 'passwd')))
# # Find password input (input tag and name = passwd)
# passwordElem = browser.find_element(By.NAME, 'passwd')
# passwordElem.send_keys('Edward.10042018')

# # Wait 3/4 a second
# time.sleep(0.75)
# submitPwdElem = wait.until(EC.presence_of_element_located((By.ID, 'idSIButton9')))
# # Find submit button (input tag and id = idSIButton9)
# submitPwdElem = browser.find_element(By.ID, 'idSIButton9')
# submitPwdElem.click()
# # Wait 3/4 a second
# time.sleep(0.75)

# yesElem = wait.until(EC.presence_of_element_located((By.ID, 'idSIButton9')))
# # If the "Stay signed in?" page appears, click "Yes"
# try:
#     yesElem = browser.find_element(By.ID, 'idSIButton9')
#     yesElem.click()
# except:
#     pass
# # Wait 3.5 seconds
# time.sleep(3.5)
### Authentication End ###

### Automated Notes Start ###

# Go to todo's : https://fwisd.branchingminds.com/#/activity/todos
browser.get('https://fwisd.branchingminds.com/#/activity/todos')

# Wait 3 seconds
time.sleep(3)
# %%
# For each key in the dictionary of Group & Notes
for groupname in groupNotes:
    # Wait 3/4 a second
    time.sleep(0.75)
    print('Entering notes for {}'.format(groupname))
    print('-------------------------')
    print('Note: {}'.format(groupNotes[groupname]))
    print('-------------------------')
# Find the the Notes button inside a div with a link that has the text of the Group in it
    try:
        link = browser.find_element(By.XPATH, "//a[normalize-space()='{}']".format(groupname))
    except:
        print('Error on Group Lookup:')
        print('Group {} not found, skipping...'.format(groupname))
        continue
    
    try:
        # Find the "Note for All" button that is a following sibling of the link
        note_for_all_button = link.find_element(By.XPATH, "following-sibling::div/button[normalize-space()='Note for All']")
    except:
        print('Error on Note for All Lookup:')
        print('Note for All button not found for {}, skipping...'.format(groupname))
        continue

    # Check if the button is disabled
    if note_for_all_button.get_attribute('disabled'):
        # If it is, skip to the next group
        continue

    try:
        # Click the button
        note_for_all_button.click()
    except:
        print('Error on Note for All Click:')
        print('Note for All button not clickable for {}, skipping...'.format(groupname))
        continue
    # Wait 3/4 a second
    time.sleep(0.75)

    # %%
    try:
        # Find the text area for the notes 
        notesTextArea = browser.find_element(By.XPATH, "//textarea[@id='notes']")
        # Enter the notes
        notesTextArea.send_keys(groupNotes[groupname])
    except:
        print('Error on Notes Lookup:')
        print('Notes text area not found for {}, skipping...'.format(groupname))
        continue
    # Wait 3/4 a second
    time.sleep(1.75)

    # %%
    # Find the Cancel button
    cancelButton = browser.find_element(By.XPATH, "//button[normalize-space()='Cancel']")
    try:
        # Find the Ok button
        okButton = notesTextArea.find_element(By.XPATH, "//button[normalize-space()='Ok']")
    except:
        print('Error on Ok button Lookup:')
        print('Ok button not found for {}, skipping...'.format(groupname))
        continue
    # Click Cancel - Debug
    # cancelButton.click()
    try:
        # Click Ok
        okButton.click()
        print('Ok button clicked for {}'.format(groupname))
    except:
        print('Error on Ok button Click:')
        print('Ok button not clickable for {}, skipping...'.format(groupname))
        continue
    # Wait 3/4 a second
    time.sleep(1.75)

    # %%
    try:
        # Find "Mark Done" button
        markDoneButton = link.find_element(By.XPATH, "following-sibling::div/button[normalize-space()='Mark Done']")
    except:
        print('Error on Mark Done Lookup:')
        print('Mark Done button not found for {}, skipping...'.format(groupname))
        continue

    # Check if the button is disabled
    if markDoneButton.get_attribute('disabled'):
        # If it is, skip to the next group
        continue
    try:
        # Click "Mark Done" button
        # markDoneButton.click()
        print('Mark Done button clicked for {}'.format(groupname))
    except:
        print('Error on Mark Done Click:')
        print('Mark Done button not clickable for {}, skipping...'.format(groupname))
        continue

    # Wait 3/4 a second
    time.sleep(0.75)

### Automated Notes End ###
# %%
