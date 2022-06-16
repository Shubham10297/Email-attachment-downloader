import os
from selenium import webdriver
from tempfile import mkdtemp
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from helpers import *
import traceback
import time
from bs4 import BeautifulSoup
import email
import imaplib
from email.header import decode_header


def configure_chrome_options():
    """
    These are the chrome options which can configured according to needs.
    """
    download_path = mkdtemp()
    options = webdriver.ChromeOptions()
    options.binary_location = '/opt/chrome/chrome'
    options.add_argument('--headless')
    options.add_argument('--no-sandbox')
    options.add_argument("--disable-gpu")
    options.add_argument("--window-size=1280x1696")
    options.add_argument("--single-process")
    options.add_argument("--disable-dev-shm-usage")
    options.add_argument("--disable-dev-tools")
    options.add_argument("--no-zygote")
    options.add_argument(f"--user-data-dir={mkdtemp()}")
    options.add_argument(f"--data-path={mkdtemp()}")
    options.add_argument(f"--disk-cache-dir={mkdtemp()}")
    options.add_argument("--remote-debugging-port=9222")

    prefs = {"download.default_directory": download_path,
             'download.prompt_for_download': False,
             'download.directory_upgrade': True,
             'safebrowsing.enabled': False,
             'safebrowsing.disable_download_protection': True,
             'profile.default_content_setting_values.automatic_downloads': 1,
             "profile.default_content_settings.popups": 0,
             }

    options.add_experimental_option("prefs", prefs)

    return options, download_path


def enter_name_details(driver: webdriver.Chrome, username, password, wait):
    email_field = wait.until(EC.element_to_be_clickable((By.NAME, "username")))
    email_field.send_keys(username)

    next_button = wait.until(EC.element_to_be_clickable(
        (By.ID, "idp-discovery-submit")))
    next_button.click()

    password_field = wait.until(
        EC.element_to_be_clickable((By.NAME, "password")))
    password_field.send_keys(password)

    login_button = wait.until(EC.element_to_be_clickable(
        (By.ID, "okta-signin-submit")))
    login_button.click()


def Download_Excel(driver: webdriver.Chrome, wait, download_path):
    """
    This will download the Excel file from the site. This is a custom function.
    """
    download_button = wait.until(
        EC.element_to_be_clickable((By.ID, "downloadReport")))
    download_button.click()

    excel_button = wait.until(
        EC.element_to_be_clickable((By.ID, "exportExcelLink")))
    excel_button.click()

    os.chdir(download_path)
    wait = WebDriverWait(driver, 180)
    time.sleep(10)
    status = wait.until(EC.text_to_be_present_in_element(
        (By.XPATH, '//*[@id="downloadManager-grid"]/div[2]/table/tbody/tr[1]/td[4]'), "FINISHED"))

    if status:
        print("File loaded...")
        download_button_final = wait.until(EC.element_to_be_clickable(
            (By.XPATH, '//*[@id="downloadManager-grid"]/div[2]/table/tbody/tr[1]/td[5]/div/a[1]')))
        download_button_final.click()
    else:
        raise Exception(
            "Could not download the file as the status is showing pending from the last 3 mins.")

    file_not_found = True
    while(file_not_found):
        if os.listdir():
            x = os.listdir()[0]
            if 'Purchase Report' in x and 'crdownload' not in x:
                print("file downloaded")
                file_not_found = False
                driver.close()


def main(username, password, senders_mail, subcribers_emails, bucket_name, browse_link, email_id, email_pass):
    try:

        if browse_link == None:
            raise Exception("Could not find the browse link.")

        options, download_path = configure_chrome_options()

        print(download_path)

        driver = webdriver.Chrome("/opt/chromedriver", options=options)

        wait = WebDriverWait(driver, 60)

        driver.get(browse_link)

        enter_name_details(driver, username=username,
                           password=password, wait=wait)

        Download_Excel(driver, wait, download_path)

        file_name = os.listdir()[0]

        _content = open(os.path.join(download_path, file_name), "rb")

        send_mail(email_id=email_id, email_pass=email_pass, source=senders_mail, email_ids=subcribers_emails, subject="Report Downloaded",
                  body_text=f"The file name is {file_name}", attachment_file=_content, as_attachment=True, attachment_filename="Report.xlsx")
    except Exception as e:
        print(traceback.print_exc())
        send_mail(email_id=email_id, email_pass=email_pass, source=senders_mail, email_ids=subcribers_emails, subject="Error",
                  body_text=f"The error is : {e}")


def get_link(body):
    """
    This is a custom function to get the link from the email body. The link will redirect to a site from which the reports 
    needs to be downloaded.
    """
    browse_link = None
    msg = email.message_from_bytes(body)
    if msg.is_multipart():
        # iterate over email parts
        for part in msg.walk():
            # extract content type of email
            content_type = part.get_content_type()
            print(content_type)
            content_disposition = str(part.get("Content-Disposition"))
            print(content_disposition)
            try:
                # get the email body
                body = part.get_payload(decode=True).decode()
                print(body)
            except:
                pass
            if content_type == "text/html" and "attachment" not in content_disposition:
                soup = BeautifulSoup(body)
                browse_link = soup.findAll('a')[0].get('href')
                if 'sysco/purchase' in browse_link:
                    return browse_link
    else:
        content_type = msg.get_content_type()
        # get the email body
        body = msg.get_payload(decode=True).decode()
        print(body)
        if content_type == "text/html":
            soup = BeautifulSoup(body)
            browse_link = soup.findAll('a')[0].get('href')
            if 'sysco/purchase' in browse_link:
                return browse_link


def extract_info_from_mail(username, password):
    """
    This is a function which extract all the useful things from the email.
    This included attachments which will be saved inside a directory,
    HTML content which will be saved as INDEX.html file in a directory,
    the mail body which will get printed.
    One can set the number of emails to be parsed inside the mailbox in the variable N.

    Returning the mail body as futher things needs to be extracted from it.

    """

    imap = imaplib.IMAP4_SSL("imap-mail.outlook.com", 993)
    # authenticate
    imap.login(username, password)

    status, messages = imap.select("INBOX")
    # number of top emails to fetch. This is one as I want to see the latest mail only
    N = 1
    # total number of emails
    messages = int(messages[0])

    for i in range(messages, messages-N, -1):
        # fetch the email message by ID
        res, msg = imap.fetch(str(i), "(RFC822)")
        for response in msg:
            if isinstance(response, tuple):
                # parse a bytes email into a message object
                msg = email.message_from_bytes(response[1])
                # decode the email subject
                subject, encoding = decode_header(msg["Subject"])[0]
                if isinstance(subject, bytes):
                    # if it's a bytes, decode to str
                    subject = subject.decode(encoding)
                # decode email sender
                From, encoding = decode_header(msg.get("From"))[0]
                if isinstance(From, bytes):
                    From = From.decode(encoding)

                datestring = msg['date']
                # if the email message is multipart
                if msg.is_multipart():
                    # iterate over email parts
                    for part in msg.walk():
                        # extract content type of email
                        content_type = part.get_content_type()
                        content_disposition = str(
                            part.get("Content-Disposition"))
                        try:
                            # get the email body
                            body = part.get_payload(decode=True).decode()
                        except:
                            pass
                        if content_type == "text/plain" and "attachment" not in content_disposition:
                            # print text/plain emails and skip attachments
                            print(body)
                        elif "attachment" in content_disposition:
                            # download attachment
                            filename = part.get_filename()
                            if filename:

                                # enter the folder name where you want to download the attachment
                                folder_name = r"C:\Users\Shubham Sharma\Downloads"
                                filepath = os.path.join(folder_name, filename)

                                # download attachment and save it
                                open(filepath, "wb").write(
                                    part.get_payload(decode=True))
                else:
                    # extract content type of email
                    content_type = msg.get_content_type()
                    # get the email body
                    body = msg.get_payload(decode=True).decode()
                    if content_type == "text/plain":
                        # print only text email parts
                        print(body)
                if content_type == "text/html":
                    print(body)
                    # if it's HTML, create a new HTML file and open it in browser
                    folder_name = r"C:\Users\Shubham Sharma\Downloads"
                    filename = "index.html"
                    filepath = os.path.join(folder_name, filename)
                    # write the file
                    open(filepath, "w").write(body)
                    # open in the default browser

    # close the connection and logout
    imap.close()
    imap.logout()

    return body


def handler():
    # username for the site from which report needs to be downloaded
    username_site = str(os.environ['Login_Id'])
    # password for the site from which report needs to be downloaded
    password_site = str(os.environ['Password'])

    # mail address from which the email will be sent
    # Email account credentials
    email_id = "enter your email id"
    email_pass = "Enter you email password"

    # enter the mail id to whom you want to send the Email notifications
    subcribers_emails = ["shubham@gmail.com"]
    body = extract_info_from_mail(email_id, email_pass)

    browse_link = get_link(body)

    main(username_site, password_site, email_id,
         subcribers_emails, browse_link, email_id, email_pass)


if __name__ == "__main__":
    handler()
