<h1>Email Attachment Downloader</h1>

The purpose of this code is to connect to connect to a mailbox and retrieve the useful information from it.
It basically retreives the following information:
  * Email Body
  * Html Content
  * Attachment

These things are extracted and stored inside the system for further analysis in **extract_info_from_mail** function.

Apart from implementing this, I have implemented the second part of code written in **Selenium** in which the extracted information from the email body (*i.e a link address*) is used to traverse to a website in order to download the daily report. It's just the extension to logic of extracting useful information from the mail.

In the end the downlaoded report is sent as email attachment using **send_mail** helper function. We can also send a HTML or a text in the body of the mail. It also has the functionality of sending mail to mulltiple users at once.
