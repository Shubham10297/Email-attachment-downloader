from logging import error
from typing import List, Tuple
import boto3
import base64
from botocore.exceptions import ClientError
from email.mime.text import MIMEText
from email.mime.application import MIMEApplication
from email.mime.multipart import MIMEMultipart
import io
import json
import smtplib


def run_query(conn, query) -> list[tuple]:
    cursor = conn.cursor()
    cursor.execute(query)
    results = cursor.fetchall()
    cursor.close()
    return results


def export_csv(df, index=False):
    with io.StringIO() as buffer:
        df.to_csv(buffer, index=index)
        return buffer.getvalue()


def validate_html(html: str) -> str:
    if "<body>" not in html:
        html = "<body>"+html+"</body>"
    if "<html>" not in html:
        html = "<html>"+html+"</html>"

    return html


def send_mail(email_id: str, email_pass: str, source: str, email_ids: list[str], subject: str, html_part: str = None, body_text: str = None, as_attachment: bool = False, attachment_file=None, attachment_filename: str = None) -> None:
    """
    This is a custom function written to send mail to multiple or single users.
    It also has the feature to include any attachment in the mail.

    Arguments Information

    email_id : The email id from which the mail needs to be sent.
    email_pass : The email password.
    email_ids: These are the list of email addresses to which the mail needs to be sent
    subject: The subject of the mail.
    html_part: The html part that needs to be included inside the mail. Eg. <p></b>This is paragraph which is bold.</b></p>
    body_text: The text part which needs to be included inside the mail.
    as_attachment: This needs to be True if there is an attachment to be sent.
    attachment_file: The Bytes IO object of the file to be attached.
    attachment_filename: The name of the attached file.


    """

    CHARSET = "UTF-8"

    # creates SMTP session
    client = smtplib.SMTP("imap-mail.outlook.com", 993)
    # start TLS for security
    client.starttls()
    # Authentication
    client.login("sender_email_id", "sender_email_id_password")

    message = MIMEMultipart('mixed')
    message['Subject'] = subject
    message['From'] = source
    message['To'] = ", ".join(email_ids)
    message_body = MIMEMultipart('alternative')

    if html_part and body_text:
        html_part = validate_html(html_part)
        html_part = html_part.replace(
            "</body>", "<p><b>{}</b></p></body>".format(body_text))
        htmlpart = MIMEText(html_part, 'html', CHARSET)
        message_body.attach(htmlpart)
    elif body_text:
        textpart = MIMEText(body_text, "plain", CHARSET)
        message_body.attach(textpart)

    message.attach(message_body)

    try:
        if as_attachment:
            attachmentpart = MIMEApplication(attachment_file)
            attachmentpart.add_header(
                'Content-Disposition', 'attachment', filename=attachment_filename)
            message.attach(attachmentpart)

        response = client.send_raw_email(Source=source, Destinations=email_ids, RawMessage={
                                         'Data': message.as_string(), })
    except ClientError as e:
        print(e.response['Error']['Message'])
    else:
        print("Email sent! Message ID:")
