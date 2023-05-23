"""Importing Email"""
import os
import smtplib
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from scripts.core.schema.model import Email
from scripts.logging.logger import logger
from scripts.exceptions.exception_codes import Email_HandlerException


def send_email(body, email: Email):
    """Function to send email"""
    sender_email = os.environ.get("sender_email")
    sender_password = os.environ.get("sender_password")
    receiver_email = email.rec_email

    message = MIMEMultipart()
    message["From"] = sender_email
    message["To"] = receiver_email
    message["Subject"] = "Total price of billing items"

    message.attach(MIMEText(body, "html"))

    try:
        logger.info("Handlers:send_email")
        server = smtplib.SMTP("smtp.gmail.com", 587)
        server.starttls()
        server.login(sender_email, sender_password)

        server.send_message(message)

        server.quit()
        logger.info("send_email:Email sent successfully")
        return {"message": "Email sent successfully"}
    except Exception as e:
        logger.error(Email_HandlerException.EX006.format(error=str(e)))
        return {"message": str(e)}
