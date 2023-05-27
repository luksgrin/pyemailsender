# Global imports
import smtplib
from os import path
from time import sleep
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from email.mime.application import MIMEApplication

# Class definitions
class EmailSenderBase():

    def __init__(self, config, delay=0):

        # Email configuration
        self._smtp_server = config["SMTP_SERVER"]
        self._sender   = config["SENDER"]
        self._password = config["PASSWORD"]

        # Sleep config
        self._delay = delay

    # Check if the SMTP server comes from microsoft
    def __isMicrosoftSMTP(self, candidate):
        return ("outlook" in candidate) or ("office365" in candidate)
    
    def _connect(self, SMTPclass):

        # Instantiate SMTP server
        smtp_obj = SMTPclass(
            self._smtp_server,
            self._smtp_PORT
        )

        return smtp_obj
            
    
    def __createEmail(self, receiver, subject, message, CC=None, BCC=None, attachmentPaths=None):

        # Create the email message
        msg = MIMEMultipart()
        msg["From"] = self._sender
        msg["To"] = receiver
        msg["Subject"] = subject

        if CC:
            if isinstance(CC, list):
                CC = ", ".join(CC)
            msg['Cc'] = CC

        if BCC: 
            if isinstance(BCC, list):
                BCC = ", ".join(BCC)
            msg['Bcc'] = BCC

        msg.attach(MIMEText(message, "plain"))

        # Attach multiple files if provided
        if attachmentPaths:
            
            if isinstance(attachmentPaths, str):
                attachmentPaths = [attachmentPaths]

            for attachmentPath in attachmentPaths:

                filename = path.split(attachmentPath)[-1]

                with open(attachmentPath, "rb") as attachment:
                    part = MIMEApplication(
                        attachment.read(),
                        Name=filename
                    )
                    part["Content-Disposition"] = f"attachment; filename={filename}"
                    msg.attach(part)

        return msg

    def _sendEmails(self, smtp_obj, messagesDict):

        for k,v in messagesDict.items():

            if not(
                {"receiver", "subject", "message"}.issubset(set(v.keys()))
            ):
                print(f"Incomplete arguments for {k}...\nSkipping...")

            msg = self.__createEmail(**v)
            
            # Send the email
            smtp_obj.send_message(msg)

            print(f"Email sent to {k}")

            sleep(self._delay)

    def sendEmails(self, messagesDict):
        smtp_obj = self.__connect()
        self._sendEmails(smtp_obj, messagesDict)

        # Disconnect from the SMTP server
        smtp_obj.quit()
        return None
    
class EmailSender_TLS(EmailSenderBase):

    def __init__(self, config, delay=0):

        super().__init__(config, delay)

        try:
            self._smtp_PORT = int(config["SMTP_PORT"])
        except:
            self._smtp_PORT = 587

    def __connect(self):

        try:
            smtp_obj = super()._connect(self, smtplib.SMTP)

            # For microsoft... sigh
            if super().__isMicrosoftSMTP(self._smtp_server):
                smtp_obj.ehlo('mylowercasehost')
                smtp_obj.starttls()
                smtp_obj.ehlo('mylowercasehost')
            else:
                smtp_obj.starttls()
                
            smtp_obj.login(self._sender, self._password)

        except Exception as e:
            raise Exception(
                f"Something went wrong during the connection. See:\n\n{e}"
            )

        return smtp_obj
    
    def sendEmails(self, messagesDict):
        smtp_obj = self.__connect()
        self._sendEmails(smtp_obj, messagesDict)

        # Disconnect from the SMTP server
        smtp_obj.quit()
        return None
    
class EmailSender_SSL(EmailSenderBase):

    def __init__(self, config, delay=0):

        super().__init__(config, delay)

        try:
            self._smtp_PORT = int(config["SMTP_PORT"])
        except:
            self._smtp_PORT = 465

    def __connect(self):

        try:
            smtp_obj = super()._connect(smtplib.SMTP_SSL)
            smtp_obj.login(self._sender, self._password)

        except Exception as e:
            raise Exception(
                f"Something went wrong during the connection. See:\n\n{e}"
            )

        return smtp_obj
    
    def sendEmails(self, messagesDict):
        smtp_obj = self.__connect()
        self._sendEmails(smtp_obj, messagesDict)

        # Disconnect from the SMTP server
        smtp_obj.quit()
        return None
