'''
Send emails to a list of contacts through an Office365 exchage account.
Example:
> python office365_send_emails.py config.json contacts.txt email.html
'''
import sys
import json
import smtplib
import time
from os import path
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from email.header import Header

host = 'smtp.office365.com'
port = '587'
delay_between_messages = 5  # Office365 has a limit of 30 messages per minute


def smtp_connect(email, password):
    smtp = smtplib.SMTP(host, port)
    smtp.starttls()
    smtp.login(email, password)
    return smtp


def send_emails(
    email,
    password,
    from_email,
    subject,
    email_body,
    email_type,
    contacts
):
    smtp = smtp_connect(email, password)

    index = 0
    count = len(contacts)
    last_error = None
    while index < count:
        contact = contacts[index]
        msg = MIMEMultipart('alternative')
        msg['From'] = from_email
        msg['To'] = contact
        msg['Subject'] = subject
        msg.attach(MIMEText(email_body, email_type))

        try:
            smtp.send_message(msg)
            print(f'{contact} - OK')
            last_error = None
        except smtplib.SMTPAuthenticationError as e:
            print(f'{contact} - FAILED (SMTPAuthenticationError: {e})')
            last_error = 'smtplib.SMTPAuthenticationError'
        except smtplib.SMTPDataError as e:
            print(f'{contact} - FAILED (SMTPDataError: {e})')
            last_error = 'smtplib.SMTPDataError'
        except smtplib.SMTPConnectError as e:
            print(f'{contact} - FAILED (SMTPConnectError: {e})')
            last_error = 'smtplib.SMTPConnectError'
        except smtplib.SMTPHeloError as e:
            print(f'{contact} - FAILED (SMTPHeloError: {e})')
            last_error = 'smtplib.SMTPHeloError'
        except smtplib.SMTPSenderRefused as e:
            print(f'{contact} - FAILED (SMTPSenderRefused: {e})')
            last_error = 'smtplib.SMTPSenderRefused'
        except smtplib.SMTPResponseException as e:
            print(f'{contact} - FAILED (SMTPResponseException: {e})')
            last_error = 'smtplib.SMTPResponseException'
        except smtplib.SMTPServerDisconnected as e:
            if last_error is not 'smtplib.SMTPServerDisconnected':
                last_error = 'smtplib.SMTPServerDisconnected'
                print(f'{contact} - FAILED (SMTPServerDisconnected: {e}) '
                      '[RETRY]')
                smtp = smtp_connect(email, password)
                continue
            last_error = 'smtplib.SMTPServerDisconnected'
            print(f'{contact} - FAILED (SMTPServerDisconnected: {e})')
        except smtplib.SMTPRecipientsRefused as e:
            print(f'{contact} - FAILED (SMTPRecipientsRefused: {e})')
            last_error = 'smtplib.SMTPRecipientsRefused'
        except smtplib.SMTPNotSupportedError as e:
            print(f'{contact} - FAILED (SMTPNotSupportedError: {e})')
            last_error = 'smtplib.SMTPNotSupportedError'
        except smtplib.SMTPException as e:
            print(f'{contact} - FAILED (SMTPException: {e})')
            last_error = 'smtplib.SMTPException'
        except Exception as e:
            print(f'{contact} - FAILED (Exception: {e})')
            last_error = 'Exception'

        del(msg)
        index += 1
        time.sleep(delay_between_messages)

    smtp.quit()


if __name__ == '__main__':
    config_file = sys.argv[1]
    contacts_file = sys.argv[2]
    email_file = sys.argv[3]

    config = {}
    with open(config_file, 'r', encoding='utf-8') as f:
        config = json.loads(f.read())

    contacts = []
    with open(contacts_file, 'r', encoding='utf-8') as f:
        contacts = f.readlines()
        contacts = [contact.strip().replace('\n', '')
                    for contact in contacts if contact != '\n']

    email_body = ''
    with open(email_file, 'r', encoding='utf-8') as f:
        email_body = f.read()

    send_emails(
        config['email'],
        config['password'],
        config['from_email'],
        config['subject'],
        email_body,
        config['email_type'],
        contacts
    )

    print('Done.')
