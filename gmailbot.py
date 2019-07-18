from os import remove, rename
from os.path import splitext, dirname, abspath, exists
from time import sleep
from subprocess import Popen, CREATE_NEW_CONSOLE
from logging import getLogger, FileHandler, StreamHandler, Formatter, DEBUG, WARNING
from base64 import urlsafe_b64encode
from httplib2 import Http, ServerNotFoundError
from email.mime.text import MIMEText
from googleapiclient.discovery import build
from googleapiclient.errors import HttpError
from google_auth_oauthlib.flow import InstalledAppFlow
from pickle import load, dump
from google.auth.transport.requests import Request


def connect():
    creds = None
    SCOPES = ['https://mail.google.com/', 'https://www.googleapis.com/auth/drive']
    if exists('token.pickle'):
        with open('token.pickle', 'rb') as token:
            creds = load(token)
    if not creds or not creds.valid:
        if creds and creds.expired and creds.refresh_token:
            creds.refresh(Request())
        else:
            flow = InstalledAppFlow.from_client_secrets_file(
                'credentials.json', SCOPES)
            creds = flow.run_local_server()
        with open('token.pickle', 'wb') as token:
            dump(creds, token)

    gmail_service = build('gmail', 'v1', credentials=creds)
    drive_service = build('drive', 'v3', credentials=creds)
    return gmail_service, drive_service


def create_message(sender, to, subject, message_text):
    message = MIMEText(message_text)
    message['to'] = to
    message['from'] = sender
    message['subject'] = subject
    return {'raw': urlsafe_b64encode(message.as_string().encode()).decode()}


def send_message(user_id, message):
    try:
        message = (gmail_service.users().messages().send(userId=user_id, body=message).execute())
        return message
    except HttpError as e:
        logger.error(e)


def clean_mail():
    # Delete mail that is not from whitelisted emailadresses.
    filter = ""
    for email in whitelisted_emails:
        filter = f'{filter}-from: {email} AND '
    for retry in range(max_retries):
        try:
            messages = gmail_service.users().messages().list(userId='me', q=filter).execute().get('messages')
        except (TimeoutError, ConnectionResetError, ServerNotFoundError) as e:
            logger.error(e)
            delay_exponentially(retry, exponent)
        else:
            break
    if messages is None:
        return

    for message in messages:
        for retry in range(max_retries):
            try:
                gmail_service.users().messages().trash(userId='me', id=message['id']).execute()
            except (TimeoutError, ConnectionResetError, ServerNotFoundError) as e:
                logger.error(e)
                delay_exponentially(retry, exponent)
            else:
                break


def get_messages():
    # Fetch unread email from bot_mail
    for retry in range(max_retries):
        try:
            messages = gmail_service.users().messages().list(userId='me', q=f'label:unread subject:{subject_to_monitor} from:{master_mail}').execute().get('messages')
        except (TimeoutError, ConnectionResetError, ServerNotFoundError) as e:
            logger.error(e)
            delay_exponentially(retry, exponent)
        else:
            break

    if messages is None:
        return

    messagecontent = []
    for message in reversed(messages):
        for retry in range(max_retries):
            try:
                result = gmail_service.users().messages().get(userId='me', format='raw', id=message['id']).execute()
                gmail_service.users().messages().modify(userId='me', id=message['id'], body={'removeLabelIds': ['UNREAD']}).execute()
            except (TimeoutError, ConnectionResetError, ServerNotFoundError) as e:
                logger.error(e)
                delay_exponentially(retry, exponent)
            else:
                break
        messagecontent.append(result['snippet'])
    return messagecontent


def delay_exponentially(base, exponent):
    delay = (base + 1) ** exponent
    logger.info("Sleeping %d seconds." % delay)
    sleep(delay)


def watch_mail():
    # Check email periodically
    while True:
        logger.info("Watching email...")
        clean_mail()
        mailcontent = get_messages()
        if mailcontent is None:
            sleep(delay_between_mailcheck)
            continue
        for commandqueue in mailcontent:
            commands = commandqueue.split(',')
            for command in commands:
                command = command.strip()
                if command.lower() == 'start notepad':
                    logger.info('Starting notepad.')
                    send_message('me', create_message(bot_mail, master_mail, bot_name, 'Starting notepad.'))
                    Popen(r'notepad.exe', creationflags=CREATE_NEW_CONSOLE)
                elif command.lower() == 'stop notepad':
                    logger.info('Stopping notepad.')
                    send_message('me', create_message(bot_mail, master_mail, bot_name, 'Stopping notepad.'))
                    Popen(r'taskkill /IM notepad.exe', creationflags=CREATE_NEW_CONSOLE)
                elif command.lower() == 'update':
                    # to update the bot, upload new version of this file to the bots google drive and send the command.
                    logger.info('Starting botupdate.')
                    send_message('me', create_message(bot_mail, master_mail, bot_name, 'Starting botupdate.'))

                    try:
                        remove(f"{splitext(__file__)[0]}.BAK")
                    except FileNotFoundError as e:
                        pass
                    try:
                        rename(__file__, f"{splitext(__file__)[0]}.BAK")
                    except OSError as e:
                        logger.error(f"Error: {e}. Aborting to prevent loss of previous version of bot.")
                        break

                    for retry in range(max_retries):
                        try:
                            response = drive_service.files().list(q=f"name = '{__file__}'").execute()
                        except (TimeoutError, ConnectionResetError, ServerNotFoundError) as e:
                            logger.error(e)
                            delay_exponentially(retry, exponent)
                        else:
                            break

                    update_file = response.get('files')
                    if len(update_file) == 0:
                        logger.info('Found no new update file. Aborting.')
                        send_message('me', create_message(bot_mail, master_mail, bot_name, 'Found no new update file. Aborting.'))
                        break
                    elif len(update_file) > 1:
                        logger.info('More than one update file. Aborting.')
                        send_message('me', create_message(bot_mail, master_mail, bot_name, 'More than one update file. Aborting.'))
                        break

                    for retry in range(max_retries):
                        try:
                            request = drive_service.files().get_media(fileId=update_file[0].get('id')).execute()
                        except (TimeoutError, ConnectionResetError, ServerNotFoundError) as e:
                            logger.error(e)
                            delay_exponentially(retry, exponent)
                        else:
                            break

                    for retry in range(max_retries):
                        try:
                            with open(__file__, 'wb') as file:
                                file.write(request)
                        except OSError as e:
                            logger.error(e)
                            delay_exponentially(retry, exponent)
                        else:
                            break

                    logger.info('Done with update. Restarting.')
                    send_message('me', create_message(bot_mail, master_mail, bot_name, 'Done with update. Restarting.'))
                    Popen(rf'python {__file__}', cwd=rf'{dirname(abspath(__file__))}', creationflags=CREATE_NEW_CONSOLE)
                    return
                sleep(delay_between_mailcommands)
        sleep(delay_between_mailcheck)


logger = getLogger(__name__)
logger.setLevel(DEBUG)
file_handler = FileHandler('logfile.log')
file_handler.setLevel(WARNING)
console_handler = StreamHandler()
console_handler.setLevel(DEBUG)
formatter = Formatter('%(asctime)s %(levelname)s %(name)s Error:%(exc_info)s %(message)s')
file_handler.setFormatter(formatter)
console_handler.setFormatter(formatter)
logger.addHandler(file_handler)
logger.addHandler(console_handler)

# Retry variables
max_retries = 10
exponent = 3

# Establish connection
for retry in range(max_retries):
    try:
        gmail_service, drive_service = connect()
    except (ServerNotFoundError) as e:
        logger.error(e)
        delay_exponentially(retry, exponent)
    else:
        break

# Settings below
delay_between_mailcheck = 60
delay_between_mailcommands = 20
bot_name = 'I was sent from a bot.'  # Used as subject in response emails.

# Necessary to change:
whitelisted_emails = ('sendsmecommands@gmail.com', )  # Tuple of emailsenders to keep in bots inbox. Make sure to add master_mail to this one.
subject_to_monitor = 'hey listen'  # Emailsubject to monitor.
bot_mail = 'bottestmail@gmail.com'  # Email with gmail api and drive api activated (recommend new email).
master_mail = 'sendsmecommands@gmail.com'  # Bot only listens to commands from this email with the subject above.
watch_mail()
