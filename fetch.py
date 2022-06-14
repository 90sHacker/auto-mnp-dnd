import imaplib
import email
import textract
import tempfile
import boto3
import os
from pathlib import Path
from email.header import decode_header
from datetime import datetime
from dateutil.parser import *
from dotenv import load_dotenv

load_dotenv()

base_mnp= os.environ.get("BASE_MNP")
base_dnd= os.environ.get("BASE_DND")
dnd_username= os.environ.get("DND_USERNAME")
dnd_password= os.environ.get("DND_PASSWORD")
mnp_username= os.environ.get("MNP_USERNAME")
mnp_password= os.environ.get("MNP_PASSWORD")
aws_access_key= os.environ.get("ACCESS_KEY")
aws_secret_key= os.environ.get("SECRET_KEY")





class FetchEmail():

    def __init__(self, mail_server, username, password):
        self.connection = imaplib.IMAP4_SSL(mail_server)

        self.connection.login(username, password)
        self.connection.select('Inbox', readonly=False)

    def close_connection(self):
        """
        Close connection to the IMAP server
        """
        self.connection.close()

    def search(self, value, key='HEADER FROM'):
        """
        Search the Mailbox for matching messages
        """
        print(f'({key} "{value}")')
        result, data = self.connection.search(None, f'({key} "{value}")')
        print(data, result)
        if result == 'OK':
            return result, data

    def fetch_recent_email(self, value):
        """
        Retrieve the most recent email for a search criterion
        """
        msgs = []
        result, messages = self.search(value)

        if (messages):
            # get the id's of the messages as a string
            id = email.message_from_bytes(messages[0]).as_string()
            #create an array of id's and pick the last element which is the most recent
            id = id.strip()
            id_array = id.split(' ')
            # if there is one or multiple message id's, pick the last one
            if(len(id_array) >= 1):
                message_id = id_array[-1]

        if result == 'OK':
            try:
                res, data = self.connection.fetch(message_id, "(RFC822)")
            except:
                print("No email found")
                self.close_connection()
                exit()

            msgs.append(data[0][1])
            response, data = self.connection.store(message_id, '+FLAGS', '\\Seen')

        return msgs

    def download_attachments(self, value, download_folder="attachments"):
        """
        Download the attachments for the recent email
        """

        att_err = "No current email found."
        att_dir = "No attachment found"

        msgs = self.fetch_recent_email(value)
        for response in msgs:
            msg = email.message_from_bytes(response)
            date = decode_header(msg["Date"])[0][0]
            then = parse(date).date()
            print(then)
            now = datetime.now().strftime('%Y-%m-%d')
            print(now)
            if str(then) != str(now):
                return att_err

            for part in msg.walk():
                if part.get_content_maintype() == 'multipart':
                    continue

                if part.get('Content-disposition') is None:
                    continue
                
                file_name = part.get_filename()
                if bool(file_name):
                    if value.__contains__('mtn'):
                        file_name = 'mtn_' + file_name
                    elif value.__contains__('glo'):
                        file_name = 'glo_' + file_name
                    elif value.__contains__('airtel'):
                        file_name = 'airtel_' + file_name
                    elif value.__contains__('9mobile'):
                        file_name = '9mobile_' + file_name
                    elif value.__contains__('gmail'):
                        file_name = 'gmail_' + file_name
                    
                    file_name = Path(file_name)
                    att_dir = Path(download_folder)
                    file_path = att_dir / file_name
                    if str(file_path.suffix) in ['.txt', '.xls', '.csv']:
                        if att_dir.exists() == False:
                            att_dir.mkdir()
                        file_path.write_bytes(part.get_payload(decode=True))
                        # self.convert_files(att_dir)
    
        if(isinstance(att_dir, Path)):
            return att_dir.absolute()

        return att_dir

    def check_file_name(self, file_name, keyword):
        """
        Checks a file name for a specified key
        """
        fn = file_name.lower()
        for key in keyword:
            if fn.__contains__(key):
                return True
            break
        return False

    def convert_files(self, file_path, convert_folder="converted"):
        """
        File conversion to .txt
        """
        con_dir = Path(convert_folder)
        file_path = Path(file_path)
        s3 = boto3.client('s3', aws_access_key, aws_secret_key)
        dt = datetime.now()

        for path in file_path.rglob('*.*'):
            if path.stem.__contains__('mtn'):
                if self.check_file_name(path.stem, ['mnp']):
                    #convert file to .txt
                    text = textract.process(str(path.absolute()))
                    temp = tempfile.NamedTemporaryFile(delete=False, suffix=path.suffix)
                    temp.write(text)
                    #upload to s3 as specified name format
                    try:
                        with open(Path(temp.name), "rb") as tp:
                            s3.upload_fileobj(tp, f'{base_mnp}mtn/', f'{dt.strftime("%Y")}{dt.strftime("%m")}{dt.strftime("%d")}.txt')
                        print('Upload successful!')
                        temp.close()
                        os.unlink(temp)
                    except FileNotFoundError:
                        print('File not found')
            

                if self.check_file_name(path.stem, ['dnd']):
                    #convert file to .txt
                    text = textract.process(str(path.absolute()))
                    temp = tempfile.NamedTemporaryFile(delete=False, suffix=path.suffix)
                    temp.write(text)
                    #upload to s3 as specified name format
                    try:
                        with open(Path(temp.name), "rb") as tp:
                            s3.upload_fileobj(tp, f'{base_dnd}mtn/', f'{dt.strftime("%Y")}{dt.strftime("%m")}{dt.strftime("%d")}.txt')
                        print('Upload successful!')
                        temp.close()
                        os.unlink(temp)
                    except FileNotFoundError:
                        print('File not found')

            elif path.stem.__contains__('airtel'):
                if self.check_file_name(path.stem, ['mnp']):
                    #convert file to .txt
                    text = textract.process(str(path.absolute()))
                    temp = tempfile.NamedTemporaryFile(delete=False, suffix=path.suffix)
                    temp.write(text)
                    #upload to s3 as specified name format
                    try:
                        with open(Path(temp.name), "rb") as tp:
                            s3.upload_fileobj(tp, f'{base_mnp}airtel/', f'{dt.strftime("%Y")}{dt.strftime("%m")}{dt.strftime("%d")}.txt')
                        print('Upload successful!')
                        temp.close()
                        os.unlink(temp)
                    except FileNotFoundError:
                        print('File not found')

                if self.check_file_name(path.stem, ['dnd']):
                    #convert file to .txt
                    text = textract.process(str(path.absolute()))
                    temp = tempfile.NamedTemporaryFile(delete=False, suffix=path.suffix)
                    temp.write(text)
                    #upload to s3 as specified name format
                    try:
                        with open(Path(temp.name), "rb") as tp:
                            s3.upload_fileobj(tp, f'{base_dnd}airtel/', f'{dt.strftime("%Y")}{dt.strftime("%m")}{dt.strftime("%d")}.txt')
                        print('Upload successful!')
                        temp.close()
                        os.unlink(temp)
                    except FileNotFoundError:
                        print('File not found')

            elif path.stem.__contains__('glo'):
                if self.check_file_name(path.stem, ['mnp']):
                    #convert file to .txt
                    text = textract.process(str(path.absolute()))
                    temp = tempfile.NamedTemporaryFile(delete=False, suffix=path.suffix)
                    temp.write(text)
                    #upload to s3 as specified name format
                    try:
                        with open(Path(temp.name), "rb") as tp:
                            s3.upload_fileobj(tp, f'{base_mnp}glo/', f'{dt.strftime("%Y")}{dt.strftime("%m")}{dt.strftime("%d")}.txt')
                        print('Upload successful!')
                        temp.close()
                        os.unlink(temp)
                    except FileNotFoundError:
                        print('File not found')

                if self.check_file_name(path.stem, ['dnd']):
                    #convert file to .txt
                    text = textract.process(str(path.absolute()))
                    temp = tempfile.NamedTemporaryFile(delete=False, suffix=path.suffix)
                    temp.write(text)
                    #upload to s3 as specified name format
                    try:
                        with open(Path(temp.name), "rb") as tp:
                            s3.upload_fileobj(tp, f'{base_dnd}glo/', f'{dt.strftime("%Y")}{dt.strftime("%m")}{dt.strftime("%d")}.txt')
                        print('Upload successful!')
                        temp.close()
                        os.unlink(temp)
                    except FileNotFoundError:
                        print('File not found')
            
            elif path.stem.__contains__('9mobile'):
                if self.check_file_name(path.stem, ['mnp']):
                    #convert file to .txt
                    text = textract.process(str(path.absolute()))
                    temp = tempfile.NamedTemporaryFile(delete=False, suffix=path.suffix)
                    temp.write(text)
                    #upload to s3 as specified name format
                    try:
                        with open(Path(temp.name), "rb") as tp:
                            s3.upload_fileobj(tp, f'{base_mnp}9mobile/', f'{dt.strftime("%Y")}{dt.strftime("%m")}{dt.strftime("%d")}.txt')
                        print('Upload successful!')
                        temp.close()
                        os.unlink(temp)
                    except FileNotFoundError:
                        print('File not found')
                    

                if self.check_file_name(path.stem, ['dnd']):
                    #convert file to .txt
                    text = textract.process(str(path.absolute()))
                    temp = tempfile.NamedTemporaryFile(delete=False, suffix=path.suffix)
                    temp.write(text)
                    #upload to s3 as specified name format
                    try:
                        with open(Path(temp.name), "rb") as tp:
                            s3.upload_fileobj(tp, f'{base_dnd}9mobile/', f'{dt.strftime("%Y")}{dt.strftime("%m")}{dt.strftime("%d")}.txt')
                        print('Upload successful!')
                        temp.close()
                        os.unlink(temp)
                    except FileNotFoundError:
                        print('File not found')
            else:
                if path.stem != '.DS_Store':
                    text = textract.process(str(path.absolute()))
                    print(path.absolute())
                    if con_dir.exists() == False:
                        con_dir.mkdir()
                    con_dir.joinpath(f'{path.stem}.txt').write_bytes(text)


if __name__ == '__main__':
    dnd_mail = FetchEmail('imap.gmail.com',
                      dnd_username, dnd_password)
    
    mnp_mail = FetchEmail('imap.gmail.com',
                      mnp_username, mnp_password)

    airtel_dnd_result = dnd_mail.download_attachments('airtel-dnd@terragonltd.com')
    print(airtel_dnd_result)
    mtn_dnd_result = dnd_mail.download_attachments('sdp@mtn.com')
    print(mtn_dnd_result)
    
    dnd_mail.convert_files(airtel_dnd_result)
    dnd_mail.convert_files(mtn_dnd_result)

    mtn_mnp_result = mnp_mail.download_attachments('DAAS_note_ng@mtn.com')
    print(mtn_mnp_result)

    mnp_mail.convert_files(mtn_mnp_result)

    dnd_mail.close_connection()
    mnp_mail.close_connection()
