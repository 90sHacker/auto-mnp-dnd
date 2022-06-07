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

mtn_mnp= os.environ.get("MTN_MNP")
mtn_dnd= os.environ.get("MTN_DND")
airtel_mnp= os.environ.get("AIRTEL_MNP")
airtel_dnd= os.environ.get("AIRTEL_DND")
glo_mnp= os.environ.get("GLO_MNP")
glo_dnd= os.environ.get("GLO_DND")
mobile_mnp= os.environ.get("9MOBILE_MNP")
mobile_dnd= os.environ.get("9MOBILE_DND")
email_username= os.environ.get("EMAIL_USERNAME")
email_password= os.environ.get("EMAIL_PASSWORD")
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
                        file_name = 'mtn' + file_name
                    elif value.__contains__('glo'):
                        file_name = 'glo' + file_name
                    elif value.__contains__('airtel'):
                        file_name = 'airtel' + file_name
                    elif value.__contains__('9mobile'):
                        file_name = '9mobile' + file_name
                    
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
                    with open(Path(temp.name), "rb") as tp:
                        s3.upload_fileobj(tp, mtn_mnp, f'{dt.strftime("%Y")}{dt.strftime("%m")}{dt.strftime("%d")}.txt')
                    temp.close()
                    os.unlink(temp)

                if self.check_file_name(path.stem, ['dnd']):
                    #convert file to .txt
                    text = textract.process(str(path.absolute()))
                    temp = tempfile.NamedTemporaryFile(delete=False, suffix=path.suffix)
                    temp.write(text)
                    #upload to s3 as specified name format
                    with open(Path(temp.name), "rb") as tp:
                        s3.upload_fileobj(tp, mtn_dnd, f'{dt.strftime("%Y")}{dt.strftime("%m")}{dt.strftime("%d")}.txt')
                    temp.close()
                    os.unlink(temp)

            elif path.stem.__contains__('airtel'):
                if self.check_file_name(path.stem, ['mnp']):
                    #convert file to .txt
                    text = textract.process(str(path.absolute()))
                    temp = tempfile.NamedTemporaryFile(delete=False, suffix=path.suffix)
                    temp.write(text)
                    #upload to s3 as specified name format
                    with open(Path(temp.name), "rb") as tp:
                        s3.upload_fileobj(tp, airtel_mnp, f'{dt.strftime("%Y")}{dt.strftime("%m")}{dt.strftime("%d")}.txt')
                    temp.close()
                    os.unlink(temp)

                if self.check_file_name(path.stem, ['dnd']):
                    #convert file to .txt
                    text = textract.process(str(path.absolute()))
                    temp = tempfile.NamedTemporaryFile(delete=False, suffix=path.suffix)
                    temp.write(text)
                    #upload to s3 as specified name format
                    with open(Path(temp.name), "rb") as tp:
                        s3.upload_fileobj(tp, airtel_dnd, f'{dt.strftime("%Y")}{dt.strftime("%m")}{dt.strftime("%d")}.txt')
                    temp.close()
                    os.unlink(temp)

            elif path.stem.__contains__('glo'):
                if self.check_file_name(path.stem, ['mnp']):
                    #convert file to .txt
                    text = textract.process(str(path.absolute()))
                    temp = tempfile.NamedTemporaryFile(delete=False, suffix=path.suffix)
                    temp.write(text)
                    #upload to s3 as specified name format
                    with open(Path(temp.name), "rb") as tp:
                        s3.upload_fileobj(tp, glo_mnp, f'{dt.strftime("%Y")}{dt.strftime("%m")}{dt.strftime("%d")}.txt')
                    temp.close()
                    os.unlink(temp)

                if self.check_file_name(path.stem, ['dnd']):
                    #convert file to .txt
                    text = textract.process(str(path.absolute()))
                    temp = tempfile.NamedTemporaryFile(delete=False, suffix=path.suffix)
                    temp.write(text)
                    #upload to s3 as specified name format
                    with open(Path(temp.name), "rb") as tp:
                        s3.upload_fileobj(tp, glo_dnd, f'{dt.strftime("%Y")}{dt.strftime("%m")}{dt.strftime("%d")}.txt')
                    temp.close()
                    os.unlink(temp)
            
            elif path.stem.__contains__('9mobile'):
                if self.check_file_name(path.stem, ['mnp']):
                    #convert file to .txt
                    text = textract.process(str(path.absolute()))
                    temp = tempfile.NamedTemporaryFile(delete=False, suffix=path.suffix)
                    temp.write(text)
                    #upload to s3 as specified name format
                    with open(Path(temp.name), "rb") as tp:
                        s3.upload_fileobj(tp, mobile_mnp, f'{dt.strftime("%Y")}{dt.strftime("%m")}{dt.strftime("%d")}.txt')
                    temp.close()
                    os.unlink(temp)

                if self.check_file_name(path.stem, ['dnd']):
                    #convert file to .txt
                    text = textract.process(str(path.absolute()))
                    temp = tempfile.NamedTemporaryFile(delete=False, suffix=path.suffix)
                    temp.write(text)
                    #upload to s3 as specified name format
                    with open(Path(temp.name), "rb") as tp:
                        s3.upload_fileobj(tp, mobile_dnd, f'{dt.strftime("%Y")}{dt.strftime("%m")}{dt.strftime("%d")}.txt')
                    temp.close()
                    os.unlink(temp)
            else:
                if path.stem != '.DS_Store':
                    text = textract.process(str(path.absolute()))
                    print(path.absolute())
                    if con_dir.exists() == False:
                        con_dir.mkdir()
                    con_dir.joinpath(f'{path.stem}.txt').write_bytes(text)


if __name__ == '__main__':
    mail = FetchEmail('imap.gmail.com',
                      'bshobanke@terragonltd.com', 'uedijmdxacoxmlos')
    # respond = mail.download_attachments('Rajesh Chopra')
    result = mail.download_attachments('bshobanke2@gmail.com')
    print(result)
    mail.convert_files(result)

    mail.close_connection()
