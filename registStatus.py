import imaplib 
import email
from email.header import decode_header
from bs4 import BeautifulSoup as bs
import xlrd
import openpyxl
import datetime
import time



def getEmails():
    '''return array with subject / Body / requester_user'''

    user = 'aqacademy@aquinos.pt'
    pass_ = 'P@ssw0rd01'
    imap_server = 'mail.aquinos.org'

    # create an IMAP4 class with SSL 
    imap = imaplib.IMAP4_SSL(imap_server)
    # authenticate
    imap.login(user, pass_)

    status, messages = imap.select("INBOX")

    # number of top emails to fetch
    N = 9
    # total number of emails
    messages = int(messages[0])

    data = []

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
                subject_ = subject
                from_ = From

                # if the email message is multipart
                if msg.is_multipart():
                    # iterate over email parts
                    for part in msg.walk():
                        # extract content type of email
                        content_type = part.get_content_type()
                        content_disposition = str(part.get("Content-Disposition"))
                        try:
                            # get the email body
                            body = part.get_payload(decode=True).decode()
                        except:
                            pass
                        if content_type == "text/plain" and "attachment" not in content_disposition:
                            # print text/plain emails and skip attachments
                            body_ = getBodyMsg(body)
                        elif "attachment" in content_disposition:
                            # download attachment
                            filename = part.get_filename()
                            body_ = getBodyMsg(body)
                            if filename:
                                pass
                else:
                    # extract content type of email
                    content_type = msg.get_content_type()
                    # get the email body
                    body = msg.get_payload(decode=True)
                    if content_type == "text/plain":
                        # print only text email parts
                        body_ = body
                if content_type == "text/html":
                    body_ = getBodyMsg(body)

                # from / Subject / Body
                to_apend = [from_, subject_, body_]
                data.append(to_apend)

    # close the connection and logout
    imap.close()
    imap.logout()
    return data


#helpers

def getBodyMsg(body):
    html_file = body
    parsed_html = bs(html_file, features="html.parser")
    # get bodymsg -> p tag
    try:
        p_tag = parsed_html.findAll('p')
        all_ps = []
        for i in p_tag:
            if 'STATUS' in i.text.upper():
                all_ps.append(i.text)
            else:
                pass
        return all_ps

    except AttributeError:
        return body
    

data =  getEmails()


def parseData(data):
    '''Return array of dict: {
        {
           drsnumber:drsnumber;
           requester:requester;
           status:status
           date:date
        }
    }'''
    if(len(data) > 1):
        status_types = ['aprovada', 'recusada', 'finalizada', 'anulada']
        data_parsed = []
        for x in data:
            #get requester
            requester = x[0].split('<')[0]
            #get drs and status
            for i in x[2]:
                drsnumber = [int(j) for j in i.split() if j.isdigit()][0]
                #get drs status
                for status in status_types:
                    if status.upper() in i.upper():
                        status_ = status
                        data = {
                            'drsnumber':drsnumber,
                            'requester':requester,
                            'status':status_
                        }
                        data_parsed.append(data)
        return data_parsed
    print('Error getting data')
    return



drs = parseData(data)


db_file = 'DRS_DB.xlsx'


def getDrsStatus(path_db, drs_data):

    #open db
    db_file = xlrd.open_workbook(path_db).sheet_by_index(0)

    #loop to drs_data
    for _ in drs_data:
        drs_number = _['drsnumber']     
        requester = _['requester']
        status = _['status']
        
        #check if drs in db and the status
        for x in range(db_file.nrows):
            drs_n = db_file.cell_value(x, 0)

            if(type(drs_n) != str):
                if int(drs_n) == int(drs_number):
                    if status.upper() == 'FINALIZADO':
                        #check if this status are empty

                        isStatusEmpty = db_file.cell_value(x, 27)
                        if len(isStatusEmpty) == 0:
                            #add status to DB
                            file_ = openpyxl.load_workbook(path_db)
                            dbsheet = file_.active
                            dbsheet[f'AB{x + 1}'] = requester
                            today = time.strftime("%m/%d/%Y")
                            dbsheet[f'AC{x + 1}'] = datetime.datetime.strptime(today, "%m/%d/%Y")
                            file_.save(path_db)
                            file_.close()
                            print(f'DRS {drs_n} add to db with status: {status.upper()}')

                    if status.upper() == 'APROVADO':
                        #check if this status are empty

                        isStatusEmpty = db_file.cell_value(x, 29)
                        if len(isStatusEmpty) == 0:
                            #add status to DB
                            file_ = openpyxl.load_workbook(path_db)
                            dbsheet = file_.active
                            dbsheet[f'AD{x + 1}'] = requester
                            today = time.strftime("%m/%d/%Y")
                            dbsheet[f'AF{x + 1}'] = datetime.datetime.strptime(today, "%m/%d/%Y")
                            file_.save(path_db)
                            file_.close()
                            print(f'DRS {drs_n} add to db with status: {status.upper()}')

                    if status.upper() == 'RECUSADO':
                        #check if this status are empty

                        isStatusEmpty = db_file.cell_value(x, 29)
                        isStatusEmpty_2 = db_file.cell_value(x, 30)
                        if len(isStatusEmpty) == 0 and len(isStatusEmpty_2) == 0:
                            #add status to DB
                            file_ = openpyxl.load_workbook(path_db)
                            dbsheet = file_.active
                            dbsheet[f'AE{x + 1}'] = requester
                            today = time.strftime("%m/%d/%Y")
                            dbsheet[f'AF{x + 1}'] = datetime.datetime.strptime(today, "%m/%d/%Y")
                            file_.save(path_db)
                            file_.close()
                            print(f'DRS {drs_n} add to db with status: {status.upper()}')
                    
                    if status.upper() == 'ANULADO':
                        #check if this status are empty

                        isStatusEmpty = db_file.cell_value(x, 32)
                        if len(isStatusEmpty) == 0:
                            #add status to DB
                            file_ = openpyxl.load_workbook(path_db)
                            dbsheet = file_.active
                            dbsheet[f'AG{x + 1}'] = requester
                            today = time.strftime("%m/%d/%Y")
                            dbsheet[f'AH{x + 1}'] = datetime.datetime.strptime(today, "%m/%d/%Y")
                            file_.save(path_db)
                            file_.close()
                            print(f'DRS {drs_n} add to db with status: {status.upper()}')
        
        

getDrsStatus(db_file, drs_data=drs)

