import imaplib
import datetime
import email
def Check_Reply(recipients,subject,email_user,email_pass):
    mail = imaplib.IMAP4_SSL('imap.gmail.com')
    mail.login(email_user, email_pass)
    mail.list()
    mail.select('inbox')
    result, data = mail.uid('search', None, '(FROM "'+recipients+'" SUBJECT "'+subject+'")')
    try:
        uid = data[0].split()[0]
        return True
    except:
        return False

#this function is only to write file for a good format
def Format_Write_File(file_handle,column_name):
            if len(str(column_name))>13:
                file_handle.write(str(column_name)+(30-(len(str(column_name))-13))*" ")
            elif len(str(column_name))<13:
                file_handle.write(str(column_name)+(30+(13-len(str(column_name))))*" ")
            elif len(str(column_name))==13:
                file_handle.write(str(column_name)+" "*30)


def Fetch_Unreply():
    #all month hash
    all_month={"Jan":1,"Feb":2,"Mar":3,"Apr":4,"May":5,"Jun":6,"Jul":7,"Aug":8,"Sep":9,"Oct":10,"Nov":11,"Dec":12}
    #time span last week
    date_last_week = datetime.date.today() - datetime.timedelta(7)
    file_handle=open("email_not_reply_list.txt","a+")
    file_handle.write("All people that hasn't reply to your email yet!!!!!!\n\n")
    file_handle.write("Sent Date"+" "*40)
    file_handle.write("Receiver"+" "*32)
    file_handle.write("Subject"+" "*30+"\n\n\n")
    email_user=raw_input("Input your email account please:")
    email_pass=raw_input("Input the password of your account:")
    mail = imaplib.IMAP4_SSL('imap.gmail.com')
    mail.login(email_user, email_pass)
    mail.list()
    mail.select('[Gmail]/Sent Mail')
    status, data = mail.search(None, "ALL")
  #  date = (datetime.date.today() - datetime.timedelta(1000)).strftime("%d-%b-%Y")
  #  status, data = mail.uid('search', None, '(SENTSINCE '+date+')')
    mails = data[0].split()
    if data[0] != '':
       for num in mails:
           tpe, raw_msg = mail.fetch(num, '(RFC822)')
           try:
               msg = email.message_from_string(raw_msg[0][1])
               sbj, ecode = email.Header.decode_header(msg['subject'])[0]
            #   print sbj
               #from, sender
               frm = ''
               for fts, ecode in email.Header.decode_header(msg['Date']):
                   frm = frm + fts
                   year=msg['Date'][12:16];month=all_month[msg['Date'][8:11]];day=msg['Date'][5:7]
                   sent_date=datetime.date(int(year),int(month),int(day))
                   if sent_date>date_last_week and Check_Reply(msg['to'],sbj,email_user,email_pass) is False:
                        Format_Write_File(file_handle,str(sent_date))
                        Format_Write_File(file_handle,msg['to'])
                        Format_Write_File(file_handle,sbj)
                        file_handle.write("\n")

                   #print msg['Date']
           except:
               pass
    print "successful to write in text file!"

Fetch_Unreply()