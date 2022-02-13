#!/usr/bin/env python3
# coding: utf-8

#Программа напоминалка, рассылает по email-ам указанным в файле xslx в поле email напоминания о приеме ССТУ в заданное время. 
#Затем проверяет служебный почтовый ящик на ответы пользователей, если ответ не получен или получен ответ "ПРИНЯТО"
#сообщает о результатах пользователям указанным в поле back_email
 

# In[1]:


import pandas as pd
import datetime, pytz
import email, smtplib, imaplib, ssl
from email.mime.text import MIMEText
from email.iterators import typed_subpart_iterator
from email.header import Header, decode_header, make_header
import hashlib
from time import sleep

TIMESTART = datetime.datetime.now()
SECFRSTART = (datetime.datetime.now() - TIMESTART).total_seconds()
LASTREM = TIMESTART

#email_server = '11.11.11.11'
email_server = 'mail.email.local'



def main():
    
    global TIMESTART, SECFRSTART, LASTREM
    
    print('Sheduler start working ' + t2s(datetime.datetime.now()) + '\n')
    
    
    
    
    #reminder_rules = ['9:00', 1]

    reminder_rules = ['8:30', 1]

    #imap = imaplib.IMAP4_SSL('mail.fss.local')
    #imap.login('fss\mailrobot.10', 'mailrobot10')
    #imap.list()
    #imap.select('INBOX')
    
    dfs = pd.read_excel('./email-reminder.xlsx', sheet_name='shedule2')         
    shedule = create_shedule(dfs, reminder_rules)    
    
    
    print_shedule(shedule)
    
    md5hash0 = md5hash('./email-reminder.xlsx')
    md5hash1 = md5hash0

    last = last_date(shedule)
    
    #while datetime.datetime.now() < last:
    while datetime.datetime.now() < datetime.datetime(2131, 6, 22, 19, 0):    
        
        sleep(1)
        
        md5hash1 = md5hash('./email-reminder.xlsx') #проверка изменения файла
        
        if md5hash1 != 'md5err' and md5hash1 != md5hash0:
            try:
                dfs = pd.read_excel('./email-reminder.xlsx', sheet_name='shedule2')
                shedule = create_shedule(dfs, reminder_rules)
                md5hash0 = md5hash1
                last = last_date(shedule)
                print('Shedule Update Success')
                print_shedule(shedule)
            except:
                print('Could not read shedule file email-reminder.xlsx!')

        
        SECFRSTART = (datetime.datetime.now() - TIMESTART).total_seconds()
        
        #try:
        
        send_check_notify(shedule)
        
        #print('Unexp exit') 
        
        #except:
        #    print('Err connecting')  
     

    print('End')
    


# In[4]:


def md5hash(fname):
    
    hash_md5 = hashlib.md5()
    
    try:
        f = open(fname, 'rb')    
        d = f.read()
        f.close()
        hash_md5.update(d)
        return hash_md5.hexdigest()
    except:
        print('md5err')
        return 'md5err'

def print_shedule(shedule):
    for key_i, val in shedule.items():
        #извлекаем
        email = val['email'] #email получателя
        back_email = val['back_email'] #email-ы  для оборатного ответа (список)
        shed = val['shedule'] #расписание приема АРМ ОДПГ
        name = val['name']
        sname = val['sname']
        
        print(name, sname, ' ', email)
        
        #пробегаем по расписанию   
        for key_j, reminder_list in shed.items():
            
            start = key_j[0]
            end = key_j[0]
            
            print(t2s(start), ' ', t2s(end))
    


def date_times_str_to_datetime(str_datetimes):
    
    datetime_list = str_datetimes.replace(' ', '').split(',')
    
    date = datetime.datetime.strptime(datetime_list[0], '%d.%m.%Y').date()
    time_start = datetime.datetime.strptime(datetime_list[1], '%H:%M').time()
    time_end = datetime.datetime.strptime(datetime_list[2], '%H:%M').time()
    
    return (datetime.datetime.combine(date, time_start), datetime.datetime.combine(date, time_end))

def reminder(shedule, reminder_rules): #формированеие расписания увдомлений
    for key, val in shedule.items():
        start = key[0]
        end = key[1]
        
        shedule[key] = []
        rem_time = 1
        
        for x in reminder_rules:
            if type(x) == str:
                rem_time = datetime.datetime.combine(start.date(), datetime.datetime.strptime(x, '%H:%M').time())
            else:
                rem_time = start - datetime.timedelta(hours=x)
            
            if [rem_time, 0, 0, 0, 0] not in shedule[key]: #если уже есть случайно
                shedule[key].append([rem_time, 0, 0, 0, 0])
                
        shedule[key].sort() #сортируем по возрастанию времени
               
   
        
def create_shedule(dfs, reminder_rules):
    shedule = dfs['name'].to_dict()
    
    for key, val in shedule.items():        
        
        shedule[key] = dfs.iloc[key][1:5].to_dict()
        
        back_email = shedule[key]['back_email']
        
        shedule[key]['back_email'] = back_email.replace(' ', '').split(',')
        
        date_time = dfs.iloc[key][5:]
        
        shed = {}
        
        for x in date_time:
            if x == 'none':
                #print('none')
                continue
             
            shed.update({date_times_str_to_datetime(x):[]})

        
        shedule[key].update({'shedule' : shed}) #график приема
        reminder(shedule[key]['shedule'], reminder_rules) 
        #теперь струкутура выглядти так (начало приема, конец):[время напоминания 1,0,0]
    
    return shedule


# In[22]:


def send_check_notify(d):
    
    global TIMESTART, SECFRSTART, LASTREM
    
    for key_i, val in d.items():
        #извлекаем
        email = val['email'] #email получателя
        back_email = val['back_email'] #email-ы  для оборатного ответа (список)
        shedule = val['shedule'] #расписание приема
        name = val['name']
        sname = val['sname']
        
        #пробегаем по расписанию   
        for key_j, reminder_list in shedule.items():
            
            start = key_j[0] #начало приема
                             #ключ словаря это tuple из двух значенией начало приема, конец приема
                             #reminder_list это список с датой временем напоминания и др.
            
            
            
            for x in reminder_list:
                
                #x[0]      x[1]       x[2]        x[3]
                rem_time, send_time, conf_time, adm_notify, adm_neg = x

                
                now = datetime.datetime.now()
                
                sec = (now - rem_time).total_seconds()
        
                if sec >= 0 and sec < 30 and send_time == 0: #начало уведомления за  30 сек. если ранее
                                                             #не посылали
                    
                    print('Sending..', name, email, 
                          'start:', t2s(start), 'remind:', t2s(rem_time), 'now:', t2s(now.time()))
                    
                    message = 'Напоминаем, что ' + t2s(start) + ' ' + name + ' ' + sname + ' осуществляет личный прием ССТУ. '+                              'Просьба в указанное время запустить АРМ. Если Вы планируете осуществить прием ССТУ '+                              'просьба ответить на это письмо словом "ПРИНЯТО" в теле письма, '+                              'в противном случае сообщить по телефону о невозможности проведениея приема в Отдел информатизации. '+                              'Если Вы сегодня уже отправляли подтверждение о приеме ССТУ, повторное подтверждение не требуется.'
                    
                    if isinstance(conf_time, datetime.datetime):
                        subj = 'ССТУ прием. Напоминание. ' + t2s(start) + ' прием осуществляет ' +                            name + ' ' + sname + ' ' 
                    else:
                        subj = 'ССТУ прием. ТРЕБУЕТСЯ ОТВЕТ. ' + t2s(start) + ' прием осуществляет ' +                            name + ' ' + sname + ' ' 

                    send_simple_email(email, subj, message)
                    
                    x[1] = datetime.datetime.now() #время отправления напоминания send_time
                    
                    #отправляем уведомление админу
                    
                    text_for_admin = 'ССТУ прием. ' + t2s(start) + ' ' + name + ' ' + sname
                    for m in back_email:
                        send_simple_email(m, text_for_admin, text_for_admin)

                   
                elif send_time !=0: #напоминание было отправлено
                    
                    now = datetime.datetime.now()
                    
                    before_confirm = (start - now).total_seconds()
                    after_last_notificatiob = (send_time - now).total_seconds()
                    
                    
                    last_reminder = reminder_list[-1][0]
                    first_reminder = reminder_list[0][0]
                    
                    if (before_confirm >= 0 and 
                        before_confirm <= 1800 and      #за полчаса до начала приема проверяем подтверждение
                        conf_time == 0):                #подтверждение получения не проверялось
                        
                        #время с прошлого напоминания
                        drem = (now - LASTREM).total_seconds()
                    
                        if drem > 60:   #10:  #120: #проверяем раз в две минуты     
                            
                            print('Check email', t2s(now))
                            
                            #получено подтверждение?
                            if check_confirm2(email, now.date(), first_reminder, start) :
                                                                
                                #получено подтверждение больше проверять не надо
                                mark_reminder_list_as_confirm(reminder_list, datetime.datetime.now())
                                
                                print('Conf received.!!')
                                
                                subj = 'ССТУ прием. Получено подтверждение. Прием ' + t2s(start) +                                       ' осуществляет ' + name + ' ' + sname
                                body = 'ССТУ прием. Получено подтверждение. Прием ' + t2s(start) +                                       ' осуществляет ' + name + ' ' + sname
                                
                                #отправляем уведомление админу
                                for m in back_email:
                                    send_simple_email(m, subj, body)
                                
                                
                            elif before_confirm < 900 and conf_time != -1: #900:   #прекращаем напоминания за 15 мин
                                
                                mark_reminder_list_as_non_confirm(reminder_list)
                                
                                print('None Conf received.!!')
                                
                                subj = 'ССТУ прием. Подтверждение НЕ ПОЛУЧЕНО. Прием ' + t2s(start) +                                       ' осуществляет ' + name + ' ' + sname
                                body = 'ССТУ прием. Подтверждение НЕ ПОЛУЧЕНО. Прием ' + t2s(start) +                                       ' осуществляет ' + name + ' ' + sname
                                
                                #отправляем уведомление админу
                                for m in back_email:
                                    send_simple_email(m, subj, body)
                                
                          
                            
                            LASTREM = datetime.datetime.now()
                                
  

def check_confirm2(email_from, now_date, first_reminder, start):
    
    tdelta = datetime.timedelta(days=1)
    
    msg_list = getimap_mail(now_date, now_date + tdelta, email_from)
    
    for x in msg_list:
        msg_datetime = x[0].replace(tzinfo=None)
        msg_body = x[4]
        
        pos = msg_body.upper().find('ПРИНЯТО')
        if pos >= 0 and msg_datetime < start and msg_datetime > first_reminder:
            return True
        
    return False





def mark_reminder_list_as_confirm(reminder_list, datetime_confirmation):
    for x in reminder_list:
        x[2] = datetime_confirmation
        

        
def mark_reminder_list_as_non_confirm(reminder_list):
    for x in reminder_list:
        x[2] = -1
    



def send_email(receiver_email, msg):
    smtp_server = email_server
    port = 587  # For starttls
    sender_email = 'mailrobot@eee.local.ru'   #служебный email
    
    context = ssl.SSLContext(ssl.PROTOCOL_SSLv23)
    
    with smtplib.SMTP(smtp_server, port) as server:        
        server.starttls(context=context)
        server.login('eee.local\mailrobot', 'password')
        server.sendmail(sender_email, receiver_email, msg.as_string()) 

        
def send_simple_email(receiver_email, subj, body):
    
    msg = MIMEText(body, 'plain', 'utf-8')
    msg['Subject'] = Header(subj, 'utf-8')
    msg['From'] = 'mailrobot@eee.local.ru'
    msg['To'] = receiver_email
    
    smtp_server = email_server
    port = 587  # For starttls
    sender_email = 'mailrobot@eee.local.ru'
    
    context = ssl.SSLContext(ssl.PROTOCOL_SSLv23)
    
    with smtplib.SMTP(smtp_server, port) as server:        
        server.starttls(context=context)
        server.login('eee.local\mailrobot', 'password')
        server.sendmail(sender_email, receiver_email, msg.as_string()) 

        
        
def t2str(t):
    r = 'datetime_err'
    if type(t) is datetime.date:
        r = datetime.datetime.strftime(t, '%d.%m.%Y')
    elif type(t) is datetime.datetime:
        r = datetime.datetime.strftime(t, '%d.%m.%Y %H:%M')
    elif type(t) is datetime.time:
        r = datetime.time.strftime(t, '%H:%M')
    
    return r      


def t2s(t):
    r = 'datetime_err'
    if type(t) is datetime.date:
        r = datetime.datetime.strftime(t, '%d.%m.%Y')
    elif type(t) is datetime.datetime:
        r = datetime.datetime.strftime(t, '%d.%m.%Y %H:%M')
    elif type(t) is datetime.time:
        r = datetime.time.strftime(t, '%H:%M')
    
    return r      


def str2t(t):
    r = 'datetime_err'
    
    try:
        r = datetime.datetime.strptime(t, '%d.%m.%Y %H:%M')
    except:
        try:
            r = datetime.datetime.strptime(t, '%d.%m.%Y').date()
        except:
            try:
                r = datetime.datetime.strptime(t, '%H:%M').time()
            except:
                pass
    
    return r    


def last_date(d):
    sh = []
    for key_i, val in d.items():
        shedule = val['shedule'] #расписание приема
        #пробегаем по расписанию   
        for key_j, reminder_list in shedule.items():
            start = key_j[0] #ключ словаря это tuple из двух значенией начало, конец приема
                             #reminder_list это список с датой временем напоминания
            sh.append(start)
    
    return max(sh)
    


# In[10]:


def getimap_mail(*args): #args:  since, before, email_from
    #соединение с сервером
    imap = imaplib.IMAP4_SSL(email_server)
    imap.login('local\mailrobot', 'password')
    #imap.list()
    imap.select('INBOX')
    
    
    if len(args) == 0:
        r, d = getimapids_all(imap)
    elif len(args) == 1:
        email_from = args[0]
        r, d = getimapids_all_from_email(imap, email_from)
    elif len(args) == 2:
        since = args[0]
        before = args[1]
        r, d = getimapids_fordates(imap, since, before)
    elif len(args) == 3:
        since = args[0]
        before = args[1]
        email_from = args[2]
        r, d = getimapids_fordates_email(imap, since, before, email_from)
    else:
        print('getimap_mail args errorr')
        imap.logout()
        return 0
    
    ids = d[0]
    id_list = ids.split() #идентификаторы сообщений
    
    msg_list = []

    for x in id_list:
        status, mail_data = imap.fetch(x, '(RFC822)')
        raw_email = mail_data[0][1]
        msg = email.message_from_bytes(raw_email, _class = email.message.EmailMessage)
        
        msg_date = email.utils.parsedate_to_datetime(getheader(msg['Date']))
        msg_from = email.utils.parseaddr(getheader(msg['From']))[1]
        msg_to   = email.utils.parseaddr(getheader(msg['To']))[1]
        msg_subj = getheader(msg['Subject'])
        msg_body = get_body(msg)
        
        msg_list.append([msg_date, msg_from, msg_to, msg_subj, msg_body])
        
        #print('Date:', email.utils.parsedate_to_datetime(getheader(msg['Date'])))
        #print(getheader(msg['Date']))
        #print('From1:', getheader(msg['From']))
        #print('From:', email.utils.parseaddr(getheader(msg['From']))[1])
        #print('To:', email.utils.parseaddr(getheader(msg['To']))[1])
        #print('Subj:', getheader(msg['Subject']))
        #print(get_charset(msg))
        #print('Body:\n')
        #print('Body:')
        #print(get_body(msg))
        #print('\n\n')
        
    imap.logout()
    return msg_list
                
def getheader(header_text, default="ascii"):
    """Decode the specified header"""
    headers = decode_header(header_text)
    header_sections = [(text.encode(charset or default)).decode() if type(text) == str else text.decode(charset or default)
                       for text, charset in headers]
    
    return ''.join(header_sections)


def get_charset(message, default="ascii"):
    """Get the message charset"""
    if message.get_content_charset():
        #print('1', message.get_content_charset())
        return message.get_content_charset()

    if message.get_charset():
        #print('2', message.get_charset())
        return message.get_charset()

    return default

def get_body(message):
    """Get the body of the email message"""
    if message.is_multipart():
        #get the plain text version only
        text_parts = [part
                      for part in typed_subpart_iterator(message,
                                                         'text',
                                                         'plain')]
        body = []
        for part in text_parts:
            charset = get_charset(part, get_charset(message))
            body.append(str(part.get_payload(decode=True),
                                charset,
                                "replace"))

        return '\n'.join(body).strip()

    else: # if it is not multipart, the payload will be a string
          # representing the message body
        body = str(message.get_payload(decode=True),
                       get_charset(message),
                       'replace')
        return body.strip()
    
    

#почта за период с даты(включая) до (исключая) imap работает только с датами (без часов и секунд)
def getimapids_fordates(imap, since, before):
    t1 = since.strftime("%d-%b-%Y") #из datetime в текстовой вид типа 25-Jun-2020
    t2 = before.strftime("%d-%b-%Y")
    date_str = '(since \"' + t1 + '\" before \"' + t2 + '\")'
      
    return imap.search(None, date_str)


def getimapids_all(imap):
    return imap.search(None, 'ALL')


def getimapids_fordates_email(imap, since, before, email_from):
    t1 = since.strftime("%d-%b-%Y") #из datetime в текстовой вид типа 25-Jun-2020
    t2 = before.strftime("%d-%b-%Y")
    
    search_str1 = '(since \"' + t1 + '\" before \"' + t2 + '\")'
    search_str2 = '(HEADER FROM ' + '"' + email_from + '")'
          
    return imap.search(None, search_str1, search_str2)


def getimapids_today(imap):
    #now = datetime.datetime.now()
    now = datetime.datetime.today()
    d = datetime.timedelta(days=1)
    
    return getimap_ids_fordate(imap, now, now + d) 


def getimapids_all_from_email(imap, email_from):
    search_str1 = '(HEADER FROM ' + '"' + email_from + '")'
    return imap.search(None, search_str1)

def delete_imapids(imap, ids_list):
    for x in ids_list:
        #print(x)
        imap.store(x, '+FLAGS', '\\Deleted') #удаление писем
    imap.expunge()
    
def delete_email(*args):
    #соединение с сервером
    imap = imaplib.IMAP4_SSL(email_server)
    imap.login('local\mailrobot', 'password')
    #imap.list()
    imap.select('INBOX')
    
    if len(args) == 0:
        r, d = getimapids_all(imap)
    elif len(args) == 1:
        email_from = args[0]
        r, d = getimapids_all_from_email(imap, email_from)
    elif len(args) == 2:
        since = args[0]
        before = args[1]
        r, d = getimapids_fordates(imap, since, before)
    elif len(args) == 3:
        since = args[0]
        before = args[1]
        email_from = args[2]
        r, d = getimapids_fordates_email(imap, since, before, email_from)
    else:
        print('getimap_mail args errorr')
        imap.logout()
        return 0
    
    ids = d[0]
    ids_list = ids.split()
    
    delete_imapids(imap, ids_list)
    
    imap.logout()

delete_email('email5@eee.local.ru')
    
main()
