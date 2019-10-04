from datetime import datetime as dt,timedelta
import time
import smtplib
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from email.mime.base import MIMEBase
from email.message import EmailMessage
from email import encoders
import schedule
from os import system
import gen_rep
import db
import os

def send_email():
    today = dt.today()
    days_to_subtract = 5
    date_to = dt.today() - timedelta(days=days_to_subtract)
    # Generating Report
    print('Generating Report!!!')
    gen_start = time.time()
    test = gen_rep.Report(db.get_results(),f'report_{today.strftime("%m-%d-%Y")}')
    test.generate_report()
    gen_end = time.time()
    gen_statement = f'\n{today} : Done Generating Report in {gen_end - gen_start}s'
    write_logs(gen_statement) 
    print(gen_statement)

    email_user = 'jdeus@fortmed.org'
    content = 'Good Day! \n\nPlease see attached file!'
    msg = MIMEMultipart()
    #to ='amendoza@fortmed.org,beth.ko@fortmed.org,ageronimo@fortmed.org,eugenemacalalag@gmail.com'
    to = 'jdeus@fortmed.org,dtadeo@fortmed.org'
    msg['Subject'] =f'e-Result Report from {date_to.strftime("%m/%d/%Y")} to {today.strftime("%m/%d/%Y")}'
    msg['From'] = email_user
    msg['To'] = to

    msg.attach(MIMEText(content,'plain'))
    file_name = f'reports/report_{today.strftime("%m-%d-%Y")}.xlsx'
    attach_file = open(file_name,'rb')

    payload = MIMEBase('application','octet-stream')
    payload.set_payload((attach_file).read())
    encoders.encode_base64(payload)

    payload.add_header('Content-Disposition','attachment',filename=file_name)
    msg.attach(payload)
    text = msg.as_string()
    smtp = smtplib.SMTP('just20.justhost.com',587)
    smtp.starttls()
    smtp.login(email_user,'L1f3feelsg00d@')
    smtp.sendmail(email_user,to.split(','),text)
    smtp.quit
    statement = f"\n{today} : Sending report to  {to.split(',')}"
    write_logs(statement)
        
    
def write_logs(statement):
    file = open('logs.txt','a+')
    file.write(statement)
    file.close()
        
def job():
    today = dt.today()
    print('Sending an email!!')
    send_start = time.time()
    send_email()
    text = f'\n{today} : Email Successfuly Sent in { time.time() - send_start}s'
    write_logs(text)
    print(text)
    time.sleep(5)
    screen()

def date_checking():
    days = dt.today().weekday()
    text = ''
    if (days== 5):
        text = '\nGood morning! Im sending email at exactly 7:00PM in the evening'
    else:
        text = f'\nGood morning! {6-days} day(s) before I send an email to the managers'
    print(text)
    time.sleep(5)
    screen()

def screen():
    system('cls')
    print('Automated Email for E-Result Reports is now running!')


#Schedules
# Date Checking
schedule.every(1).minutes.do(job)
schedule.every().day.at('09:00').do(date_checking)
# Schedule for emailing the managers
schedule.every().saturday.at("19:00").do(job)
screen()
while True:
    schedule.run_pending()
    time.sleep(1)
