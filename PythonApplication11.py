


import psutil
import time
import os
from email.message import EmailMessage
import ssl
import smtplib
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from email.mime.application import MIMEApplication
import servicemanager
import socket
import sys
import win32event
import win32service
import win32serviceutil
from time import sleep

class SystemWorkloadService(win32serviceutil.ServiceFramework):
    _svc_name_ = 'project1'
    _svc_display_name_ = 'work.os'

    def __init__(self, args):
        win32serviceutil.ServiceFramework.__init__(self, args)
        self.hWaitStop = win32event.CreateEvent(None, 0, 0, None)
        socket.setdefaulttimeout(60)
        self.is_alive = True

    def SvcStop(self):
        self.ReportServiceStatus(win32service.SERVICE_STOP_PENDING)
        win32event.SetEvent(self.hWaitStop)
        self.is_alive = False

    def SvcDoRun(self):
        servicemanager.LogMsg(servicemanager.EVENTLOG_INFORMATION_TYPE,
                              servicemanager.PYS_SERVICE_STARTED,
                              (self._svc_name_, ''))
        self.main()
def collect_os():
    cpu_percent = psutil.cpu_percent(interval=1)
    memory_info = psutil.virtual_memory()
    hdd_usage = psutil.disk_usage('/')
    network_info = psutil.net_io_counters()

    sent_mb = round(network_info.bytes_sent / (1024 * 1024), 2)
    recv_mb = round(network_info.bytes_recv / (1024 * 1024), 2)

    timestamp = time.strftime("%Y-%m-%d %H:%M:%S")
    data = f"Timestamp: {timestamp}\n"\
           f"CPU Usage: {cpu_percent}%\n"\
           f"Memory Usage: {memory_info.percent}%\n"\
           f"HDD Usage: {hdd_usage.percent}%\n"\
           f"Network Usage (bytes sent/received): {sent_mb}MB / {recv_mb}MB\n"

    return data

#print(collect_system_workload())
f = open("TextFile1.txt", "a")
f.write(collect_os())
f.close()


email_sender="the email sender......"
email_password='the password of the email sender'
email_receiver="the email receiver"
subject="Project 1 os"
body="please find the attached file"
em=MIMEMultipart()
em["From"]=email_sender
em["To"]=email_receiver
em["subject"]=subject
em.attach(MIMEText(body))
context=ssl.create_default_context()
while True:
    
    with open('TextFile1.txt', 'rb') as f:
        attachment = MIMEApplication(f.read(), _subtype='txt')
        attachment.add_header('Content-Disposition', 'attachment', filename='TextFile1.txt')
        em.attach(attachment)
    with smtplib.SMTP_SSL('smtp.gmail.com',465,context=context) as smtp:
        smtp.login(email_sender,email_password)
        smtp.sendmail(email_sender,email_receiver,em.as_string())
    time.sleep(3600)

if __name__ == '__main__':
    if len(sys.argv) == 1:
        servicemanager.Initialize()
        servicemanager.PrepareToHostSingle(SystemWorkloadService)
        servicemanager.StartServiceCtrlDispatcher()
    else:
        win32serviceutil.HandleCommandLine(SystemWorkloadService)