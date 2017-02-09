#!/usr/bin/python
# -*- coding: cp1250  -*-
__author__ = 'Fekete András Demeter'

import win32com.client
import time
import os
from datetime import datetime
import subprocess

class OutlookLib:

    def __init__(self, settings={}):
        self.settings = settings

    def get_messages(self):

        global another_folder
        global another_folder_4
        global another_folder_5
        global another_folder_6
        global another_folder_11
        global auto_arch_folder

        outlook = win32com.client.Dispatch("Outlook.Application")
        ns = outlook.GetNamespace("MAPI")

        another_folder =  ns.GetDefaultFolder(6).Folders.Item('REPORTS')
        another_folder_4 =  ns.GetDefaultFolder(6).Folders.Item('__TASK_BOSS')
        another_folder_5 =  ns.GetDefaultFolder(6).Folders.Item('__TASK_OPERATION')
        another_folder_11 =  ns.GetDefaultFolder(6).Folders.Item('__ANOTHER_MAILS')

        all_inbox = ns.GetDefaultFolder(6).Items
        # all_inbox=all_inbox.Restrict("[Unread] = true")  if you need only unread message!

        auto_arch_folder =  ns.GetDefaultFolder(6).Folders.Session.Folders.Item("Auto_riports_arch").Folders.Item('riports')  # Archiv pst filename » Auto_riports_arch, subfoder » riports

        return all_inbox


    def get_body(self, msg):
        return msg.Body

    def get_subject(self, msg):
        return msg.Subject

    def get_sender(self, msg):
        return msg.SenderName

    def get_cc(self, msg):
        return msg.CC

    def get_bcc(self, msg):
        return msg.BCC

    def get_recipient(self, msg):
        return msg.To

    def get_attachments(self, msg):
        return msg.Attachments

    def get_IsMarkedAsTask(self,msg):
        return msg.IsMarkedAsTask


def mail_check():

    outlook = OutlookLib()
    messages = outlook.get_messages()
    msg = messages.GetFirst()

    while msg:

        if type(msg) == 'NoneType':
            time.sleep(30)
        else:

            if msg.Subject == "" and msg.Body == "":
                msg = messages.GetNext()

            else:

                text_subj = str(msg.Subject)
                text_from = str(msg.SenderName)
                print(text_from, ' ', text_subj)

                if msg.SenderName == "REPORT_SENDER":
                    if not len(msg.Attachments) is 0:

                        attach.append((msg.Attachments, msg.Subject))
                        for attachTuple in attach:

                            for fileAtt in attachTuple[0]:
                                try:

                                    fileAtt.SaveAsFile(r"fileservicepath" + fileAtt.FileName)
                                    pass

                                except:

                                    time.sleep(2)
                                    try:
                                        fileAtt.SaveAsFile(r"fileservicepath" + fileAtt.FileName)
                                        pass

                                    except:

                                        TIMESTR = time.strftime('%H:%M:%S')
                                        date_time = str(datetime.date.today())
                                        print("I can't save: ", fileAtt.FileName, ' ', date_time, ' ', TIMESTR)

                        msg.Move(another_folder)

                    else:

                        try:

                            u = msg.SenderName
                            if u == 'XY2':
                                msg.Move(another_folder_4)
                            elif u == 'XY3':
                                msg.Move(another_folder_5)
                            else:
                                msg.Move(another_folder_11)


                        except:
                            pass

        time.sleep(2)
        msg = messages.GetNext()


def connections_test():
    i = 1
    while i < 100:

        try:
            fileservice_conn_test_1 = os.path.isdir("fileservicepath")
        except:
            fileservice_conn_test = 'False'

        if fileservice_conn_test_1 == True:
            fileservice_conn_test = 'True'
        else:
            fileservice_conn_test = 'False'

        if fileservice_conn_test == 'True':
            timex = time.strftime('%H:%M:%S')
            print('- >> Connection test success << - ', timex)
            break
        else:
            timex = time.strftime('%H:%M:%S')
            print('- >>  Connection test false, wait(15) and test again.. << - ', timex)
            time.sleep(15)
            i += 1

def archive_mail_from_report():

        try:

            lowerAccount = another_folder.Items
            msg2 = lowerAccount.GetFirst()
            f = len(lowerAccount)
            while f != 0:

                lowerAccount = another_folder.Items
                msg2 = lowerAccount.GetFirst()

                while msg2:
                    msg2.Move(auto_arch_folder)
                    msg2 = lowerAccount.GetNext()
                f = len(lowerAccount)


        except Exception as e:

            timex = time.strftime('%H:%M:%S')
            print('ERROR - >> ' + str(e) + ' << - ', timex)

            try:
                os.system('taskkill /f /im OUTLOOK.exe')
            except Exception as e:
                print("error ... ", e)

            time.sleep(600)

            pii = subprocess.Popen(r'start "C:\Program Files\Microsoft Office\Office12" OUTLOOK',
                                   shell=True,
                                   stdin=subprocess.PIPE,
                                   stdout=subprocess.PIPE,
                                   stderr=subprocess.STDOUT, )

            pii_stdout, pii_err = pii.communicate()
            p_status = pii.wait()

            if p_status == 0:
                print('reboot successful')
            else:
                print('reboot error')

            connections_test()


if __name__ == "__main__":

    while 1:

        attach = []

        try:

            i=0
            while i<10:
                try:
                    mail_check()
                    i=10
                except:
                    i+=1


        except Exception as e:

                timex = time.strftime('%H:%M:%S')
                print('ERROR - >> '+ str(e) +' << - ', timex)

                try:
                    os.system('taskkill /f /im OUTLOOK.exe')
                except Exception as e:
                    print( "error ... ", e)

                time.sleep(600)

                pii = subprocess.Popen(r'start "C:\Program Files\Microsoft Office\Office12" OUTLOOK',
                                                                                   shell=True,
                                                                                   stdin=subprocess.PIPE,
                                                                                   stdout=subprocess.PIPE,
                                                                                   stderr=subprocess.STDOUT,)
                pii_stdout, pii_err = pii.communicate()
                p_status = pii.wait()

                if p_status==0:
                    print('reboot successful')

                else:
                    print('reboot error')

                time.sleep(100)
                connections_test()

        finally:

            time_hour=int(time.strftime('%H'))

            if time_hour != time_hour2:

                time_hour2 = int(time.strftime('%H'))
                archive_mail_from_report()
