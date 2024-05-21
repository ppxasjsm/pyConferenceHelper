from appscript import app, k
from mactypes import Alias
from pathlib import Path
import argparse
import os
import pandas as pd
import numpy as np
from random import randint
from time import sleep

__author__ = "Antonia Mey"
__email__ = "antonia.mey@ed.ac.uk"
__copyright__ = "BSD" 


def create_message_with_attachment(subject='', body='', to_recip=[],cc_recip=[]):

    msg = Message(subject=subject, body=body, to_recip=to_recip, cc_recip=cc_recip)

    # attach file
    # p = Path('path/to/myfile.pdf')
    # msg.add_attachment(p)

    #msg.show()
    msg.send()

def create_body(person, message, csv_start = 0 ):
    r"""
    Assembles the message from a single file
    Parameters
    ----------
    person : String Array
        Contains the title, first name, second name and email of the receipient
    message : Filename
        Filename of the file containing the message to be send minus the opening (Dear...)
    csv_start : integer
        Interger in the csv file from which the persons information contain title and name
    """
    title = person[0]
    #first_name = person[csv_start+1]
    second_name = person[1]
    body = 'Dear '+title+' '+second_name+',\n\n'
    f = open(message, 'r')
    body+=f.read()
    f.close()
    #body = MIMEText(body, 'plain')
    return body

class Outlook(object):
    def __init__(self):
    	self.client = app('Microsoft Outlook')

class Message(object):
    def __init__(self, parent=None, subject='', body='', to_recip=[], cc_recip=[], show_=True):

        if parent is None: parent = Outlook()
        client = parent.client

        self.msg = client.make(
            new=k.outgoing_message,
            with_properties={k.subject: subject, k.content: body})

        self.add_recipients(emails=to_recip, type_='to')
        self.add_recipients(emails=cc_recip, type_='cc')

        #if show_: self.show()

    def show(self):
        self.msg.open()
        self.msg.activate()

    def send(self):
        #self.msg.open()
        #self.msg.activate()
        self.msg.send()

    def add_attachment(self, p):
        # p is a Path() obj, could also pass string

        p = Alias(str(p)) # convert string/path obj to POSIX/mactypes path

        attach = self.msg.make(new=k.attachment, with_properties={k.file: p})

    def add_recipients(self, emails, type_='to'):
        if not isinstance(emails, list): emails = [emails]
        for email in emails:
            self.add_recipient(email=email, type_=type_)

    def add_recipient(self, email, type_='to'):
        msg = self.msg

        if type_ == 'to':
            recipient = k.to_recipient
        elif type_ == 'cc':
            recipient = k.cc_recipient

        msg.make(new=recipient, with_properties={k.email_address: {k.address: email}})

if __name__ == '__main__':
    parser = argparse.ArgumentParser()

    parser.add_argument(
            "-csv_file",
            help='File containing names and email addresses',
            default='test',
            metavar='String'
        )

    parser.add_argument(
            "-attachments",
            help="File list of attachments",
            default = [],
            nargs='+',
            metavar='String'
        )
    parser.add_argument(
            "-message",
            help = "File containing the message that needs to be sent out multiple times",
            default = 'test.txt',
            metavar = 'String'
        )
    parser.add_argument(
            "-cc_mail",
            help = "cc_email address",
            default = '',
            metavar = 'String'
        )
            
    args = parser.parse_args()
    
    # checking if directories for writing exist

    csv_file = args.csv_file
    attachments = args.attachments
    message = args.message
    cc_email = args.cc_mail
    
    print('==============================================')
    print("Sending Emails with the folloing information: ")
    print('==============================================')
    print('[csv file: ]    '+csv_file)
    print('[attachments: ]         %s'%attachments)
    print('[message file: ] '+message) 
    print('[Email from: ]'+cc_email)
    print('==============================================')


    subject = 'Recent Appointees in Physical Chemistry meeting 2-4 September 2024 Edinburgh'

    df = pd.read_csv(csv_file,sep=',', header = None, skiprows=1)
    people = np.array(df)
    num_entries = people.shape[0]

    for i in range(num_entries):
        body = create_body(people[i], message)
        print(body)
        to_email = people[i][2]
        print(to_email)
        print('-----------------sending email----------------------')
        create_message_with_attachment(subject=subject, body=body, to_recip=[to_email], cc_recip=[cc_email])
        sleep(randint(1,5))
