#!/usr/bin/env python
import smtplib
import smtplib
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
import numpy as np
import pandas as pd
import time
from random import randint
import getpass
import argparse
import os

__author__ = "Antonia Mey"
__email__ = "antonia.mey@fu-berlin.de"
__copyright__ = "BSD" 

class MassEmail ( object ):
    
    def __init__( self , body = None):
        self.body = body
        
    def send_mail( self, fromEmail, toEmail, USERNAME, PASSWORD, smtpserver='smtp.gmail.com:587', files = None ):
        r"""
        Sends the actual email
        Parameters
        ----------
        fromEmail : String
            senders Email
        toEmail : String
            receivers Email
        USERNAME : String
            username for gmail account
        PASSWORD : String
            password for the gmail account
        smtpserver : string
            mailserver from which the emails are sent. Default gmail
        files : String List
            List of attchments
        """
        msg = MIMEMultipart('alternative')
        msg['Subject'] = 'AIMS-IMAGINARY workshop further information'
        msg['From'] = fromEmail
        msg['To'] = toEmail
        for filename in files:
            extension = os.path.splitext(filename)[1]
            if extension == 'pdf':
                f = file(filename)
                fp = open(filename, 'rb')
                pdfAttachment = MIMEApplication(fp.read(), 'pdf')
                fp.close()
            
                pdfAttachment.add_header('Content-Disposition', 'attachment', filename = filename)
                msg.attach(pdfAttachment)
            else:
                f = file(filename)
                attachment = MIMEText(f.read())
                attachment.add_header('Content-Disposition', 'attachment', filename=filename)           
                msg.attach(attachment)
        
        content = MIMEText(self.body, 'plain')
        msg.attach(content)
        server = smtplib.SMTP(smtpserver)
        server.starttls()
        server.login(USERNAME,PASSWORD)
        problems = server.sendmail(fromEmail, toEmail, msg.as_string())
        server.quit()
    
    def assemble_single_message( self, person, message, csv_start = 0 ):
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
        title = person[csv_start]
        first_name = person[csv_start+1]
        second_name = person[csv_start+2]
        self.body = 'Dear '+title+' '+first_name+' '+second_name+',\n\n'
        f = open(message, 'r')
        self.body+=f.read()
        f.close()

#================================
#Main
#================================
if '__main__' == __name__:
    # GET COMMAND LINE ARGUMENTS!
    parser = argparse.ArgumentParser()
    parser.add_argument(
            "-user",
            help='username for your gmail account',
            default='test',
            metavar='String'
        )
    parser.add_argument(
            "-csv_file",
            help='File containing names and email addresses',
            default='test',
            metavar='String'
        )
    parser.add_argument(
            "-csv_start",
            help='Column number statring from 0 with persons title',
            default = 0,
            type = int
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
            "-from_mail",
            help = " Email address form sender",
            default = '',
            metavar = 'String'
        )
            
    args = parser.parse_args()
    
    #checking if directories for writing exist
    user = args.user
    csv_file = args.csv_file
    from_mail = args.from_mail
    attachments = args.attachments
    csv_start = args. csv_start
    message = args.message
    
    print '=============================================='
    print "Sending Emails with the folloing information: "
    print '=============================================='
    print '[gmail user: ]  '+user
    print '[csv file: ]    '+csv_file
    print '[csv file start: ]        %d'%csv_start
    print '[attachments: ]         %s'%attachments
    print '[message file: ] '+message 
    print '[Email from: ]'+from_mail
    print '=============================================='
    promt = 'Password for gmail account for user '+user+' :'
    password = getpass.getpass(promt)
        
    
    df = pd.read_csv(csv_file,sep=',', header = None)
    people = np.array(df)
    num_entries = people.shape[0]
    
    for i in xrange(num_entries):
        print 'sending mail '+str(i+1) +'/'+str(num_entries)
        email = MassEmail()
        email.assemble_single_message(people[i], message)
        email.send_mail(from_mail, people[i][-1], user, password, files = attachments)
        time.sleep(randint(2,9))
