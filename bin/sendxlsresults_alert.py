###############################################################################
###############################################################################
##
##  SENDXLSRESULTS - a splunk email command to send reports as XLS
##
##  Dominique Vocat, contributions by Felix Sterzelmaier <felix.sterzelmaier@concat.de>
##  Version 1.1.0 - respects default filename just like the send via email alert action if no name is specified.
##  Version 1.1.1 - python 2-3 readyness
##
###############################################################################
###############################################################################
from __future__ import print_function
import sys
import os
import json
import six.moves.urllib.request, six.moves.urllib.error, six.moves.urllib.parse
import csv
import gzip
import smtplib, email
from email.MIMEMultipart import MIMEMultipart
from email.mime.text import MIMEText
from email.MIMEBase import MIMEBase
from email import Encoders

import socket
import string
import random
import time
from collections import defaultdict
import logging as logger

import re
import xlwt
import copy
from splunk.util import normalizeBoolean

#might fix the error - see https://stackoverflow.com/questions/11536764/how-to-fix-attempted-relative-import-in-non-package-even-with-init-py
os.sys.path.append(os.path.dirname(os.path.abspath('.')))

## Create a unique identifier for this invocation
NOWTIME         = time.time()
SALT            = random.randint(0, 100000)
INVOCATION_ID   = str(NOWTIME) + ':' + str(SALT)
INVOCATION_TYPE = "action"

######################################################
######################################################
# Helper functions from a canonical splunk script plus
# own extensions.
#

def unquote(val):
    if val is not None and len(val) > 1 and val.startswith('"') and val.endswith('"'):
        return val[1:-1]
    return val

def toBool(strVal):
   if strVal == None:
       return False

   lStrVal = strVal.lower()
   if lStrVal == "true" or lStrVal == "t" or lStrVal == "1" or lStrVal == "yes" or lStrVal == "y" :
       return True
   return False

def intToBool(Val):
   if Val == 0:
       return False
   if Val == 1:
       return True
   
def getarg(argvals, name, defaultVal=None):
    return unquote(argvals.get(name, defaultVal))

def request(method, url, data, headers):
    # Helper function to fetch JSON data from the given URL
    req = six.moves.urllib.request.Request(url, data, headers)
    req.get_method = lambda: method
    res = six.moves.urllib.request.urlopen(req)
    return json.loads(res.read())

###############################################################################
#
# Function:   getEmailAlertActions
#
# Descrition: This function calls the Splunk REST API to get the various alert
#             email configuration settings needed to send SMTP messages in the
#             way that Splunk does
#
# Arguments:
#    argvals  - hash of various arguments passed into the search.
#    payload  - hash of various Splunk configuration settings.
#
###############################################################################

def getEmailAlertActions(argvals, payload):
    try:
        url_tmpl = '%(server_uri)s/services/configs/conf-alert_actions/email?output_mode=json'
        record_url = url_tmpl % dict(server_uri=payload.get('server_uri'))
        headers = {
            'Authorization': 'Splunk %s' % payload.get('session_key'),
            'Content-Type': 'application/json'}

        try:
            record = request('GET', record_url, None, headers)
        except six.moves.urllib.error.HTTPError as e:
            logger.error('invocation_id=%s invocation_type="%s" msg="Could not get email alert actions from splunk" error="%s"' % (INVOCATION_ID,INVOCATION_TYPE,str(e)))
            sys.exit(2)

        argvals['server'] = record['entry'][0]['content']['mailserver']
        argvals['sender'] = record['entry'][0]['content']['from']
        argvals['use_ssl'] = record['entry'][0]['content']['use_ssl']
        argvals['use_tls'] = record['entry'][0]['content']['use_tls']
        argvals['reportFileName'] = record['entry'][0]['content']['reportFileName']
    except six.moves.urllib.error.HTTPError as e:
        logger.error('invocation_id=%s invocation_type="%s" msg="Could not get email alert actions from splunk" error="%s"' % (INVOCATION_ID,INVOCATION_TYPE,str(e)))
        raise

###############################################################################
#
# Function:   sendemail
#
# Descrition: This function sends a MIME encoded e-mail message using Splunk SMTP
#              Settings.
#
# Arguments:
#    recipient - maps the the field 'email_to' in the event returned by Search.
#    subject - maps to the field 'subject' in the event returned by Search.
#    body - maps the field 'message' in the event returned by Search.
#    argvals - hash of various arguments needed to configure the SMTP connection etc.
#
###############################################################################

def sendemail(recipient, sender, subject, body, argvals, attachment):


    print("email sender: - recipient: "+recipient+" | subject: "+subject+" | body: "+body+" | attachment:"+attachment, file=sys.stderr)
    print(argvals, file=sys.stderr)
    server    = getarg(argvals, "server", "localhost")
    use_ssl   = intToBool(argvals['use_ssl'])
    use_tls   = intToBool(argvals['use_tls'])
    username  = getarg(argvals, "username"  , "")
    password  = getarg(argvals, "password"  , "")
    
    # make sure the sender is a valid email address
    if (sender.find("@") == -1):
        sender = sender + '@' + socket.gethostname()

    if sender.endswith("@"):
        sender = sender + 'localhost'

    # Create multi part mail
    msg = MIMEMultipart()
    msg['Subject'] = subject 
    msg['From'] = sender
    msg['To'] = recipient
    part1 = MIMEText(bodyText, 'plain', 'utf-8')
    msg.attach(part1)

    try:
        part2 = MIMEBase('application', "octet-stream")
        part2.set_payload(open(os.environ['SPLUNK_HOME'] + "/var/run/splunk/" + attachment, "rb").read())
        Encoders.encode_base64(part2)
        part2.add_header('Content-Disposition', 'attachment; filename="' + attachment + '"')
        msg.attach(part2)
    except Exception as e:
        print("exception while attaching file occured", file=sys.stderr)
        logger.error('invocation_id=%s invocation_type="%s" msg="error attaching file" rcpt="%s" error="%s"' % (INVOCATION_ID,INVOCATION_TYPE,recipient,str(e)))
        raise        
    try:
        # send the mail
        if not use_ssl:
            print("plain jane mail", file=sys.stderr)
            smtp = smtplib.SMTP(server)
        else:
            print("use ssl mail", file=sys.stderr)
            smtp = smtplib.SMTP_SSL(server)

        if use_tls:
            print("use TLS", file=sys.stderr)
            smtp.ehlo()
            smtp.starttls()
        if len(username) > 0 and len(password) >0:
            print("user username/password", file=sys.stderr)
            smtp.login(username, password)

        #print >> sys.stderr, msg.as_string()
        print("go go gadget - send email!", file=sys.stderr)
        smtp.sendmail(sender, string.split(recipient, ","), msg.as_string())
        smtp.quit()
        return
    
    except smtplib.SMTPRecipientsRefused as e:
        print("exception while sending email occured - refused recipients", file=sys.stderr)
        print(e.recipients, file=sys.stderr)
        print(e, file=sys.stderr)
        logger.error('invocation_id=%s invocation_type="%s" msg="Could not send email" rcpt="%s" error="%s"' % (INVOCATION_ID,INVOCATION_TYPE,recipient,str(e)))
        raise
    except (socket.error, smtplib.SMTPException) as e:
        print("exception while sending email occured", file=sys.stderr)
        print(e, file=sys.stderr)



######################################################
######################################################
#
# Main
#

if __name__ == "__main__":
    if len(sys.argv) > 1 and sys.argv[1] == "--execute" :

        logger.basicConfig(format='%(asctime)s %(levelname)s %(message)s', filename=os.path.join(os.environ['SPLUNK_HOME'],'var','log','splunk','sendxlsresults.log'), filemode='a+', level=logger.INFO)

        argvals        = defaultdict(list)
        recipient_list = defaultdict(list)
        event_list     = defaultdict(list)
        fields         = []
        header_key     = {}
        default_format = "table {font-family:Arial;font-size:12px;border: 1px solid black;padding:3px}th {background-color:#4F81BD;color:#fff;border-left: solid 1px #e9e9e9} td {border:solid 1px #e9e9e9}"

        payload = json.loads(sys.stdin.read())
        print("Payload:", file=sys.stderr)
        print(payload, file=sys.stderr)
        
        getEmailAlertActions(argvals,payload)

        settings = payload.get('configuration')
        print("Settings we received:", file=sys.stderr)
        print(settings, file=sys.stderr)

        print("Arguments we received:", file=sys.stderr)
        print(argvals, file=sys.stderr)
        bodyText           = getarg(settings, "body", "")
        subject            = getarg(settings, "subject", "")
        sender             = getarg(settings, "sender", "")
        recipient          = getarg(settings, "recipient", "")
        filename           = getarg(settings, "filename", "")
        search_name        = getarg(argvals, "search_name", "one")
        smptHost           = getarg(argvals, "server", "localhost")
        reportFileName     = getarg(argvals, "reportFileName", "")
        
        newFilename = search_name
        if filename!="":
            newFilename = filename

        filename = reportFileName.replace("$name$",newFilename)+".xls"
        regex = ur"\$time:([^$]+)\$"
        match = re.compile(regex).search(filename)
        if(match!=None):
            if(len(match.groups())==1):
                timestampString = match.group(1)
                timestampStringReplaced = time.strftime(timestampString, time.localtime(NOWTIME))
                filename = re.sub(regex, timestampStringReplaced, filename, 0)

        with gzip.open(payload.get('results_file'),'r') as fin:
            csvreader = csv.reader(fin, delimiter=',')

            #--
            output = os.environ['SPLUNK_HOME'] + "/var/run/splunk/" + filename
            logger.info("parameters used: outputfile %s" % (output))
            int_re = re.compile(r'^\d+$')
            float_re = re.compile(r'^\d+\.\d+$')
            date_re = re.compile(r'^\d+-\d+-\d+|^\d+\/\d+\/\d+|^\d+\.\d+\.\d+')
            style = xlwt.XFStyle()

            workbook = xlwt.Workbook(encoding="UTF-8") #()
            sheet = workbook.add_sheet(search_name)

            column_num = 0
            row_num = 0
            for row in csvreader:
                column_num = 0
                for item in row:
                    if not item.startswith('__'):
                        format = 'general'
                        cellvalue = ""
                        if re.match(date_re, item):
                            cellvalue = item
                            format = 'M/D/YY'
                        elif re.match(float_re, item):
                            cellvalue = float(item)
                            format = '0.00'
                        elif re.match(int_re, item):
                            cellvalue = float(item)
                            format = '0'
                        else:
                            cellvalue = item
                            format = ''
                        style.num_format_str = format
                        sheet.write(row_num, column_num, cellvalue, style) #sheet.write(row_num, column_num, unicode(item).encode("utf-8"), style)
                        column_num += 1
                row_num += 1
                #return True
                
            #save excel sheet
            workbook.save(output)
            
            try:
                sendemail(recipient, sender, subject, bodyText, argvals, filename)

            except Exception as e:
                import traceback
                stack =  traceback.format_exc()
                logger.error('invocation_id=%s invocation_type="%s" msg="some error occured - stack trace follows" %s' % (INVOCATION_ID,INVOCATION_TYPE, stack))
                exit()

    else:
        logger.error('invocation_id=%s invocation_type="%s" msg="Unsupported execution mode (expected --execute flag)"' % (INVOCATION_ID,INVOCATION_TYPE))
        sys.exit(1)
