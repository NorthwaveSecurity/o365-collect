#!/usr/bin/env python3

import configargparse
import sys
import os
import re
import sys
import hashlib
import logging
import base64
import time
import threading

from requests.exceptions import HTTPError
from datetime import datetime
from O365 import Account

# Works on Python3 only
if int(sys.version_info[0]) != 3:
    sys.exit("This script must run using Python3 since 'O365' only supports Python3")

# Constants
LIMIT_MESSAGES = None
SUSPICIOUS_ATTACHMENT_FILETYPES='7z,bat,cmd,doc,docm,dotm,exe,gadget,hta,inf,jar,js,jse,lnk,msi,msp,potm,ppsm,ppt,pptm,ps1,py,rar,sfc,sfx,sldm,vb,vbe,vbs,xla,xlam,xltm,zip'

# Logging config
LOG = None

def createDirIfNotExists(directory):
    if not os.path.isdir(directory):
        try:
            os.mkdir(directory)
        except FileExistsError: # This could happen due to multi-threading
            return

def parseMessage(message, userEmail, settings):
    LOG.info("\tParsing message...")

    # Create hash of message
    messageHash = hashlib.md5('{}_{}_{}'.format(int(message.received.timestamp()), userEmail, message.sender.address).encode('utf-8')).hexdigest()
    attachmentsOfInterest = [a for a in message.attachments if not a.is_inline]
    if settings['has_attachment_types']:
        attachmentsOfInterest = [a for a in attachmentsOfInterest if re.search(settings['has_attachment_types'], a.name)]
    
    if (not settings['has_attachment']) or attachmentsOfInterest:
        createDirIfNotExists('{}'.format(settings['output_dir']))
        createDirIfNotExists('{}/{}'.format(settings['output_dir'], userEmail))
        message.save_as_eml(to_path='{}/{}/{}.eml'.format(settings['output_dir'], userEmail, messageHash))
        if settings['extract_attachments'] and attachmentsOfInterest:
            createDirIfNotExists('{}/{}/{}'.format(settings['output_dir'], userEmail, messageHash))
            for attachment in attachmentsOfInterest:
                try:
                    attachmentExtension = re.search('[.][^.]+$',attachment.name)
                    attachmentExtension = attachmentExtension.group(0) if attachmentExtension else ''
                    attachmentContent = base64.b64decode(attachment.content)
                    attachmentHash = hashlib.md5(attachmentContent).hexdigest()
                    with open('{}/{}/{}/{}{}'.format(settings['output_dir'], userEmail, messageHash, attachmentHash, attachmentExtension), 'wb') as attachmentFile:
                        attachmentFile.write(attachmentContent)
                except Exception as e:
                    LOG.warning("\tFailed to extract attachment ({})".format(e))


def parseInbox(userEmail, settings):
    if settings['concurrent_mailboxes'] > 1:
        # Delay so stdin can be pasted
        time.sleep(2)

    if settings['delegated']: # App registration has 'Delegated' privileges (mostly used for a few/single e-mail addresses)
        userAccount = Account(
            (settings['client_id'], settings['client_secret']),
            scopes=['basic'],
            auth_flow_type='authorization',
            tenant_id=settings['tenant_id'],
            main_resource=userEmail)
    else: # App registration has 'Application' privileges with no user interaction required
        userAccount = Account(
            (settings['client_id'], settings['client_secret']),
            auth_flow_type='credentials',
            tenant_id=settings['tenant_id'],
            main_resource=userEmail)


    LOG.info("[{}] Authenticating...".format(userEmail))
    if not userAccount.is_authenticated:
        userAccount.authenticate()

    LOG.info("[{}] Getting mailbox...".format(userEmail))
    userMailbox = userAccount.mailbox()

    # Build query
    # See: https://docs.microsoft.com/en-us/microsoft-365/compliance/keyword-queries-and-search-conditions
    # See: https://docs.microsoft.com/en-us/previous-versions/office/office-365-api/api/version-2.0/extended-properties-rest-operations
    query = userMailbox.new_query()
    if settings['date_from']:
        query = query.chain('and').on_attribute('ReceivedDateTime').greater_equal(settings['date_from'])
    if settings['date_to']:
        query = query.chain('and').on_attribute('ReceivedDateTime').less_equal(settings['date_to'])
    if settings['has_attachment']:
        query = query.chain('and').on_attribute('HasAttachments').equals(True)
    if settings['contains_subject']:
        query = query.chain('and').on_attribute('Subject').contains(settings['contains_subject'])

    LOG.info("[{}] Getting messages...".format(userEmail))
    userMessages = []
    try:
        userMessages = userMailbox.get_messages(limit=LIMIT_MESSAGES, query=query, download_attachments=True)
    except HTTPError as e:
        if(e.response.status_code == 404):
            userMessages = []  # This means no messages are found, which is assumed normal behavior
        else:
            LOG.error("[{}] {}".format(userEmail,e))

    LOG.info("[{}] Iterating messages...".format(userEmail))

    threads = [] 
    for message in userMessages:
        thread = threading.Thread(target=parseMessage, args=(message, userEmail, settings))
        thread.start()
        threads.append(thread)
        if len(threads) == settings['concurrent_messages']:
            for thread in threads:
                thread.join()
            threads = []
    for thread in threads:
        thread.join()

    LOG.info("[{}] Parsed all messages in this inbox...".format(userEmail))

def processMailboxes(settings):
    if settings['concurrent_mailboxes'] > 1:
        threads = []
        for userEmail in settings['file']:
            thread = threading.Thread(target=parseInbox, args=(userEmail, settings))
            thread.start()
            threads.append(thread)

            # Only 5 mailboxes at a time
            if len(threads) == settings['concurrent_mailboxes']:
                for thread in threads:
                    thread.join()
                threads = []
    else:
        for userEmail in settings['file']:
            parseInbox(userEmail.strip(), settings)

def _argGuid(arg):
    if re.match('[0-9a-fA-F]{8}-([0-9a-fA-F]{4}-){3}[0-9a-fA-F]{12}', arg):
        return arg
    else:
        raise configargparse.ArgumentTypeError('guid has to be in the form xxxxxxxx-xxxx-xxxx-xxxx-xxxxxxxxxxxx')

def main():
    argparser = configargparse.ArgumentParser(description="%s acquires e-mail messages from O365 environments and stores these messages as EML files in a specified output folder." % sys.argv[0])
    argparser.add_argument('-c','--config', is_config_file=True, help="Specify config file (.ini or .yaml), which includes arguments")
    argparser.add_argument('-t','--tenant-id', type=_argGuid, required=True, help="Tenant ID of the Microsoft O365 environment (e.g. 815099a1-e306-4dd8-a3c1-1e05d436698a)")
    argparser.add_argument('-i','--client-id', type=_argGuid, required=True, help="Client ID of the App Registration created in AAD in Azure Portal (e.g. 9a181509-06e3-dd84-1a3c-8a1e05d43669)")
    argparser.add_argument('-s','--client-secret', type=str, required=True, help="Client Secret of the App Registration created in AAD in Azure Portal (e.g. oy:V1IxA5YXyj@Ry3E2ExhJBR:?=77@)")
    argparser.add_argument('-d','--delegated', action='store_true', help="Authorize user(s) interactively (delegated). This is not needed when the App Registration has 'application' privileges")
    argparser.add_argument('-f','--file', type=configargparse.FileType('r'), default=sys.stdin, help="File with e-mail addresses (per line). (default=stdin)")
    argparser.add_argument('-o','--output-dir', type=str, required=True, help="Output directory where acquired e-mail contents will be stored")
    argparser.add_argument('-e','--extract-attachments', action='store_true', help="Extract attachments into subfolders of specified output directory (default=False)")
    argparser.add_argument('--concurrent-messages', type=int, default=5, choices=range(1,11), help="Number of concurrent messages to be parsed/fetched. Messages are processed multi-threaded, use with care (default=5)")
    argparser.add_argument('--concurrent-mailboxes', type=int, default=1, choices=range(1,6), help="Number of concurrent mailboxes to be parsed/fetched. Mailboxes are processed multi-threaded, use with care (default=1)")
    argparser.add_argument('--date-from', type=lambda s: datetime.strptime(s, '%Y-%m-%d'), help="Filter for messages that were received on or after this date")
    argparser.add_argument('--date-to', type=lambda s: datetime.strptime(s, '%Y-%m-%d'), help="Filter for messages that were received on or before this date")
    argparser.add_argument('--has-attachment', action='store_true', help="Filter for messages that have one or more non-inline attachments")
    argparser.add_argument('--has-attachment-types', nargs='?', metavar='FILE_EXTENSIONS', const=SUSPICIOUS_ATTACHMENT_FILETYPES, default=False, help="Filter for messages that have one or more non-inline attachments, but only attachments that have one of the specified FILE_EXTENSIONS, separated by ',' (default file extensions: {}). This option enforces --has-attachment".format(SUSPICIOUS_ATTACHMENT_FILETYPES))
    argparser.add_argument('--contains-subject', metavar='SUBJECT', help="Filter for messages that include this subject")
    settings = argparser.parse_args().__dict__

    # When set, convert --has-attachment-types to regex and enforce --has_attachment
    if settings['has_attachment_types']:
        settings['has_attachment_types'] = '[.]({})$'.format('|'.join(settings['has_attachment_types'].split(',')))
        settings['has_attachment'] = True
    
    #Set logging settings
    global LOG
    logging.basicConfig(format='%(levelname)-10s %(message)s', filename='o365-mail-acquisition-{}.log'.format(time.strftime('%Y%m%d%H%M%S')))
    logging.getLogger('O365.connection').setLevel(logging.CRITICAL)    
    LOG = logging.getLogger('O365-Extractor')
    LOG.setLevel(logging.INFO)
    LOG.addHandler(logging.StreamHandler())

    #logging.basicConfig()
    processMailboxes(settings)

if __name__ == '__main__':
    main()
