#!/usr/bin/env python3

import configargparse
import re
import sys
import logging
import time

from datetime import datetime
from O365 import Account

# Works on Python3 only
if int(sys.version_info[0]) != 3:
    sys.exit("This script must run using Python3 since 'O365' only supports Python3")

#LOG = None 

def fetchContactList(settings):
    if settings['delegated']:
        globalAccount = Account(
            (settings['client_id'], settings['client_secret']),
            scopes=['basic','users'],
            auth_flow_type='authorization',
            tenant_id=settings['tenant_id'])
    else:
        globalAccount = Account(
            (settings['client_id'], settings['client_secret']),
            auth_flow_type='credentials',
            tenant_id=settings['tenant_id'])

    #LOG.info("Authenticating...")
    if not globalAccount.is_authenticated:
        globalAccount.authenticate()
    
    #LOG.info("Retrieving contact list...")
    for userDetails in globalAccount.directory().get_users(limit=None):
        if userDetails.mail:
            print(userDetails.mail)
    
def _argGuid(arg):
    if re.match('[0-9a-fA-F]{8}-([0-9a-fA-F]{4}-){3}[0-9a-fA-F]{12}', arg):
        return arg
    else:
        raise configargparse.ArgumentTypeError('guid has to be in the form xxxxxxxx-xxxx-xxxx-xxxx-xxxxxxxxxxxx')

def main():
    argparser = configargparse.ArgumentParser(description="%s acquires a global list of e-mail addresses from O365 environments." % sys.argv[0])
    argparser.add_argument('-c','--config', is_config_file=True, help="Specify config file (.ini or .yaml), which includes arguments")
    argparser.add_argument('-t','--tenant-id', type=_argGuid, required=True, help="Tenant ID of the Microsoft O365 environment (e.g. 815099a1-e306-4dd8-a3c1-1e05d436698a)")
    argparser.add_argument('-i','--client-id', type=_argGuid, required=True, help="Client ID of the App Registration created in AAD in Azure Portal (e.g. 9a181509-06e3-dd84-1a3c-8a1e05d43669)")
    argparser.add_argument('-s','--client-secret', type=str, required=True, help="Client Secret of the App Registration created in AAD in Azure Portal (e.g. oy:V1IxA5YXyj@Ry3E2ExhJBR:?=77@)")
    argparser.add_argument('-d','--delegated', action='store_true', help="Authorize user(s) interactively (delegated). This is not needed when the App Registration has 'application' privileges")
    settings = argparser.parse_args().__dict__
    
    #Set logging settings
    # global LOG
    # logging.basicConfig(format='%(levelname)-10s %(message)s', filename='o365-contact-acquisition-{}.log'.format(time.strftime('%Y%m%d%H%M%S')))
    # logging.getLogger('O365.connection').setLevel(logging.CRITICAL)    
    # LOG = logging.getLogger('O365-Extractor')
    # LOG.setLevel(logging.INFO)
    # LOG.addHandler(logging.StreamHandler())

    fetchContactList(settings)


if __name__ == '__main__':
    main()
