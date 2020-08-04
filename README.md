# o365-collect

A collection of scripts that will help to acquire data, such lists of contacts and s e-mail contents, from office365 tenants. These acquisition scripts are written in python and use the [python O365](https://github.com/O365/python-o365) library. To be able to use these scripts, access to the Office 365 is needed and therefore an App Registration in Azure Active Directory (AAD) must be setup. See the [documentation](docs/HOWTO.md) on how to do this.

## o365-acquire-mail.py

This script acquires e-mail contents from O365 environments in EML format.

In order to communicate with the Office 365 API, a `TENANT_ID`, a `CLIENT_ID`, and a `CLIENT_SECRET` must be supplied. See the [documentation](docs/HOWTO.md) to find out how to obtain these arguments.

For input, the script needs a list of Office 365 e-mail addresses that will be queried using the Office 365 API. Optionally, this list can be supplied using the `-f` option, but by default `stdin` is used to read e-mail addresses

The script needs an `OUTPUT_DIR` (supplied by the `-o` option) which is used to store e-mail contents. If this directory does not exists, it will be created automatically.

The script is able to extract attachments from e-mail contents automatically using the `e` option. When you need to search extracted attachments for suspicious macro's, see the `extract-macros.py` script in the [office-forensics repository](https://gitlab.local.northwave.nl/mboekelo/office-forensics).

### Usage

```text
o365-acquire-mail.py [-h] [-c CONFIG] -t TENANT_ID -i CLIENT_ID -s
                            CLIENT_SECRET [-d] -o OUTPUT_DIR [-f FILE] [-e]
                            [--concurrent-messages {1,2,3,4,5,6,7,8,9,10}]
                            [--concurrent-mailboxes {1,2,3,4,5}]
                            [--date-from DATE_FROM]
                            [--date-to DATE_TO]
                            [--has-attachment]
                            [--has-attachment-types [FILE_EXTENSIONS]]
                            [--contains-subject SUBJECT]
```

### Arguments

```text
  -h, --help            show this help message and exit
  -c CONFIG, --config CONFIG
                        Specify config file (.ini or .yaml), which includes
                        arguments
  -t TENANT_ID, --tenant-id TENANT_ID
                        Tenant ID of the Microsoft O365 environment (e.g.
                        815099a1-e306-4dd8-a3c1-1e05d436698a)
  -i CLIENT_ID, --client-id CLIENT_ID
                        Client ID of the App Registration created in AAD in
                        Azure Portal (e.g.
                        9a181509-06e3-dd84-1a3c-8a1e05d43669)
  -s CLIENT_SECRET, --client-secret CLIENT_SECRET
                        Client Secret of the App Registration created in AAD
                        in Azure Portal (e.g. oy:V1IxA5YXyj@Ry3E2ExhJBR:?=77@)
  -d, --delegated       Authorize user(s) interactively (delegated). This is
                        not needed when the App Registration has 'application'
                        privileges
  -o OUTPUT_DIR, --output-dir OUTPUT_DIR
                        Output directory where acquired e-mail contents will
                        be stored
  -f FILE, --file FILE  File with e-mail addresses (per line). (default=stdin)
  -e, --extract-attachments
                        Extract attachments into subfolders of specified
                        output directory (default=False)
  --concurrent-messages {1,2,3,4,5,6,7,8,9,10}
                        Number of concurrent messages to be parsed/fetched.
                        Messages are processed multi-threaded, use with care
                        (default=5)
  --concurrent-mailboxes {1,2,3,4,5}
                        Number of concurrent mailboxes to be parsed/fetched.
                        Mailboxes are processed multi-threaded, use with care
                        (default=1)
  --date-from DATE_FROM
                        Filter for messages that were received on or after
                        this date
  --date-to DATE_TO     Filter for messages that were received on or before
                        this date
  --has-attachment      Filter for messages that have one or more non-inline
                        attachments
  --has-attachment-types [FILE_EXTENSIONS]
                        Filter for messages that have one or more non-inline
                        attachments, but only attachments that have one of the
                        specified FILE_EXTENSIONS, separated by ',' (default
                        file extensions: 7z,bat,cmd,doc,docm,dotm,exe,gadget,h
                        ta,inf,jar,js,jse,lnk,msi,msp,potm,ppsm,ppt,pptm,ps1,p
                        y,rar,sfc,sfx,sldm,vb,vbe,vbs,xla,xlam,xltm,zip). This
                        option enforces --has-attachment
  --contains-subject SUBJECT
                        Filter for messages that include this subject
```

### Example

The following example extracts e-mail contents dating from January 2019 and have at least one attachment. E-mail contents are extracted into the folder `output` and attachments will be extracted separately. No user interaction is required. E-mail addresses are supplied in a file called `o365-email-addresses.txt`.

```bash
o365-acquire-mail.py \
-f o365-email-addresses.txt
-t '9a181509-06e3-dd84-1a3c-8a1e05d43669' \
-i '815099a1-e306-4dd8-a3c1-6698a1e05d43' \
-s 'b1Y=HBnusr]zvLIVww7tmGQj]5fq0-aC' \
--has-attachment --date-from 2019-01-01 --date-to 2019-02-01 \
-o output -e
```

### Config

To simplify usage, environmental arguments (such as tenant-id, client-id and client-secret) may be stored within a config file and passed as a single argument (--config):

```text
Example config:
tenant-id : 9a181509-06e3-dd84-1a3c-8a1e05d43669
client-id : 815099a1-e306-4dd8-a3c1-6698a1e05d43
client-secret: b1Y=HBnusr]zvLIVww7tmGQj]5fq0-aC
```

## o365-acquire-contacts.py

This script acquires all e-mail addresses registered in an O365 environment and prints these.

### Usage

```text
o365-acquire-contacts.py [-h] [-c CONFIG] -t TENANT_ID -i CLIENT_ID -s
                                CLIENT_SECRET [-d]
```

### Arguments

```text
  -c CONFIG, --config CONFIG
                        Specify config file (.ini or .yaml), which includes
                        arguments
  -t TENANT_ID, --tenant-id TENANT_ID
                        Tenant ID of the Microsoft O365 environment (e.g.
                        815099a1-e306-4dd8-a3c1-1e05d436698a)
  -i CLIENT_ID, --client-id CLIENT_ID
                        Client ID of the App Registration created in AAD in
                        Azure Portal (e.g.
                        9a181509-06e3-dd84-1a3c-8a1e05d43669)
  -s CLIENT_SECRET, --client-secret CLIENT_SECRET
                        Client Secret of the App Registration created in AAD
                        in Azure Portal (e.g. oy:V1IxA5YXyj@Ry3E2ExhJBR:?=77@)
  -d, --delegated       Authorize user(s) interactively (delegated). This is
                        not needed when the App Registration has 'application'
                        privileges
```

### Example

The following example extracts Office 365 e-mail addresses using delegated authorization (so user interaction is required):

```bash
o365-extract-contacts.py \
-t '9a181509-06e3-dd84-1a3c-8a1e05d43669' \
-i '815099a1-e306-4dd8-a3c1-6698a1e05d43' \
-s 'b1Y=HBnusr]zvLIVww7tmGQj]5fq0-aC' \
-d
```

### Config

Similar to `o365-acquire-mail.py`, a config file can be passed, including one or more arguments
