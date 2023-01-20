#!c:\users\rocketbot\appdata\local\programs\python\python36-32\python.exe
# EASY-INSTALL-ENTRY-SCRIPT: 'xml2xlsx==1.0.1','console_scripts','xml2xlsx'
__requires__ = 'xml2xlsx==1.0.1'
import re
import sys
from pkg_resources import load_entry_point

if __name__ == '__main__':
    sys.argv[0] = re.sub(r'(-script\.pyw?|\.exe)?$', '', sys.argv[0])
    sys.exit(
        load_entry_point('xml2xlsx==1.0.1', 'console_scripts', 'xml2xlsx')()
    )
