#!/bin/bash
echo "install pip first..."
easy_install.exe pip
echo "pip installed!!!"
echo "install python packages...."
pip install requests
pip install lxml
pip install xlsxwriter
pip install openpyxl
pip install numpy
pip install testresources
pip install python-dateutil
echo "SSL: sudo apt-get install libffi-dev libssl-dev"
pip install pyopenssl ndg-httpsclient pyasn1
