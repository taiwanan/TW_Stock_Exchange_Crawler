#!/bin/bash
echo "install pip first..."
easy_install.exe pip
echo "pip installed!!!"
echo "install python packages...."
pip install requests
pip install lxml
pip install xlsxwriter
pip install openpyxl
