#coding=utf-8

import os
import re
import sys
import csv
import glob
import time
import string
import logging
import requests
import argparse
import codecs
from lxml import html
from datetime import datetime, timedelta

from os import mkdir
from os.path import isfile

last_date_of_file = None

class Crawler():
    def __init__(self, prefix_twse_idx="TWSE_IDX"):
        self.prefix_twse_idx = prefix_twse_idx
        ''' Check the original TWSE_IDX file is exist '''
        if not isfile('{}.csv'.format(self.prefix_twse_idx)):
            print 'File {}.csv lost, please check!!!\r'.format(self.prefix_twse_idx)
            quit()

    def process(self, date, total, count):
        print "Crawling @ " + "%15s " % date,
        print "[%3d"%(count*100/total) + "%]\r",

    def _clean_row(self, row):
        ''' Clean comma and spaces '''
        for index, content in enumerate(row):
            row[index] = re.sub(",", "", content.strip())
            row[index] = filter(lambda x: x in string.printable, row[index])
        return row

    def _twse_idx_record(self, row):
        ''' Save row to csv file '''
        if (os.path.isfile('{}.csv'.format(self.prefix_twse_idx))):
            f = open('{}.csv'.format(self.prefix_twse_idx), 'ab')
            cw = csv.writer(f, lineterminator='\r\n')
            cw.writerow(row)
            f.close()

    def _get_twse_idx_data(self, date_str): 
        global last_date_of_file

        date_str_w = date_str.replace("/","")   #strip '/' in date_str
        date_str_c = str((int(date_str[0:4])-1911))+date_str[4:]    #transfer to Taiwan year
        
        # fetch TWSE Volume and Amount
        url_vol = 'http://www.twse.com.tw/exchangeReport/FMTQIK?response=html&date={}'.format(date_str_w)
        page_vol = requests.get(url_vol)

        if not page_vol.ok:
            logging.error("Can not get TWSE index data at {}".format(date_str_c))
            return

        # fetch TWSE Index
        url_idx = 'http://www.twse.com.tw/indicesReport/MI_5MINS_HIST?response=html&date={}'.format(date_str_w)
        page_idx = requests.get(url_idx)

        if not page_idx.ok:
            logging.error("Can not get TWSE index data at {}".format(date_str_c))
            return

        # Parse page
        tree_vol = html.fromstring(page_vol.text)
        tree_idx = html.fromstring(page_idx.text)

        for (tr_idx, tr_vol) in zip(tree_idx.xpath('//table[1]/tbody/tr'),tree_vol.xpath('//table[1]/tbody/tr')):
            tds_idx = tr_idx.xpath('td/text()')
            tds_vol = tr_vol.xpath('td/text()')

            if tds_idx[0] == tds_vol[0]:    #make sure volume and index with the same day 
                date_fetch = tds_idx[0]
                date_record = '{}/{}/{}'.format(str(int(date_fetch[:3])+1911),date_fetch[4:6],date_fetch[7:9])
                    
                tds_vol[1] = tds_vol[1].replace(",","")   #strip ',' in Volume
                tds_vol[2] = tds_vol[2].replace(",","")   #strip ',' in Amount

                row = self._clean_row([
                    date_record, # 日期
                    tds_idx[1],  # Open index
                    tds_idx[2],  # High index
                    tds_idx[3],  # Low index
                    tds_idx[4],  # Close index
                    str(int(tds_vol[2])/100000000), # 成交金額 (億)
                    str(int(tds_vol[1])/1000), # 成交股數 (千股/張)
                    tds_vol[3], # 成交筆數
                ])

                # check date to prevent overwrite at the same date
                date_check_old = time.strptime(last_date_of_file, "%Y/%m/%d")
                date_check_new = time.strptime(row[0], "%Y/%m/%d")

                if date_check_old < date_check_new:
                    self._twse_idx_record(row)
                     
    def get_data(self, year, month, day):
        date_str_c = '{0}/{1:02d}/{2:02d}'.format(year - 1911, month, day)
        date_str_w = '{0}/{1:02d}/{2:02d}'.format(year,month,day)

        self._get_twse_idx_data(date_str_w)


    def get_last_date(self):
        with open('./{}.csv'.format(self.prefix_twse_idx), 'rb') as file:
            reader = csv.reader(file)
            rows = 0
            for row in reader:
                rows += 1
        file.close()
        return row[0]

def main():

    global last_date_of_file
    global total_days_to_crawl
    
    # Set logging
    if not os.path.isdir('log'):
        os.makedirs('log')
    logging.basicConfig(filename='log/crawl-error.log',
        level=logging.ERROR,
        format='%(asctime)s\t[%(levelname)s]\t%(message)s',
        datefmt='%Y/%m/%d %H:%M:%S')

    # Get arguments
    parser = argparse.ArgumentParser(description='Crawl TWSE index and volume data')
    args = parser.parse_args()

    crawler = Crawler()

    # Crawl data
    # TWSE-Volume first day is 1990/01
    # TWSE-Index first day is 1999/01
    print 'Crawling TWSE index and volume'

    # update TWSE index data till today 

    first_day = datetime.today()
    # get the last data date from 'TWSE_IDX.csv'
    lastdate = crawler.get_last_date()
    last_date_of_file = lastdate
    #lastyear_c = lastdate.split('/')[0]
    #lastyear = int(lastyear_c) + 1911  
    lastyear = lastdate.split('/')[0]
    lastmonth = lastdate.split('/')[1]
    lastday = lastdate.split('/')[2]

    last_day = datetime(int(lastyear), int(lastmonth), int(lastday))
    first_day = datetime(int(1999), int(3), int(31))
    
    if last_day.strftime('%Y/%m/%d') == first_day.strftime('%Y/%m/%d'):
        print 'your data is up-to-date'
    else:
        print 'updating data from '+last_day.strftime('%Y/%m/%d')+' to '+first_day.strftime('%Y/%m/%d')+'...'
        max_error = 5
        error_times = 0
        month_counts = 1 
        last_day = last_day + timedelta(1)
        total_months_to_crawl = (first_day.year - last_day.year) * 12 + (first_day.month - last_day.month) + 1
        fetch_month = None 
        while error_times < max_error and last_day <= first_day:  #from past till now
            try:
                #crawler.get_data(last_day.year, last_day.month, last_day.day) #from past till now
                if fetch_month != last_day.month:
                    fetch_month = last_day.month
                    crawler.get_data(last_day.year, last_day.month, last_day.day) #from past till now
                    crawler.process(last_day.strftime('%B, %Y'), total_months_to_crawl, month_counts)
                    month_counts += 1
                    time.sleep(3) # crawling data by every 3sec to prevent server block 
                    error_times = 0
            except:
                date_str = first_day.strftime('%Y/%m/%d')
                logging.error('Crawl raise error {}'.format(date_str))
                error_times += 1
                continue
            finally:
                #first_day -= timedelta(1) #from now to past
                last_day += timedelta(1) #from past till now

        print ''
        print "done!"        

if __name__ == '__main__':
    main()
