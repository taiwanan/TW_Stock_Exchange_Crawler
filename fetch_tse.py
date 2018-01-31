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
from os.path import isdir

class Crawler():
    def __init__(self, prefix_tse="tse_raw_data", prefix_otc="otc_raw_data"):
        ''' Make directory if not exist when initialize '''
        if not isdir(prefix_tse):
            mkdir(prefix_tse)
        self.prefix_tse = prefix_tse
        if not isdir(prefix_otc):
            mkdir(prefix_otc)
        self.prefix_otc = prefix_otc
        # total files in both TSE/OTC
        tse_files = sum(os.path.isfile(os.path.join('{}/'.format(self.prefix_tse), f)) 
                        for f in os.listdir('{}/'.format(self.prefix_tse))) 
        otc_files = sum(os.path.isfile(os.path.join('{}/'.format(self.prefix_otc), f)) 
                        for f in os.listdir('{}/'.format(self.prefix_otc))) 
        self.tse_files = tse_files
        self.otc_files = otc_files

    def process(self, date, total, count):
        print date + "...",
        print "%3d"%(count*100/total) + "%\r",

    def _clean_row(self, row):
        ''' Clean comma and spaces '''
        for index, content in enumerate(row):
            row[index] = re.sub(",", "", content.strip())
            row[index] = filter(lambda x: x in string.printable, row[index])
        return row

    def _tse_record(self, stock_id, row):
        ''' Save row to csv file '''
        f = open('{}/{}.csv'.format(self.prefix_tse, stock_id), 'ab')
        cw = csv.writer(f, lineterminator='\r\n')
        cw.writerow(row)
        f.close()

    def _otc_record(self, stock_id, row):
        ''' Save row to csv file '''
        f = open('{}/{}.csv'.format(self.prefix_otc, stock_id), 'ab')
        cw = csv.writer(f, lineterminator='\r\n')
        cw.writerow(row)
        f.close()

    def _get_tse_data(self, date_str):
        process_count = 0
        date_str_w = date_str.replace("/","")   #strip '/' in date_str
        date_str_c = str((int(date_str[0:4])-1911))+date_str[4:]    #transfer to Taiwan year

        url = 'http://www.twse.com.tw/exchangeReport/MI_INDEX?response=html&date={}&type={}'.format(date_str_w, 'ALLBUT0999')
        page = requests.get(url)

        if not page.ok:
            logging.error("Can not get TSE data at {}".format(date_str_c))
            return


        ''' #TWSE changed its access path, fixed on 2017/06/05
        payload = {
            'download': '',
            'qdate': date_str,
            'selectType': 'ALL'
        }
        url = 'http://www.twse.com.tw/ch/trading/exchange/MI_INDEX/MI_INDEX.php'

        # Get html page and parse as tree
        page = requests.post(url, data=payload)

        if not page.ok:
            logging.error("Can not get TSE data at {}".format(date_str))
            return
        '''
        # Parse page
        tree = html.fromstring(page.text)

        for tr in tree.xpath('//table[5]/tbody/tr'):
            tds = tr.xpath('td/text()')

            sign = tr.xpath('td/font/text()')
            sign = '-' if len(sign) == 1 and sign[0] == u'－' else ''
            stock_id = tds[0].strip()

            row = self._clean_row([
                date_str_c, # 日期
                tds[2], # 成交股數
                #tds[4], # 成交金額
                tds[5], # 開盤價
                tds[6], # 最高價
                tds[7], # 最低價
                tds[8], # 收盤價
                sign + tds[9], # 漲跌價差
                tds[3], # 成交筆數
                tds[4], # 成交金額
            ])
            if (os.path.isfile('{}/{}.csv'.format(self.prefix_tse, stock_id))):
                self._tse_record(tds[0].strip(), row)
                process_count += 1
                self.process(date_str_c, self.tse_files, process_count)

        if (tree.xpath('//table[2]/tbody/tr')):
            print ''
        else:
            print date_str_c + "... none"

    def _get_otc_data(self, date_str):
        ttime = str(int(time.time()*100))
        url = 'http://www.tpex.org.tw/web/stock/aftertrading/daily_close_quotes/stk_quote_result.php?l=zh-tw&d={}&_={}'.format(date_str, ttime)
        page = requests.get(url)
        process_count = 0

        if not page.ok:
            logging.error("Can not get OTC data at {}".format(date_str))
            return

        result = page.json()

        if result['reportDate'] != date_str:
            logging.error("Get error date OTC data at {}".format(date_str))
            return

        for table in [result['mmData'], result['aaData']]:
            for tr in table:
                row = self._clean_row([
                    date_str,
                    tr[8], # 成交股數
                    #tr[9], # 成交金額
                    tr[4], # 開盤價
                    tr[5], # 最高價
                    tr[6], # 最低價
                    tr[2], # 收盤價
                    tr[3], # 漲跌價差
                    tr[10], # 成交筆數
                    tr[9], # 成交金額
                ])
                if (os.path.isfile('{}/{}.csv'.format(self.prefix_otc, tr[0]))):
                    self._otc_record(tr[0], row)
                    process_count += 1
                    self.process(date_str, self.otc_files, process_count)
        if (table):
            print ''
        else:
            print date_str + "... none"

    def table_init(self, year, month, day):
        count = 0
        date_str = '{0}/{1:02d}/{2:02d}'.format(year - 1911, month, day)
        # create tse table
        headers = [ u'日期'.encode('utf-8'), u'成交股數'.encode('utf-8'), u'開盤價'.encode('utf-8'), 
                        u'最高價'.encode('utf-8'), u'最低價'.encode('utf-8'), u'收盤價'.encode('utf-8'), 
                        u'漲跌價差'.encode('utf-8'), u'成交筆數'.encode('utf-8'), u'成交金額'.encode('utf-8')]
        payload = {
            'download': '',
            'qdate': date_str,
            'selectType': 'ALL'
        }
        url = 'http://www.twse.com.tw/ch/trading/exchange/MI_INDEX/MI_INDEX.php'

        # Get html page and parse as tree
        page = requests.post(url, data=payload)

        if not page.ok:
            logging.error("Can not get TSE data at {}".format(date_str))
            return

        # Parse page
        tree = html.fromstring(page.text)

        for tr in tree.xpath('//table[2]/tbody/tr'):
            tds = tr.xpath('td/text()')

            sign = tr.xpath('td/font/text()')
            sign = '-' if len(sign) == 1 and sign[0] == u'－' else ''
            stock_name = tds[1]
            stock_id = tds[0].strip()
            stock_info = [stock_name.encode('utf-8'),'\t%s'%(stock_id)] 

            f = file('{}/{}.csv'.format(self.prefix_tse, stock_id), 'wb')
            f.write(codecs.BOM_UTF8)
            
            cw = csv.writer(f, lineterminator='\r\n')
            cw.writerow(stock_info)
            cw.writerow(headers)
            f.close()
            count += 1
            print 'TSE table counts:' + '%6d\r'%(count),
        print ''    

        # create OTC table
        count = 0
        ttime = str(int(time.time()*100))
        url = 'http://www.tpex.org.tw/web/stock/aftertrading/daily_close_quotes/stk_quote_result.php?l=zh-tw&d={}&_={}'.format(date_str, ttime)
        page = requests.get(url)

        if not page.ok:
            logging.error("Can not get OTC data at {}".format(date_str))
            return

        result = page.json()

        if result['reportDate'] != date_str:
            logging.error("Get error date OTC data at {}".format(date_str))
            return

        for table in [result['mmData'], result['aaData']]:
            for tr in table:
                stock_name = tr[1]
                stock_id = tr[0].strip()
                stock_info = [stock_name.encode('utf-8'),'\t%s'%(stock_id)] 

                f = file('{}/{}.csv'.format(self.prefix_otc, stock_id), 'wb')
                f.write(codecs.BOM_UTF8)
            
                cw = csv.writer(f, lineterminator='\r\n')
                cw.writerow(stock_info)
                cw.writerow(headers)
                f.close()
                count += 1
                print 'OTC table counts:' + '%6d\r'%(count),
        print ''
        print "done!"        

    def get_data(self, data_type, year, month, day):
        date_str_c = '{0}/{1:02d}/{2:02d}'.format(year - 1911, month, day)
        date_str_w = '{0}/{1:02d}/{2:02d}'.format(year,month,day)
        #print 'Crawling {}'.format(date_str)
        #print format(date_str),
    	if ((data_type == 'ALL') or (data_type == 'all')):
            self._get_tse_data(date_str_w)
            self._get_otc_data(date_str_c)
    	elif ((data_type == 'OTC') or (data_type == 'otc')):
            self._get_otc_data(date_str_c)
        else:
            self._get_tse_data(date_str_w)

def get_last_date():
    with open('./tse_raw_data/0050.csv', 'rb') as file:
        reader = csv.reader(file)
        rows = 0
        for row in reader:
            rows += 1
    file.close()
    return row[0]

def main():
    # Set logging
    if not os.path.isdir('log'):
        os.makedirs('log')
    logging.basicConfig(filename='log/crawl-error.log',
        level=logging.ERROR,
        format='%(asctime)s\t[%(levelname)s]\t%(message)s',
        datefmt='%Y/%m/%d %H:%M:%S')

    # Get arguments
    parser = argparse.ArgumentParser(description='Crawl data at assigned day')
    parser.add_argument('data', type=str, nargs='?', help='to download ALL/TSE/OTC data, default is TSE only')
    parser.add_argument('day', type=int, nargs='*',
        help='assigned day (format: YYYY MM DD), default is today')
    parser.add_argument('-b', '--back', action='store_true',
        help='crawl back from assigned day until 2004/2/11')
    parser.add_argument('-c', '--check', action='store_true',
        help='crawl back 10 days for check data')
    parser.add_argument('-i', '--init', action='store_true',
        help='init table headers')

    args = parser.parse_args()

	# Data to download TSE, OTC or ALL
    if ((args.data == 'all') or (args.data == 'ALL') or 
        (args.data == 'otc') or (args.data == 'OTC') or
		(args.data == 'tse') or (args.data == 'TSE')):
        data_type = args.data
    else: 
        data_type = 'TSE'

    #print ("Input data type :"+ data_type)

	# Day only accept 0 or 3 arguments
    if len(args.day) == 0:
        first_day = datetime.today()
    elif len(args.day) == 3:
        first_day = datetime(args.day[0], args.day[1], args.day[2])
    else:
        parser.error('Date should be assigned with (YYYY MM DD) or none')
        return

    crawler = Crawler()

    # If back flag is on, crawl till 2004/2/11, else crawl one day
    if args.back or args.check:
        # otc first day is 2007/04/20
        # tse first day is 2004/02/11
        print 'Crawling ' + data_type + '...'

        last_day = datetime(2004, 2, 11) if args.back else first_day - timedelta(10)
        max_error = 5
        error_times = 0
        print 'last_day @', last_day
        print 'first_day @', first_day
        #while error_times < max_error and first_day >= last_day: #from now to past
        while error_times < max_error and last_day <= first_day:  #from past till now
            try:
                print last_day
                #crawler.get_data(data_type, first_day.year, first_day.month, first_day.day) #from now to past
                crawler.get_data(data_type, last_day.year, last_day.month, last_day.day) #from past till now
                error_times = 0
            except:
                date_str = first_day.strftime('%Y/%m/%d')
                logging.error('Crawl raise error {}'.format(date_str))
                error_times += 1
                continue
            finally:
                #first_day -= timedelta(1) #from now to past
                last_day += timedelta(1) #from past till now
    elif args.init:
        print 'TSE/OTC table init ...'
        date = first_day - timedelta(1)
        crawler.table_init(date.year, date.month, date.day)
    else:
        print 'Crawling ' + data_type + '...'
        if len(args.day) == 0:
            # macho (2016/05/19): auto update TSE data till today 
            # get the last data date from stock '0050'
            lastdate = get_last_date()
            lastyear_c = lastdate.split('/')[0]
            lastyear = int(lastyear_c) + 1911  
            lastmonth = lastdate.split('/')[1]
            lastday = lastdate.split('/')[2]

            last_day = datetime(int(lastyear), int(lastmonth), int(lastday))

            if last_day.strftime('%Y/%m/%d') == first_day.strftime('%Y/%m/%d'):
                print 'your data is up-to-date'
            else:
                print 'updating data from '+last_day.strftime('%Y/%m/%d')+' to '+first_day.strftime('%Y/%m/%d')+'...'
                max_error = 5
                error_times = 0
                last_day = last_day + timedelta(1)
                while error_times < max_error and last_day <= first_day:  #from past till now
                    try:
                        crawler.get_data(data_type, last_day.year, last_day.month, last_day.day) #from past till now
                        error_times = 0
                    except:
                        date_str = first_day.strftime('%Y/%m/%d')
                        logging.error('Crawl raise error {}'.format(date_str))
                        error_times += 1
                        continue
                    finally:
                        #first_day -= timedelta(1) #from now to past
                        last_day += timedelta(1) #from past till now
			time.sleep(10)  # prevent blocked by remote server
        elif len(args.day) == 3:
            # update to specific day: python fetch_tse.py tse year mm dd
            crawler.get_data(data_type, first_day.year, first_day.month, first_day.day)

if __name__ == '__main__':
    main()
