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

RAW_DATA_FOLDER = 'tse_raw_data'
CRAWLING_ERR = False

class Crawler():
    def __init__(self, prefix_tse="tse_earning_raw_data", prefix_otc="otc_earning_raw_data"):
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

    def process(self, stock_id, total, count):
        print "[%5s]"%(stock_id) + "...",
        print "%3d"%(count*100/total) + "%\r",
        time.sleep(30) # crawling data by every 60sec to prevent server block 

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

    def _get_earning_data(self, data_type, stock_id, year, quarter):

        global CRAWLING_ERR

        date_col_00 = "%s, Q%s"%(year, quarter)
        # 損益表 
        data_col=[  u'　　 營業收入合計'.encode('big5'),        u'　　 營業成本合計'.encode('big5'),
                    u'　 營業毛利（毛損）淨額'.encode('big5'),  u'　　 營業費用合計'.encode('big5'),
                    u'　 營業利益（損失）'.encode('big5'),      u'　　 營業外收入及支出合計'.encode('big5'),
                    u'　 本期淨利（淨損）'.encode('big5'),      u'　　 基本每股盈餘合計'.encode('big5')]
        # 當期權益變動表: 本業淨利, 非控制淨利, 股本
        data_col_e1 = u' 本期淨利（淨損）'.encode('big5')
        data_col_e2 = u' 期末餘額'.encode('big5')   # 股本(股) =  期末餘額 / 10

        process_count = 0
        payload = {
            'step': '1',
            'CO_ID': stock_id,
            'SYEAR':year,
            'SSEASON':quarter,
            'REPORT_ID':'C'  #合併財報
        }

        url = 'http://mops.twse.com.tw/server-java/t164sb01'

        # Get html page and parse as tree
        try:
            page = requests.post(url, data=payload)
        except requests.exceptions.ConnectionError:
            logging.error(" requests.exceptions.ConnectionError, try again!")
            time.sleep(180)
            page = requests.post(url, data=payload)



        if not page.ok:
            logging.error("Can not get data at {}".format(date_str))
            return

        # check which charset are using on this web
        #print page.encoding
        #print page.text.encode('ISO-8859-1') 

        # Parse page
        tree = html.fromstring(page.text)

        # check data is valid
        valid_text = u'查無資料'.encode('big5')
        row = []
        for tr in tree.xpath('//h4/font'):
            valid = tr.xpath('text()')
            try:
                if (valid_text == valid[0].encode('ISO-8859-1')):
                    logging.error("%5s  查無資料"%(stock_id))
                    CRAWLING_ERR = True

                    row = self._clean_row([date_col_00, 'n/a']) 

                    if (data_type == 'tse'):
                        if (os.path.isfile('{}/{}.csv'.format(self.prefix_tse, stock_id))):
                            self._tse_record(stock_id, row)
                    elif (data_type == 'otc'):
                        if (os.path.isfile('{}/{}.csv'.format(self.prefix_otc, stock_id))):
                            self._otc_record(stock_id, row)

                    return
                else:
                    break
            except:
                pass 
                
        # table 1: head
        # table 2: balance sheet, table 3: profix & loss statement, table 4: cashflow statement
        row_data = []
#        test = [u'　　 營業收入合計'.encode('big5'), u'　　 母公司業主（淨利／損）', u'　　 非控制權益（淨利／損）']
        valid_data = False
        for item_data in data_col:
            for tr in tree.xpath('//table[3]/tr'):
                tds = tr.xpath('td/text()')
                try:
                    if (item_data == tds[0].encode('ISO-8859-1')):
                        row_data.append(tds[1])
                        valid_data = True
                except:
                    valid_data = False 

            if (valid_data == False):
                CRAWLING_ERR = True
                logging.error("%5s  損益表查無[%15s]資料, 填0代替"%(stock_id, item_data.decode('big5').encode('utf-8')))
                row_data.append('0')


        # 當期權益變動表: 本業淨利, 非控制淨利
        valid_data = False
        for tr in tree.xpath('//table[5]/tr'):
            try:
                tds = tr.xpath('td/text()')
                if (data_col_e1 == tds[0].encode('ISO-8859-1')):
                    row_data.append(tds[len(tds)-3])    #母公司淨利
                    row_data.append(tds[len(tds)-2])    #非控制淨利
                    valid_data = True
            except:
                valid_data = False

        if (valid_data == False):
            CRAWLING_ERR = True
            logging.error("%5s  母公司/非控制淨利查無資料"%(stock_id))
            row_data.append('0')    #母公司淨利
            row_data.append('0')    #非控制淨利

        # 當期權益變動表: 股本
        for tr in tree.xpath('//table[5]/tr'):
            try:
                tds = tr.xpath('td/text()')
                if (data_col_e2 == tds[0].encode('ISO-8859-1')):
                    row_data.append(tds[1])             #股本
                    valid_data = True
                else:
                    pass 
            except: 
                valid_data = False

        if (valid_data == False):
            CRAWLING_ERR = True
            logging.error("%5s  股本查無資料"%(stock_id))
            row_data.append('0')    #股本

        row = self._clean_row([
                date_col_00, # 季度 
                row_data[0], # 營收
                row_data[1], # 成本
                row_data[2], # 毛利
                row_data[3], # 營業費用
                row_data[4], # 營業利益
                row_data[5], # 業外收入
                row_data[6], # 淨利
                row_data[7], # EPS
                row_data[8], # 母公司淨利
                row_data[9], # 非控制淨利
                row_data[10], # 股本
            ])

        if (data_type == 'tse'):
            if (os.path.isfile('{}/{}.csv'.format(self.prefix_tse, stock_id))):
                self._tse_record(stock_id, row)
        elif (data_type == 'otc'):
            if (os.path.isfile('{}/{}.csv'.format(self.prefix_otc, stock_id))):
                self._otc_record(stock_id, row)

    def table_init(self, year, month, day):
        count = 0
        date_str = '{0}/{1:02d}/{2:02d}'.format(year - 1911, month, day)
        # create tse table
        headers = [ u'季度'.encode('utf-8'), u'營收'.encode('utf-8'), u'營業成本'.encode('utf-8'), 
                    u'營業毛利'.encode('utf-8'), u'營業費用'.encode('utf-8'), u'營業利益'.encode('utf-8'), 
                    u'業外收入'.encode('utf-8'), u'淨利'.encode('utf-8'), u'每股盈餘'.encode('utf-8'),
                    u'母公司淨利'.encode('utf-8'), u'非控制淨利'.encode('utf-8'), u'股本'.encode('utf-8')]
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
        os.system('sh rm_earnings.sh')

    def get_data(self, year, quarter):

        global CRAWLING_ERR

        process_count = 0
        print "crawling TSE earnings data @ " + str(year) + ' Q' + str(quarter) + "..."
        for filename in glob.glob("./{}/*.csv".format(self.prefix_tse)):
            (f_path, f_name) = os.path.split(filename)
            (f_short_name, f_extension) = os.path.splitext(f_name)
            self._get_earning_data('tse', f_short_name, year, quarter)
            process_count += 1
            self.process(f_short_name, self.tse_files, process_count)

        process_count = 0
        print "crawling OTC earnings data @ " + str(year) + ' Q' + str(quarter) + "..."
        for filename in glob.glob("./{}/*.csv".format(self.prefix_otc)):
            (f_path, f_name) = os.path.split(filename)
            (f_short_name, f_extension) = os.path.splitext(f_name)
            self._get_earning_data('otc', f_short_name, year, quarter)
            process_count += 1
            self.process(f_short_name, self.otc_files, process_count)

        if (CRAWLING_ERR == True):
            today = time.strftime("%Y%m%d") 
            os.system('cat ./log/%s.log'%(today))
        else:
            print("\r\ndone")


def main():
    # Set logging
    today = time.strftime("%Y%m%d") 
    if not os.path.isdir('log'):
        os.makedirs('log')
    logging.basicConfig(filename='log/%s.log'%(today),
        level=logging.ERROR,
        format='%(asctime)s\t[%(levelname)s]\t%(message)s',
        datefmt='%Y/%m/%d %H:%M')


    # Get arguments
    parser = argparse.ArgumentParser(description='Crawl TSE/OTC earnings data at assigned date')
    parser.add_argument('date', type=int, nargs='*', help='assigned date (format: YYYY Q)')
    parser.add_argument('-i', '--init', action='store_true', help='init table headers (-i YYYY MM DD)')

    args = parser.parse_args()

	# Data to download TSE and OTC
    #data_type = 'ALL'

    #print ("Input data type :"+ data_type)

	# Day only accept 1 or 2 arguments
    if len(args.date) == 0:
        parser.print_help()
    elif len(args.date) == 1:
        parser.error('lost arguments')
        return
    elif len(args.date) == 2:
        earnings_year = args.date[0]
        earnings_quarter = args.date[1]
        if (earnings_year < 2013):
            parser.error('new IFRS: year must after 2013')
            return
        elif (earnings_quarter > 4):
            parser.error('season must Q1 to Q4')
            return
    elif len(args.date) == 3:
        if (args.init == True):
            date = datetime(args.date[0], args.date[1], args.date[2])
        else:
            parser.print_help()
    else:
        parser.error('Date should be assigned with (YYYY Q) or none')
        return

    crawler = Crawler()

    if args.init:
        print 'TSE/OTC earnings table init ...'
        if (date.weekday() >4):
            print "NOT a working day!!!"
        else:
            crawler.table_init(date.year, date.month, date.day)
        
    elif len(args.date) == 2:
        # update to specific day: python fetch_tse.py YEAR Q 
        crawler.get_data(earnings_year, earnings_quarter)

if __name__ == '__main__':
    main()