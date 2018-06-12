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
from dateutil.relativedelta import relativedelta

from os import mkdir
from os.path import isdir

class Crawler():
    def __init__(self, prefix_monthly="monthly_raw_data"):
        ''' Make directory if not exist when initialize '''
        if not isdir(prefix_monthly):
            mkdir(prefix_monthly)
        self.prefix_monthly = prefix_monthly
        # total files in both TSE/OTC
        monthly_files = sum(os.path.isfile(os.path.join('{}/'.format(self.prefix_monthly), f)) 
                        for f in os.listdir('{}/'.format(self.prefix_monthly))) 
        self.monthly_files = monthly_files

    def process(self, date, total, count):
        print date + "...",
        print "%3d"%(count*100/total) + "%\r",

    def _clean_row(self, row):
        ''' Clean comma and spaces '''
        for index, content in enumerate(row):
            row[index] = re.sub(",", "", content.strip())
            row[index] = filter(lambda x: x in string.printable, row[index])
        return row

    def _monthly_record(self, stock_id, row):
        ''' Save row to csv file '''
        f = open('{}/{}.csv'.format(self.prefix_monthly, stock_id), 'ab')
        cw = csv.writer(f, lineterminator='\r\n')
        cw.writerow(row)
        f.close()

    def _get_monthly_data(self, date_str_w):
        process_count = 0
        #date_str_w = date_str.replace("/","")   #strip '/' in date_str
        date_str_c_year = str((int(date_str_w[0:4])-1911))              #transfer to Taiwan year
        date_str_c_month = str('{:01d}'.format(int(date_str_w[5:7])))   #transfer to Taiwan year


        url_0 = 'http://mops.twse.com.tw/nas/t21/sii/t21sc03_{}_{}_0.html'.format(date_str_c_year, date_str_c_month) # 國內上市
        url_1 = 'http://mops.twse.com.tw/nas/t21/sii/t21sc03_{}_{}_1.html'.format(date_str_c_year, date_str_c_month) # 國外上市
        url_2 = 'http://mops.twse.com.tw/nas/t21/otc/t21sc03_{}_{}_0.html'.format(date_str_c_year, date_str_c_month) # 國內上櫃
        url_3 = 'http://mops.twse.com.tw/nas/t21/otc/t21sc03_{}_{}_1.html'.format(date_str_c_year, date_str_c_month) # 國外上櫃

        fetch_url = [url_0, url_1, url_2, url_3]
        fetch_name= [u'國內上市', u'國外上市', u'國內上櫃', u'國外上櫃']

        for url, url_name in zip(fetch_url,fetch_name):

            page = requests.get(url)
            page.encoding = 'big5'
 
            if not page.ok:
                logging.error("Can not get {} data at {}".format(url_name, date_str_w))
                return

            # Parse page
            tree = html.fromstring(page.text)
 
            if not (tree.xpath('//table[@bgcolor="#FFFFFF"]/tr')):
                print date_str_w + url_name + "... none"
            else:
                for tr in tree.xpath('//table[@bgcolor="#FFFFFF"]/tr'):
                    tds = tr.xpath('td/text()')
                    if (not tds) or (len(tds) != 11):
                        # empty data or elements is not equal to 11, read next!!
                        continue
  
                    #sign = tr.xpath('td/font/text()')
                    #sign = '-' if len(sign) == 1 and sign[0] == u'－' else ''
  
                    stock_id = tds[0].strip()
                    
                    row = self._clean_row([
                        date_str_w, # 日期
                        tds[2],     # 當月營收 
                        tds[3],     # 上月營收
                        tds[4],     # 去年同月營收
                        tds[5],     # MOM(%)
                        tds[6],     # YOY(%)
                        tds[7],     # 當月累計營收
                        tds[8],     # 去年累計營收
                        tds[9],     # 累計YOY(%)
                        ])
                    row.append(tds[10].encode('utf-8')) # 備註
 
                    if (os.path.isfile('{}/{}.csv'.format(self.prefix_monthly, stock_id))):
                        self._monthly_record(tds[0].strip(), row)
                        process_count += 1
                        self.process(date_str_w, self.monthly_files, process_count)

        print ''

    def table_init(self, year, month, day):
        process_count = 0
        date_str = '{0}/{1:02d}/{2:02d}'.format(year,month,day)
        date_str_w = date_str.replace("/","")   #strip '/' in date_str
        date_str_c = str((int(date_str[0:4])-1911))+date_str[4:]    #transfer to Taiwan year

        headers = [ u'日期'.encode('utf-8'), u'當月營收'.encode('utf-8'), u'上月營收'.encode('utf-8'), 
                    u'去年同月營收'.encode('utf-8'), u'MOM(%)'.encode('utf-8'), u'YOY(%)'.encode('utf-8'), 
                    u'當月累計營收'.encode('utf-8'), u'去年累計營收'.encode('utf-8'), u'累計YOY(%)'.encode('utf-8'),
                    u'Notes'.encode('utf-8')]

        
        url_tse = 'http://www.twse.com.tw/exchangeReport/MI_INDEX?response=html&date={}&type={}'.format(date_str_w, 'ALLBUT0999')
        url_otc = 'http://www.tpex.org.tw/web/stock/aftertrading/otc_quotes_no1430/stk_wn1430_result.php?l=zh-tw&o=htm&d={}&se=EW&s=0,asc,0'.format(date_str_c)

        xpath_tse = '//table[5]/tbody/tr'
        xpath_otc = '//table[1]/tbody/tr'
        
        fetch_url   = [url_tse, url_otc]
        fetch_name  = ['TSE', 'OTC']
        xpath_set_t = [xpath_tse, xpath_otc]

        for url, url_name, xpath_set in zip(fetch_url, fetch_name, xpath_set_t):
            ## create tse table
            # Get html page and parse as tree
            page = requests.get(url)
 
            if not page.ok:
                logging.error("Can not get {} data at {}".format(url_name, date_str_w))
                return
 
            # Parse page
            tree = html.fromstring(page.text)
 
            for tr in tree.xpath(xpath_set):
                tds = tr.xpath('td/text()')
 
                #sign = tr.xpath('td/font/text()')
                #sign = '-' if len(sign) == 1 and sign[0] == u'－' else ''
                stock_id = tds[0].strip()
                stock_name = tds[1]
                stock_info = [stock_name.encode('utf-8'),'\t%s'%(stock_id)] 
 
                f = file('{}/{}.csv'.format(self.prefix_monthly, stock_id), 'wb')
                f.write(codecs.BOM_UTF8)
                
                cw = csv.writer(f, lineterminator='\r\n')
                cw.writerow(stock_info)
                cw.writerow(headers)
                f.close()
                process_count += 1
                print '%s table counts: %6d\r'%(url_name, process_count),
            print ''    

        print "done!"        

        os.system('sh rm_monthly.sh')

    def check_new_stock(self, year, month, day):
        process_count = 0
        date_str = '{0}/{1:02d}/{2:02d}'.format(year,month,day)
        date_str_w = date_str.replace("/","")   #strip '/' in date_str
        date_str_c = str((int(date_str[0:4])-1911))+date_str[4:]    #transfer to Taiwan year

        headers = [ u'日期'.encode('utf-8'), u'當月營收'.encode('utf-8'), u'上月營收'.encode('utf-8'),
                    u'去年同月營收'.encode('utf-8'), u'MOM(%)'.encode('utf-8'), u'YOY(%)'.encode('utf-8'),
                    u'當月累計營收'.encode('utf-8'), u'去年累計營收'.encode('utf-8'), u'累計YOY(%)'.encode('utf-8'),
                    u'Notes'.encode('utf-8')]


        url_tse = 'http://www.twse.com.tw/exchangeReport/MI_INDEX?response=html&date={}&type={}'.format(date_str_w, 'ALLBUT0999')
        url_otc = 'http://www.tpex.org.tw/web/stock/aftertrading/otc_quotes_no1430/stk_wn1430_result.php?l=zh-tw&o=htm&d={}&se=EW&s=0,asc,0'.format(date_str_c)

        xpath_tse = '//table[5]/tbody/tr'
        xpath_otc = '//table[1]/tbody/tr'

        fetch_url   = [url_tse, url_otc]
        fetch_name  = ['TSE', 'OTC']
        xpath_set_t = [xpath_tse, xpath_otc]

        for url, url_name, xpath_set in zip(fetch_url, fetch_name, xpath_set_t):
            ## create tse table
            # Get html page and parse as tree
            page = requests.get(url)

            if not page.ok:
                logging.error("Can not get {} data at {}".format(url_name, date_str_w))
                return

            # Parse page
            tree = html.fromstring(page.text)

            for tr in tree.xpath(xpath_set):
                tds = tr.xpath('td/text()')

                stock_id = tds[0].strip()
                stock_name = tds[1]
                stock_info = [stock_name.encode('utf-8'),'\t%s'%(stock_id)]

                if (stock_id[0] == '0' or stock_id[:2] =='91' or len(stock_id) != 4):
                    continue

                if not (os.path.isfile('{}/{}.csv'.format(self.prefix_monthly, stock_id))):
                    f = file('{}/{}.csv'.format(self.prefix_monthly, stock_id), 'wb')
                    f.write(codecs.BOM_UTF8)

                    cw = csv.writer(f, lineterminator='\r\n')
                    cw.writerow(stock_info)
                    cw.writerow(headers)
                    f.close()
                    process_count += 1
                    print 'add new %s stock [%s] %s \r'%(url_name, stock_id, stock_name)

        print ''
        print 'Check new stock is not warrants or ETF type... \r'
        os.system('sh rm_monthly.sh')

    def get_data(self, year, month):
        date_str_c = '{0}/{1:02d}'.format(year - 1911, month)   # 20xx/xx
        date_str_w = '{0}/{1:02d}'.format(year,month)           # 10x/xx
        #print 'Crawling {}'.format(date_str)
        #print format(date_str),
        self._get_monthly_data(date_str_w)

def get_last_date():
    with open('./monthly_raw_data/1101.csv', 'rb') as file:
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
    parser = argparse.ArgumentParser(description='Crawl monthly-revenue data at assigned month')
    parser.add_argument('day', type=int, nargs='*',
        help='assigned day (format: YYYY MM), default is this month')
    parser.add_argument('-b', '--back', action='store_true',
        help='crawl back from assigned day until Jan. 2013 ')
    parser.add_argument('-i', '--init', action='store_true',
        help='init table headers')

    args = parser.parse_args()

    # Day only accept 0 or 2 arguments
    if len(args.day) == 0:
        first_day = datetime.today()
    elif len(args.day) == 2:
        first_day = datetime(args.day[0], args.day[1], 15)
    else:
        parser.error('Date should be assigned with (YYYY MM) or none')
        return

    crawler = Crawler()

    # If back flag is on, crawl till Jan. 2013, else crawl one month
    if args.back:
        print 'Crawling TSE/OTC monthly revenue...'

        last_day = datetime(2013, 1, 15)
        max_error = 3 
        error_times = 0
        print 'last_day @', last_day

        # TWSE announce monthlyi-revenue on every 15th day of month,
        # use this as dummy data and prevent data is not updated on TWSE
        if first_day.day < 15:
            first_day -= relativedelta(months=1)

        #while error_times < max_error and first_day >= last_day: #from now to past
        while error_times < max_error and last_day <= first_day:  #from past till now
            try:
                #crawler.get_data(first_day.year, first_day.month) #from now to past
                crawler.get_data(last_day.year, last_day.month) #from past till now
                error_times = 0
            except:
                date_str = first_day.strftime('%Y/%m')
                logging.error('Crawl raise error {}'.format(date_str))
                error_times += 1
                continue
            finally:
                #first_day -= relativedelta(months=1) #from now to past
                last_day += relativedelta(months=1) #from past till now
    elif args.init:
        print 'TSE/OTC table init ...'
        date = first_day - timedelta(1)
        crawler.table_init(date.year, date.month, date.day)
    else:
        print 'Checking is there a new stock...'
        weekend = first_day.weekday()
        if weekend < 5:
            date = first_day
        else:
            date = first_day - timedelta(weekend % 4)
        crawler.check_new_stock(date.year, date.month, date.day)
        print ''
        print 'Crawling TSE/OTC monthly revenue...'
        if len(args.day) == 0:
            # auto update data till this month 
            # get the last data date from stock '1101'
            lastdate = get_last_date()
            lastyear = lastdate.split('/')[0]
            #lastyear = int(lastyear_c) + 1911  
            lastmonth = lastdate.split('/')[1]
            #lastday = lastdate.split('/')[2]

            # TWSE announce monthlyi-revenue on every 15th day of month,
            # use this as dummy data and prevent data is not updated on TWSE
            last_day = datetime(int(lastyear), int(lastmonth), 15)
            if first_day.day < 15:
                first_day -= relativedelta(months=1)
            update_data_to = first_day-relativedelta(months=1)

            # start fetching data
            if last_day.strftime('%Y/%m') == first_day.strftime('%Y/%m'):
                print 'your data is up-to-date'
            else:
                print 'updating data from '+last_day.strftime('%Y/%m')+' to '+update_data_to.strftime('%Y/%m')+'...'
                max_error = 5
                error_times = 0
                last_day = last_day + relativedelta(months=1)
                while error_times < max_error and last_day <= first_day:  #from past till now
                    try:
                        crawler.get_data(last_day.year, last_day.month) #from past till now
                        error_times = 0
                    except:
                        date_str = first_day.strftime('%Y/%m')
                        logging.error('Crawl raise error {}'.format(date_str))
                        error_times += 1
                        continue
                    finally:
                        last_day += relativedelta(months=1) #from past till now
			time.sleep(10)  # prevent blocked by remote server
                print 'done'
        elif len(args.day) == 2:
            # update to specific day: python fetch_tse.py tse year mm dd
            crawler.get_data(first_day.year, first_day.month)

if __name__ == '__main__':
    main()
