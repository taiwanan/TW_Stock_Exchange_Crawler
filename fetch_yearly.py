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

last_date_of_file = None
net_asset = []

class Crawler():
    def __init__(self, prefix_trading_yearly="yearly_trading_raw_data", prefix_earning_yearly="yearly_earning_raw_data"):
        self.prefix_trading_yearly = prefix_trading_yearly
        self.prefix_earning_yearly = prefix_earning_yearly
        ''' Make directory if not exist when initialize '''
        if not (isdir(prefix_trading_yearly) and isdir(prefix_earning_yearly)):
            print 'Warning! there is no directory [{} or {}]'.format(prefix_trading_yearly, prefix_earning_yearly)
            #mkdir(prefix_trading_yearly)
        else:
            # total earning files in both TSE/OTC
            number_of_yearly_trading_files = sum(os.path.isfile(os.path.join('{}/'.format(self.prefix_trading_yearly), f)) 
                            for f in os.listdir('{}/'.format(self.prefix_trading_yearly))) 
            self.number_of_yearly_trading_files = number_of_yearly_trading_files
            number_of_yearly_earning_files = sum(os.path.isfile(os.path.join('{}/'.format(self.prefix_earning_yearly), f)) 
                            for f in os.listdir('{}/'.format(self.prefix_earning_yearly))) 
            self.number_of_yearly_earning_files = number_of_yearly_earning_files

    def get_last_date(self, stock_id, fetch_type):
        if fetch_type == 'trading':
            directory = self.prefix_trading_yearly
        elif fetch_type == 'earning':
            directory = self.prefix_earning_yearly
             
        with open('./{}/{}.csv'.format(directory, stock_id), 'rb') as file:
            reader = csv.reader(file)
            rows = 0
            for row in reader:
                rows += 1
        file.close()
        return row[0]

    def process(self, message, total, count):
        print message + "...",
        print "%3d"%(count*100/total) + "%\r",

    def _clean_row(self, row):
        ''' Clean comma and spaces '''
        for index, content in enumerate(row):
            row[index] = re.sub(",", "", content.strip())
            row[index] = filter(lambda x: x in string.printable, row[index])
        return row

    def _raw_data_record(self, stock_id, row, fetch_type):
        ''' Save row to csv file '''
        if fetch_type == 'trading':
            directory = self.prefix_trading_yearly
        elif fetch_type == 'earning':
            directory = self.prefix_earning_yearly
        f = open('{}/{}.csv'.format(directory, stock_id), 'ab')
        cw = csv.writer(f, lineterminator='\r\n')
        cw.writerow(row)
        f.close()

    def _get_trading_data(self, stock_id):

        global last_date_of_file
        max_error = 5
        error_times = 0 
        process_count = 0
        valid_data = False 
        #date_str_w = date_str.replace("/","")   #strip '/' in date_str
        tds_data=[]

        invalid_text_tse = u'很抱歉，沒有符合條件的資料!'.encode('utf-8')
        invalid_text_otc = u'查無該筆資料,請重新查詢!!'.encode('utf-8')

        url_tse = 'http://www.twse.com.tw/exchangeReport/FMNPTK?response=html&stockNo={}'.format(stock_id)
        url_otc = 'http://www.tpex.org.tw/web/stock/statistics/monthly/st42.php'

        payload_otc = {'input_stock_code':stock_id}

        xpath_tse = '//table[1]/tbody/tr'
        xpath_otc = '//table[2]/tr'

        ## load web data
        # try TSE first
        for error_times in range(max_error):
            try:
                page = requests.get(url_tse)
            except:
                time.sleep(300)  # prevent remote server block
                if error_times == (max_error-1):
                    logging.error('Crawl raise error in TSE stock[{}]'.format(stock_id))
                    return
            else:
                page.encoding = 'utf-8'
                break

        #print page.text.encode('utf-8')

        if not page.ok:
            logging.error("Can not get TSE yearly-trading data {}".format(stock_id))
            return

        # Parse page
        tree = html.fromstring(page.text)

        valid = tree.xpath('//div/text()')
        if (invalid_text_tse == valid[0].strip().encode('utf-8')):
            valid_data = False 
        else:
            valid_data = True
            xpath_t = xpath_tse
            valid_data_type = 'TSE'

        # try OTC then 
        if not valid_data:
            error_times = 0
            for error_times in range(max_error):
                try:
                    page = requests.post(url_otc, data=payload_otc)
                except:
                    time.sleep(300)  # prevent remote server block
                    if error_times == (max_error-1):
                        logging.error('Crawl raise error in TSE stock[{}]'.format(stock_id))
                        return
                else:
                    page.encoding = 'utf-8'
                    break

 
            if not page.ok:
                logging.error("Can not get OTC yearly-trading data {}".format(stock_id))
                return
        
            # Parse page
            tree = html.fromstring(page.text)
 
            valid = tree.xpath('//table[1]/tr/td/text()')
            if (invalid_text_otc == valid[0].strip().encode('utf-8')):
                logging.error("Can not get {} data".format(stock_id))
                return
            else:
                xpath_t = xpath_otc
                valid_data_type = 'OTC'

        ## save data to tds_data
        for tr_idx in tree.xpath(xpath_t):
            tds_idx = tr_idx.xpath('td/text()')

            if (not tds_idx) or (len(tds_idx) != 9):
                # empty data or elements is not equal to 11, read next!!
                continue

            if valid_data_type == 'TSE':
                tds_data.append(tds_idx)
            elif valid_data_type == 'OTC':
                # reverse date list - starts from old to new
                tds_data.insert(0,tds_idx)


        ## write data to file
        for tds_idx in tds_data:

            date_fetch = tds_idx[0]
            date_record = '{}'.format(str(int(date_fetch[:3])+1911))    # raw data is Chinese Year, trans to Western Year
                
            row = self._clean_row([
                date_record, # 年度
                tds_idx[4],  # 最高價 
                tds_idx[5],  # 最高價日期
                tds_idx[6],  # 最低價
                tds_idx[7],  # 最低價日期
                tds_idx[8],  # 收盤平均價
            ])

            # check date to prevent overwrite at the same date
            date_check_old = time.strptime(last_date_of_file, "%Y")
            date_check_new = time.strptime(row[0], "%Y")

            if date_check_old < date_check_new:
                self._raw_data_record(stock_id, row, 'trading')

        time.sleep(60)  # prevent blocked by remote server

    def _get_balance_data(self, year_w):    # 取第四季淨值當成年度淨值

        global last_date_of_file
        global net_asset
        net_asset = []
        net_asset_tmp = []
        process_count = 0
        date_str_c_year = str((int(year_w)-1911))              #transfer to Taiwan year

        payload_tse = {
            'ifrs':'N', # apply IFRS after 2011 (>2011, partial), and ALL after 2014 (>2014)
            'step':'1',
            'firstin':'ture',
            'off':'1',
            'isQuery':'Y',
            'TYPEK':'sii',  # TSE => 'sii' ; OTC => 'otc'
            'season':'04',
            'year':date_str_c_year
        }

        payload_otc = {'ifrs':'N', 'step':'1', 'firstin':'ture', 'off':'1', 'isQuery':'Y', 'TYPEK':'otc', 'season':'04', 'year':date_str_c_year}
        payload_tse_ifrs = {'ifrs':'Y', 'step':'1', 'firstin':'ture', 'off':'1', 'isQuery':'Y', 'TYPEK':'sii', 'season':'04', 'year':date_str_c_year}
        payload_otc_ifrs = {'ifrs':'Y', 'step':'1', 'firstin':'ture', 'off':'1', 'isQuery':'Y', 'TYPEK':'otc', 'season':'04', 'year':date_str_c_year}

        payload_type_old = [payload_tse, payload_otc]
        payload_type_mix = [payload_tse, payload_otc, payload_tse_ifrs, payload_otc_ifrs]
        payload_type_new = [payload_tse_ifrs, payload_otc_ifrs]

        url_ifrs_yes = 'http://mops.twse.com.tw/mops/web/ajax_t163sb05'
        url_ifrs_no = 'http://mops.twse.com.tw/mops/web/ajax_t51sb07'

        if int(year_w) < 2012:
            payload = payload_type_old
        elif int(year_w) <2015:
            payload = payload_type_mix
        else:
            payload = payload_type_new

        for payload_to_fetch in (payload):

            if payload_to_fetch['ifrs'] == 'Y': # IFRS
                url = url_ifrs_yes
            else:
                url = url_ifrs_no

            try:
                page = requests.post(url, data=payload_to_fetch)
            except requests.exceptions.ConnectionError:
                logging.error(" requests.exceptions.ConnectionError, try again!")
                time.sleep(180)
                page = requests.post(url, data=payload_to_fetch)

            page.encoding = 'utf-8'
            #print page.text.encode('utf-8')

            if not page.ok:
                logging.error("Can not get {} data".format(payload_to_fetch))
                return

            # Parse page
            tree = html.fromstring(page.text)

            for table_count in range(1,8):  # assume have 8 tables to fetch
                for tr_idx in tree.xpath('//table[%s][@class="hasBorder"]/tr'%table_count):
                    tds_idx = tr_idx.xpath('td/text()')

                    if not tds_idx:     # no data found
                        continue

                    net_asset_tmp = []
                    net_asset_tmp.append(tds_idx[0])

                    if payload_to_fetch['ifrs'] == 'Y': # IFRS
                        net_asset_tmp.append(tds_idx[-1].strip())
                    else:
                        if int(payload_to_fetch['year']) < 95:  # < 2006, last value is net asset
                            net_asset_tmp.append(tds_idx[-1].strip())
                        else:
                            net_asset_tmp.append(tds_idx[-3].strip())

                    net_asset.append(net_asset_tmp)



    def _get_earning_data(self, year_w):

        global last_date_of_file
        global net_asset
        process_count = 0
        #date_str_w = date_str.replace("/","")   #strip '/' in date_str
        date_str_c_year = str((int(year_w)-1911))              #transfer to Taiwan year

        # get net asset
        self._get_balance_data(year_w)

        payload_tse = {
            'ifrs':'N', # apply IFRS after 2011 (>2011, partial), and ALL after 2014 (>2014)
            'step':'1',
            'firstin':'ture',
            'off':'1',
            'TYPEK':'sii',  # TSE => 'sii' ; OTC => 'otc'
            'year':date_str_c_year
        }

        payload_otc = {'ifrs':'N', 'step':'1', 'firstin':'ture', 'off':'1', 'TYPEK':'otc', 'year':date_str_c_year}
        payload_tse_ifrs = {'ifrs':'Y', 'step':'1', 'firstin':'ture', 'off':'1', 'TYPEK':'sii', 'year':date_str_c_year}
        payload_otc_ifrs = {'ifrs':'Y', 'step':'1', 'firstin':'ture', 'off':'1', 'TYPEK':'otc', 'year':date_str_c_year}

        payload_type_old = [payload_tse, payload_otc] 
        payload_type_mix = [payload_tse, payload_otc, payload_tse_ifrs, payload_otc_ifrs]
        payload_type_new = [payload_tse_ifrs, payload_otc_ifrs]

        url = 'http://mops.twse.com.tw/mops/web/ajax_t51sb02'

        if int(year_w) < 2012:
            payload = payload_type_old
        elif int(year_w) <2015:
            payload = payload_type_mix
        else:
            payload = payload_type_new

        for payload_to_fetch in (payload):

            try:
                page = requests.post(url, data=payload_to_fetch)
            except requests.exceptions.ConnectionError:
                logging.error(" requests.exceptions.ConnectionError, try again!")
                time.sleep(180)
                page = requests.post(url, data=payload_to_fetch)

            page.encoding = 'utf-8'
            #print page.text.encode('utf-8')
 
            if not page.ok:
                logging.error("Can not get {} data".format(payload_to_fetch))
                return
 
            # Parse page
            tree = html.fromstring(page.text)
 
            for tr_idx in tree.xpath('//table[3]/tr'):
                tds_idx = tr_idx.xpath('td/text()')
 
                if not tds_idx:     # no data found
                    continue 
            
                stock_id = tds_idx[0].strip()

                for net_asset_tmp in net_asset:
                   if net_asset_tmp[0] == stock_id:
                       net_asset_value = net_asset_tmp[1]
                       break
                   else:
                       net_asset_value = 'n/a'

                if payload_to_fetch['ifrs'] == 'Y': # IFRS
                    row = self._clean_row([
                        year_w,      # 年度(西元)
                        tds_idx[2],  # 負債佔資產比(%) 
                        tds_idx[4],  # 流動比(%)
                        tds_idx[5],  # 速動比(%)
                        tds_idx[7],  # 應收款周轉率(次) 
                        tds_idx[9],  # 存貨周轉率(次)
                        tds_idx[10], # 平均銷貨日 
                        tds_idx[13], # 資產報酬率(%)
                        tds_idx[14], # ROE(%)
                        tds_idx[16], # 純益率(%)
                        tds_idx[17], # EPS
                        tds_idx[18], # 現金流量比(%)
                        tds_idx[20], # 現金再投比(%)
                        net_asset_value, #淨值
                    ])
                else:   # before IFRS
                    row = self._clean_row([
                        year_w,      # 年度(西元)
                        tds_idx[2],  # 負債佔資產比(%) 
                        tds_idx[4],  # 流動比(%)
                        tds_idx[5],  # 速動比(%)
                        tds_idx[7],  # 應收款周轉率(次) 
                        tds_idx[9],  # 存貨周轉率(次)
                        tds_idx[10], # 平均銷貨日 
                        tds_idx[13], # 資產報酬率(%)
                        tds_idx[14], # ROE(%)
                        tds_idx[17], # 純益率(%)
                        tds_idx[18], # EPS
                        tds_idx[19], # 現金流量比(%)
                        tds_idx[21], # 現金再投比(%)
                        net_asset_value, #淨值
                    ])
 
                # check date to prevent overwrite at the same date
                if (os.path.isfile('{}/{}.csv'.format(self.prefix_earning_yearly, stock_id))):
                    lastdate = self.get_last_date(stock_id, 'earning')
                    if lastdate == 'Year':
                        last_date_of_file = '1991'  # crawling data from 1991 (>1990)
                    else:
                        last_date_of_file = lastdate
                    date_check_old = time.strptime(last_date_of_file, "%Y")
                    date_check_new = time.strptime(row[0], "%Y")
 
                    if (date_check_old < date_check_new) or (last_date_of_file == '1991'):
                        self._raw_data_record(stock_id, row, 'earning')

                        process_count += 1
                        self.process('%.4s, stock[%.5s]'%(year_w,stock_id), self.number_of_yearly_earning_files, process_count)
            time.sleep(60)  # prevent blocked by remote server

    def table_init(self, year, month, day):
        process_count = 0
        date_str = '{0}/{1:02d}/{2:02d}'.format(year,month,day)
        date_str_w = date_str.replace("/","")   #strip '/' in date_str
        date_str_c = str((int(date_str[0:4])-1911))+date_str[4:]    #transfer to Taiwan year
        
        headers_trading = [ u'Year'.encode('utf-8'), u'最高價'.encode('utf-8'), u'日期'.encode('utf-8'), 
                            u'最低價'.encode('utf-8'), u'日期'.encode('utf-8'), u'收盤平均價'.encode('utf-8')]

        headers_earning = [ u'Year'.encode('utf-8'), u'負債佔資產比(%)'.encode('utf-8'), u'流動比(%)'.encode('utf-8'), 
                            u'速動比(%)'.encode('utf-8'), u'應收款周轉率(次)'.encode('utf-8'), u'存貨周轉率(次)'.encode('utf-8'),
                            u'平均銷貨日'.encode('utf-8'), u'資產報酬率(%)'.encode('utf-8'), u'ROE(%)'.encode('utf-8'),
                            u'純益率(%)'.encode('utf-8'), u'EPS'.encode('utf-8'), u'現金流量比(%)'.encode('utf-8'),u'現金再投比(%)'.encode('utf-8'),
                            u'淨值'.encode('utf-8')]

        
        url_tse = 'http://www.twse.com.tw/exchangeReport/MI_INDEX?response=html&date={}&type={}'.format(date_str_w, 'ALLBUT0999')
        url_otc = 'http://www.tpex.org.tw/web/stock/aftertrading/otc_quotes_no1430/stk_wn1430_result.php?l=zh-tw&o=htm&d={}&se=EW&s=0,asc,0'.format(date_str_c)

        xpath_tse = '//table[5]/tbody/tr'
        xpath_otc = '//table[1]/tbody/tr'
        
        fetch_url   = [url_tse, url_otc]
        fetch_name  = ['TSE', 'OTC']
        xpath_set_t = [xpath_tse, xpath_otc]


        for url, url_name, xpath_set in zip(fetch_url, fetch_name, xpath_set_t):

            process_count = 0

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
 
                # trading table
                f_trading = file('{}/{}.csv'.format(self.prefix_trading_yearly, stock_id), 'wb')
                f_trading.write(codecs.BOM_UTF8)
                
                cw = csv.writer(f_trading, lineterminator='\r\n')
                cw.writerow(stock_info)
                cw.writerow(headers_trading)
                f_trading.close()

                # earning table
                f_earning = file('{}/{}.csv'.format(self.prefix_earning_yearly, stock_id), 'wb')
                f_earning.write(codecs.BOM_UTF8)
                
                cw = csv.writer(f_earning, lineterminator='\r\n')
                cw.writerow(stock_info)
                cw.writerow(headers_earning)
                f_earning.close()

                process_count += 1
                print '%s table counts: %6d\r'%(url_name, process_count),
            print ''    

        print "done!"        

        os.system('sh rm_yearly.sh')

    def check_new_stock(self, year, month, day):
        date_str = '{0}/{1:02d}/{2:02d}'.format(year,month,day)
        date_str_w = date_str.replace("/","")   #strip '/' in date_str
        date_str_c = str((int(date_str[0:4])-1911))+date_str[4:]    #transfer to Taiwan year

        headers_trading = [ u'Year'.encode('utf-8'), u'最高價'.encode('utf-8'), u'日期'.encode('utf-8'), 
                            u'最低價'.encode('utf-8'), u'日期'.encode('utf-8'), u'收盤平均價'.encode('utf-8')]

        headers_earning = [ u'Year'.encode('utf-8'), u'負債佔資產比(%)'.encode('utf-8'), u'流動比(%)'.encode('utf-8'), 
                            u'速動比(%)'.encode('utf-8'), u'應收款周轉率(次)'.encode('utf-8'), u'存貨周轉率(次)'.encode('utf-8'),
                            u'平均銷貨日'.encode('utf-8'), u'資產報酬率(%)'.encode('utf-8'), u'ROE(%)'.encode('utf-8'),
                            u'純益率(%)'.encode('utf-8'), u'EPS'.encode('utf-8'), u'現金流量比(%)'.encode('utf-8'),u'現金再投比(%)'.encode('utf-8')]

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

                if not (os.path.isfile('{}/{}.csv'.format(self.prefix_trading_yearly, stock_id))):
                    # trading table
                    f_trading = file('{}/{}.csv'.format(self.prefix_trading_yearly, stock_id), 'wb')
                    f_trading.write(codecs.BOM_UTF8)
                    
                    cw = csv.writer(f_trading, lineterminator='\r\n')
                    cw.writerow(stock_info)
                    cw.writerow(headers_trading)
                    f_trading.close()
 
                    # earning table
                    f_earning = file('{}/{}.csv'.format(self.prefix_earning_yearly, stock_id), 'wb')
                    f_earning.write(codecs.BOM_UTF8)
                    
                    cw = csv.writer(f_earning, lineterminator='\r\n')
                    cw.writerow(stock_info)
                    cw.writerow(headers_earning)
                    f_earning.close()

                    print 'add new %s stock [%s] %s \r'%(url_name, stock_id, stock_name.encode('utf-8'))

        print ''
        print 'Check new stock is not warrants or ETF type... \r'
        os.system('sh rm_yearly.sh')

    def get_data(self, year_w):

        global last_date_of_file
        process_count = 0

        # fetch yearly data of trading
        print '### crawling yearly data --- TRADING ###'
        for filename in sorted(glob.glob("./{}/*.csv".format(self.prefix_trading_yearly))):
            (f_path, f_name) = os.path.split(filename)
            (f_short_name, f_extension) = os.path.splitext(f_name)

            # get last date of each stock 
            lastdate = self.get_last_date(f_short_name, 'trading')
            if lastdate == 'Year':
                last_date_of_file = '1999'  # dummy number, just crawling data
            else:
                last_date_of_file = lastdate
            # start crawling
            process_count += 1
            self.process(f_short_name, self.number_of_yearly_trading_files, process_count)
            self._get_trading_data(f_short_name)

        print ''

        # fetch yearly data of earning
        lastdate = self.get_last_date('1101', 'earning')
        if lastdate == 'Year':
            last_date_of_file = '1999'  # crawling data from 2000 (>1999)
        else:
            last_date_of_file = lastdate

        # start crawling
        print '### crawling yearly data --- EARNING ###'
        #self._get_earning_data('2017')
        for fetch_year_w in range(int(last_date_of_file)+1, int(year_w)+1):
            self._get_earning_data(str(fetch_year_w))
            print ''
def main():

    global last_date_of_file
    crawler = Crawler()

    # Set logging
    if not os.path.isdir('log'):
        os.makedirs('log')
    logging.basicConfig(filename='log/crawl-error.log',
        level=logging.ERROR,
        format='%(asctime)s\t[%(levelname)s]\t%(message)s',
        datefmt='%Y/%m/%d %H:%M:%S')

    # Get arguments
    parser = argparse.ArgumentParser(description='Crawl trading data for each year')
    
    #parser.add_argument('year', type=int, nargs='*',
    #    help='assigned year (format: YYYY), default is this year')
    parser.add_argument('-i', '--init', action='store_true',
        help='init table headers')
    parser.add_argument('-d', '--disable', action='store_true',
        help='disable checking new stock')
    parser.add_argument('-s', '--skip', action='store_true', 
        help='skip checking new stock')

    args = parser.parse_args()
    this_year = datetime.today()

    # make sure fetch day is not weekend
    weekend = this_year.weekday()
    if weekend < 5:
        date = this_year 
    else:
        date = this_year - timedelta(weekend % 4)

    # main function 
    if args.init:
        if not isdir(crawler.prefix_trading_yearly):
            mkdir(crawler.prefix_trading_yearly)
        if not isdir(crawler.prefix_earning_yearly):
            mkdir(crawler.prefix_earning_yearly)
        print 'TSE/OTC table init ...'
        crawler.table_init(date.year, date.month, date.day)
    else:
        if not (isdir(crawler.prefix_trading_yearly) and isdir(crawler.prefix_earning_yearly)):
            parser.error('Please do table-init first!')
            return

        # get the last date from '1101.csv'
        lastdate = crawler.get_last_date('1101', 'trading')
        #lastdate = crawler.get_last_date('1101', 'earning')
        if lastdate == 'Year':          # no data filled in table
            last_date_of_file = '1911'  # dummy number, just crawling data  
        else:
            last_date_of_file = lastdate

        if int(this_year.year) < (int(last_date_of_file) + 2):
            print 'Data is up-to-date!'
            return
        else:
            fetch_year = str(int(this_year.year)-1)
                    
        if (args.skip == False):
            print 'Checking is there a new stock...'
            crawler.check_new_stock(date.year, date.month, date.day)

        print ''
        print 'Crawling TSE/OTC yearly data to ' + fetch_year + ' ...'

        # start crawling data
        crawler.get_data(fetch_year)

        print '\r\rdone'
        print ''

        os.system('python -u create_yearly_chart.py')
if __name__ == '__main__':
    main()
