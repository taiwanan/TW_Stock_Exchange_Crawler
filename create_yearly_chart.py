#coding=utf-8
import xlsxwriter
import glob
import csv
import os

from os import mkdir
from os.path import isdir
from xlsxwriter.utility import xl_rowcol_to_cell
from datetime import datetime, timedelta
from dateutil.relativedelta import relativedelta
from dateutil.rrule import rrule, MONTHLY

WorkingDirectory = os.getcwd()
TRADING_YEAR_RAW_DATA_FOLDER = 'yearly_trading_raw_data'
EARNING_YEAR_RAW_DATA_FOLDER = 'yearly_earning_raw_data'
OUTPUT_CHART_FILE = 'yearly_chart.xlsx'
sheet_name = 'YEARLY'
TOTAL_FETCH_YEARS = 5 
TOTAL_MONTHS = 24

def merge_data(worksheet, f_name, date_start, date_end, start_row, pb_value):

    global TRADING_YEAR_RAW_DATA_FOLDER
    global TOTAL_MONTHS
    global y_axis_min

    this_year = datetime.today()

    start_to_fetch = False
    csv_rows_trading_tmp = []
    csv_rows_earning_tmp = []
    csv_rows_tmp = []   # for profit, eps and net-asset in this year
    date_start_orig = date_start

    (f_short_name, f_extension) = os.path.splitext(f_name)
    spamReader_trading = csv.reader(open('{}/{}'.format(TRADING_YEAR_RAW_DATA_FOLDER, f_name), 'rb'), delimiter=',',quotechar='"')
    spamReader_earning = csv.reader(open('{}/{}'.format(EARNING_YEAR_RAW_DATA_FOLDER, f_name), 'rb'), delimiter=',',quotechar='"')

    ### trading and earning data would have different period of year, read seperately!
    ## trading data 
    count = 0
    for row_trading in spamReader_trading:
        count += 1
        if count > 2:   #row 1 and 2 are headers, ignore it
            csv_rows_trading_tmp.append(row_trading)

    # write data to worksheet
    col_init = 2    # start to fill data at column 2
    date_check_valid = False
    for date_period in range(TOTAL_FETCH_YEARS):
        date_check = str(int(date_start) + date_period)
        for row_trading in csv_rows_trading_tmp:
            if (row_trading[0] == date_check):
                date_check_valid = True 
                # write data to worksheet
                # 最高
                if (row_trading[1] == ''):
                    worksheet.write(start_row, col_init, 'n/a')
                else:
                    worksheet.write_number(start_row, col_init, float(row_trading[1]))
                # 最低
                if row_trading[3] == '':
                    worksheet.write(start_row, TOTAL_FETCH_YEARS+col_init, 'n/a')
                else:
                    worksheet.write(start_row, TOTAL_FETCH_YEARS+col_init, float(row_trading[3]))
                # 收盤平均價
                if row_trading[5] == '':
                    worksheet.write(start_row, 2*TOTAL_FETCH_YEARS+col_init, 'n/a')
                else:
                    worksheet.write(start_row, 2*TOTAL_FETCH_YEARS+col_init, float(row_trading[5]))
 
                break

            else:
                date_check_valid = False

        if date_check_valid == False:
            worksheet.write(start_row, col_init, 'n/a')
            worksheet.write(start_row, TOTAL_FETCH_YEARS+col_init, 'n/a')
            worksheet.write(start_row, 2*TOTAL_FETCH_YEARS+col_init, 'n/a')

        col_init += 1

    ## earning data
    count = 0   # reset count
    for row_earning in spamReader_earning:
        count += 1
        if count > 2:   #row 1 and 2 are headers, ignore it
            csv_rows_earning_tmp.append(row_earning)

    # write data to worksheet
    col_init = 2    # start to fill data at column 2
    date_start = date_start_orig
    date_check_valid = False
    for date_period in range(TOTAL_FETCH_YEARS):
        date_check = str(int(date_start) + date_period)
        for row_earning in csv_rows_earning_tmp:
            if (row_earning[0] == date_check):
                date_check_valid = True 
                # write data to worksheet
                if (row_earning[1] != 'n/a'):
                    # EPS
                    if row_earning[10] == 'NA':
                        worksheet.write(start_row, 3*TOTAL_FETCH_YEARS+col_init, 'n/a')
                    else:
                        worksheet.write(start_row, 3*TOTAL_FETCH_YEARS+col_init, float(row_earning[10]))

                    # 淨值
                    if row_earning[13] == 'n/a':
                        worksheet.write(start_row, 4*TOTAL_FETCH_YEARS+col_init, 'n/a')
                    else:
                        worksheet.write(start_row, 4*TOTAL_FETCH_YEARS+col_init, float(row_earning[13]))
 
                    #col_init += 1
                else:
                    worksheet.write(start_row, 3*TOTAL_FETCH_YEARS+col_init, 'n/a')

                break

            else:
                date_check_valid = False

        if date_check_valid == False:
            worksheet.write(start_row, 3*TOTAL_FETCH_YEARS+col_init, 'n/a')
            worksheet.write(start_row, 4*TOTAL_FETCH_YEARS+col_init, 'n/a')
 
        col_init += 1

    ## profit, eps and net-asset in this year
    # write data to worksheet
    col_init = 2    # start to fill data at column 2
    quarter = 1
    valid_data = False
    valid_data_fetch = False


    # what year to fetch
    if this_year.month < 6: # Q1 earning data of this year has not ready
        fetch_year =  str(int(this_year.year)-1) # fetch last year
    else:
        fetch_year = str(int(this_year.year))

    # open earning data
    try:
        spamReader_Q = csv.reader(open('./tse_earning_raw_data/{}'.format(f_name), 'rb'), delimiter=',',quotechar='"')
        valid_data = True
    except:
        try:
            spamReader_Q = csv.reader(open('./otc_earning_raw_data/{}'.format(f_name), 'rb'), delimiter=',',quotechar='"')
            valid_data = True
        except:
            valid_data = False
            #print 'stock [{}] has no earning data, fill n/a'.format(f_name)

    if valid_data == True:
        count = 0   # reset count
        for row in spamReader_Q:
            count += 1
            if count > 2:   #row 1 and 2 are headers, ignore it
                csv_rows_tmp.append(row)

        while quarter <= 4:
            fetch_date = '{} Q{}'.format(fetch_year, quarter)
            for row in csv_rows_tmp:
                if (row[0] == fetch_date):
                    if row[1] == 'n/a' or row[1] == '0' or row[3] == 'n/a' or row[3] == '0':
                        worksheet.write(start_row, 5*TOTAL_FETCH_YEARS+col_init+quarter-1, 'n/a') # 當季毛利
                        worksheet.write(start_row, 5*TOTAL_FETCH_YEARS+col_init+4+quarter-1, row[-2]) # 當季eps
                    else:
                        worksheet.write(start_row, 5*TOTAL_FETCH_YEARS+col_init+quarter-1, float('%.2f'%(float(row[3])/float(row[1])*100))) # 當季毛利
                        worksheet.write(start_row, 5*TOTAL_FETCH_YEARS+col_init+4+quarter-1, row[-2]) # 當季eps
                    valid_data_fetch = True
                    break
                else:
                    valid_data_fetch = False

            if valid_data_fetch == False:
                worksheet.write(start_row, 5*TOTAL_FETCH_YEARS+col_init+quarter-1, 'n/a') # 當季毛利
                worksheet.write(start_row, 5*TOTAL_FETCH_YEARS+col_init+4+quarter-1, 'n/a') # 當季eps

            quarter += 1

    worksheet.write(start_row, 5*TOTAL_FETCH_YEARS+col_init+8, pb_value) # 當季淨值

def process(name, total, count):
    print "fetching data [%6s]..."%(name),
    print "%3d"%(count*100/total) + "%\r",

def get_date_list():

    date_list = []  # [ date_start ...  date_end]

    with open('./{}/1101.csv'.format(TRADING_YEAR_RAW_DATA_FOLDER), 'rb') as file:
        reader = csv.reader(file)
        rows = 0
        for row in reader:
            rows += 1
            if rows > 2:   # get date from 1st-data to the end
                date_list.append(row[0])
    file.close()

    return date_list

def main():
    global OUTPUT_CHART_FILE 
    global TOTAL_YEARS
    stock_count = 0
    total_years = 0
    pb_data = []

    print '***** Create TSE/OTC Yearly Chart *****'
    WorkingDirectory = os.getcwd()
    today = datetime.today()

    # get the last data date from stock '1101'
    date_list = get_date_list()
    # TSE/OTC total files
    total_files = sum(os.path.isfile(os.path.join('{}/'.format(TRADING_YEAR_RAW_DATA_FOLDER), f)) 
                    for f in os.listdir('{}/'.format(TRADING_YEAR_RAW_DATA_FOLDER)))

    ### create worksheet style for monthly revenue data
    # column lists
    row_headers = [ u'最高股價',u'最低股價',u'平均股價',u'EPS',u'每股淨值(PB)',u'當季毛利率',u'當季EPS']
    row_titles = [u'ID', u'Stock']
    row_titles_extra = ['Q1','Q2','Q3','Q4','Q1','Q2','Q3','Q4','PB']
    for columns in range(5):
        for date in date_list[-TOTAL_FETCH_YEARS:]:
            row_titles.append('%s'%(date))
    row_titles.extend(row_titles_extra)

    # create uniform worksheets
    spreadbook = xlsxwriter.Workbook('./{}'.format(OUTPUT_CHART_FILE))

    title_format = spreadbook.add_format({
    'bold':     True,
    'border':   1,
    'align':    'center',
    'valign':   'vcenter',
    'bg_color': '#cccccc',
    })

    id_format = spreadbook.add_format({
    'border':   1,
    'align':    'center',
    'valign':   'vcenter',
    'bg_color': '#cccccc',
    'text_wrap': True,
    })

    merge_format_1 = spreadbook.add_format({
    'border': 1,
    'align': 'center',
    'valign': 'vcenter',
    'text_wrap': True,
    #'bg_color': '#777777',
    })

    merge_format_1a= spreadbook.add_format({
    'border': 1,
    'align': 'center',
    'valign': 'vcenter',
    'text_wrap': True,
    #'bg_color': '#777777',
    'num_format': '0%',
    })

    merge_format_2 = spreadbook.add_format({
    'border': 1,
    'align': 'center',
    'valign': 'vcenter',
    'text_wrap': True,
    #'bg_color': '#cccccc',
    })

    merge_format_2a = spreadbook.add_format({
    'border': 1,
    'align': 'center',
    'valign': 'vcenter',
    'text_wrap': True,
    'bg_color': '#cccccc',
    'num_format': '0%',
    })

    hyperlink_format = spreadbook.add_format({
    'border': 1,
    'align': 'center',
    'valign': 'vcenter',
    'text_wrap': True,
    'bg_color': '#cccccc',
    'color': 'blue',
    })

    hyperlink_format_2 = spreadbook.add_format({
    'border': 1,
    'align': 'center',
    'valign': 'vcenter',
    'text_wrap': True,
    #'bg_color': '#cccccc',
    'color': 'blue',
    })

    number_format = spreadbook.add_format({'num_format': '0.00'})

    ## open worksheet
    yearly_spreadsheet = spreadbook.add_worksheet(sheet_name)
    yearly_spreadsheet.freeze_panes(2, 2)

    ## load PB data
    spamReader_PB = csv.reader(open('./yearly_tmp_data.csv', 'rb'), delimiter=',',quotechar='"')
    for row in spamReader_PB:
        pb_data.append(row)

    row_headers.append(pb_data[0][0])

    # header & title
    for t in range(5):
        yearly_spreadsheet.merge_range(0,2+TOTAL_FETCH_YEARS*t,0,1+TOTAL_FETCH_YEARS+TOTAL_FETCH_YEARS*t, row_headers[t], title_format)
    for t in range(2):
        yearly_spreadsheet.merge_range(0,(2+TOTAL_FETCH_YEARS*5)+4*t,0,(5+TOTAL_FETCH_YEARS*5)+4*t, row_headers[t+5], title_format)
    yearly_spreadsheet.write(0, 10+TOTAL_FETCH_YEARS*5, row_headers[-1])
    yearly_spreadsheet.set_row(0, 15, title_format)
    yearly_spreadsheet.write_row('A2', row_titles, title_format)
    yearly_spreadsheet.set_column(0, 1, 9) # A ~ B: set column width to 9 
    yearly_spreadsheet.set_column(2, len(date_list)+1, 11, number_format) # C ~ : set column width to 11 and number format #.##

    ### merge data
    for filename in sorted(glob.glob("./{}/*.csv".format(TRADING_YEAR_RAW_DATA_FOLDER))):
        (f_path, f_name) = os.path.split(filename)
        (f_short_name, f_extension) = os.path.splitext(f_name)

        with open('{}/{}.csv'.format(TRADING_YEAR_RAW_DATA_FOLDER, f_short_name), 'rb') as file:
            reader = csv.reader(file)
            for row in reader:
                stock_id = u'%s'%(f_short_name)
                stock_name = u'%s'%(row[0].decode('utf-8').strip(" "))
                stock_name_hyperlink = u"=HYPERLINK(\"https://tw.stock.yahoo.com/d/s/company_{}.html\",\"{}\")".format(f_short_name, stock_name)
                break

        # format: excel table style & an alternative color for next ID
        stock_count_row = 2 + stock_count 
        color_format = spreadbook.add_format()
        color_format.set_bg_color('#cccccc')
        color_format.set_border(1) 
        yearly_spreadsheet.set_row(stock_count_row, 15, merge_format_1)
        yearly_spreadsheet.write(stock_count_row,0,int(stock_id),id_format)
        yearly_spreadsheet.write(stock_count_row,1,stock_name_hyperlink, hyperlink_format)

        for row_data in pb_data:
            if row_data[1] == f_short_name:
                pb_value = row_data[2]
                break
            else:
                pb_value = 'n/a'

        # merge data
        merge_data(yearly_spreadsheet, f_name, date_list[-TOTAL_FETCH_YEARS], date_list[-1], stock_count_row, pb_value)

        # merge data
        stock_count += 1

        process(f_short_name, total_files, stock_count)

    print ''
    print 'export excel : %s \r'%(OUTPUT_CHART_FILE)

    spreadbook.close()

    print 'done!'

if __name__ == '__main__':
    main()
