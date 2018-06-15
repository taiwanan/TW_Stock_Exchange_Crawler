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
MONTHLY_RAW_DATA_FOLDER = 'monthly_raw_data'
OUTPUT_CHART_FILE = 'tse_otc_monthly_chart.xlsx'
sheet_name = 'MONTHLY-REVENUE'
TOTAL_MONTHS = 24
MA = 4 # mean average
y_axis_min = ['none', 0]

def chart_trend(spreadbook, worksheet, stock_count_row):

    global y_axis_min

    # chart
    line_chart = spreadbook.add_chart({'type': 'line'})
    bar_chart = spreadbook.add_chart({'type': 'column'})
    # row to present
    revenue    = stock_count_row + 2 # 月營收 
    revenue_mom= stock_count_row + 3 # MOM
    revenue_yoy= stock_count_row + 4 # YOY
    ma         = stock_count_row + 5 # xMA
    ma_momentum= stock_count_row + 6 # MA-xMA

    last_col_chr = chr(ord('C')+(TOTAL_MONTHS-1))  # character for the last column

    ## data item
    # 月營收
    line_chart.add_series({
        'name':       '=%s!$B2'%(sheet_name),
        'categories': '=%s!$C$1:$%c$1'%(sheet_name, last_col_chr),
        'values':     '=%s!$C$%s:$%c$%s'%(sheet_name,revenue, last_col_chr ,revenue),
        'line':       {'color': 'red'},
        'marker':     {'type': 'circle', 'size': 2},
    })

    # xMA 
    line_chart.add_series({
        'name':       '=%s!$B5'%(sheet_name),
        'categories': '=%s!$C$1:$%c$1'%(sheet_name, last_col_chr),
        'values':     '=%s!$C$%s:$%c$%s'%(sheet_name, ma , last_col_chr, ma),
        'line':       {'color': 'blue'},
        #'y2_axis':    True,
    })

    # MA-xMA
    bar_chart.add_series({
        'name':       '=%s!$B6'%(sheet_name),
        'categories': '=%s!$C$1:$%c$1'%(sheet_name, last_col_chr),
        'values':     '=%s!$C$%s:$%c$%s'%(sheet_name, ma_momentum , last_col_chr, ma_momentum),
        'y2_axis':    True,
    })
    
    ## chart characteristic 
    #line_chart.set_title({'name': u'月營收 '})
    line_chart.set_size({'width': 510, 'height': 140})
    line_chart.set_legend({'position': 'right'}) #圖例註釋
    line_chart.set_y_axis({ 
        'name': u'(千)',
        #'display_units': 'thousands', 
        'display_units_visible': False, 
        'min': y_axis_min[1]
    })
    line_chart.set_x_axis({
       #'name': 'date',
        'minor_unit_type': 'months',
        'interval_unit': 6,
    })
    #bar_chart.set_style(13)
    #bar_chart.set_y_axis({'num_format': '0%'})

    line_chart.combine(bar_chart)
    worksheet.insert_chart(stock_count_row+1, TOTAL_MONTHS+3, line_chart) #display chart at specific location

def merge_data(worksheet, f_name, date_start, start_row):

    global MONTHLY_RAW_DATA_FOLDER
    global TOTAL_MONTHS
    global y_axis_min

    col_init = 2    # start to fill data at column 2

    count = 0
    y_axis_min = ['none', 0]
    start_to_fetch = False
    csv_rows_tmp = []
    ma_tmp = [0] * MA 
    ma_offset = 0

    (f_short_name, f_extension) = os.path.splitext(f_name)
    spamReader = csv.reader(open('{}/{}'.format(MONTHLY_RAW_DATA_FOLDER, f_name), 'rb'), delimiter=',',quotechar='"')


    ## read data to memory and filter out invalid data
    for row in spamReader:
        count += 1
        if count > 2:   #row 1 and 2 are headers, ignore it
            csv_rows_tmp.append(row)

    ## write data to worksheet
    count = 0   # reset count
    for row in csv_rows_tmp:

        # special case : if date of raw data after date_start, re-define date_start
        if (row[0] > '{}/{:02d}'.format(date_start.year, date_start.month)) and (start_to_fetch == False):
            row_date = datetime(int(row[0][:4]), int(row[0][-2:]), 15)
            r = relativedelta(row_date, date_start)
            date_start = date_start + relativedelta(years=r.years, months = r.months)
            ma_offset = r.years * 12 + r.months
            col_init += ma_offset 

        # normal case
        if (row[0] == '{}/{:02d}'.format(date_start.year, date_start.month)) or (start_to_fetch == True):
            start_to_fetch = True 

            # write data to worksheet
            if (row[1] != 'n/a'):
                worksheet.write_number((start_row*5)+1, col_init, int(row[1])) # 當月營收
                if y_axis_min[0] == 'none':
                    y_axis_min = ['true', int(row[1])]
                elif y_axis_min[0] == 'true' and y_axis_min[1] > int(row[1]):
                    y_axis_min[1] = int(row[1])

                # MOM
                if row[4] == '':
                    worksheet.write((start_row*5)+2, col_init, 'n/a')
                else:
                    worksheet.write_number((start_row*5)+2, col_init, float(row[4]))
                # YOY
                if row[5] == '':
                    worksheet.write((start_row*5)+3, col_init, 'n/a')
                else:
                    worksheet.write_number((start_row*5)+3, col_init, float(row[5]))
                # xMA 
                if ((col_init-2) >= MA+ma_offset-1):
                    ma_tmp[(col_init-2)%MA] = int(row[1])
                    ma_average = sum(ma_tmp)/len(ma_tmp)
                    worksheet.write_number((start_row*5)+4, col_init, ma_average ) # MA
                    worksheet.write_number((start_row*5)+5, col_init, int(row[1])-int(ma_average)) # MA - xMA

                ma_tmp[(col_init-2)%MA] = int(row[1])
                col_init +=1
            else:
                worksheet.write((start_row*5)+1, col_init, 'n/a')
                worksheet.write((start_row*5)+2, col_init, 'n/a')
                worksheet.write((start_row*5)+3, col_init, 'n/a')
                worksheet.write((start_row*5)+4, col_init, 'n/a')
                ma_tmp[(col_init-2)%MA] = int(0)

def process(name, total, count):
    print "fetching data [%6s]..."%(name),
    print "%3d"%(count*100/total) + "%\r",

def get_last_date():
    with open('./monthly_raw_data/1101.csv', 'rb') as file:
        reader = csv.reader(file)
        rows = 0
        for row in reader:
            rows += 1
    file.close()
    return row[0]

def main():
    global OUTPUT_CHART_FILE 
    global TOTAL_YEARS
    stock_count  = 0

    print '***** Create TSE/OTC Monthly-Revenue Chart *****'
    WorkingDirectory = os.getcwd()
    today = datetime.today()

    # get the last data date from stock '1101'
    lastdate = get_last_date()
    lastyear = lastdate.split('/')[0]
    lastmonth = lastdate.split('/')[1]
    datetime_end = datetime(int(lastyear), int(lastmonth), 15)
    datetime_start = datetime_end - relativedelta(months=TOTAL_MONTHS-1)

    # TSE/OTC total files
    total_files = sum(os.path.isfile(os.path.join('{}/'.format(MONTHLY_RAW_DATA_FOLDER), f)) 
                    for f in os.listdir('{}/'.format(MONTHLY_RAW_DATA_FOLDER)))

    ### create worksheet style for monthly revenue data
    # column lists
    row_titles = [u'Stock', u'Item']
    for date in rrule(MONTHLY, dtstart=datetime_start,until=datetime_end):
        row_titles.append('%s/%s'%(date.year, date.month))

    # create uniform worksheets
    spreadbook = xlsxwriter.Workbook('./{}'.format(OUTPUT_CHART_FILE))

    title_format = spreadbook.add_format({
    'bold':     True,
    'border':   1,
    'align':    'center',
    'valign':   'vcenter',
    'bg_color': '#cccccc',
    })

    merge_format_1 = spreadbook.add_format({
    'border': 1,
    'align': 'center',
    'valign': 'vcenter',
    'text_wrap': True,
    'bg_color': '#777777',
    })

    merge_format_1a= spreadbook.add_format({
    'border': 1,
    'align': 'center',
    'valign': 'vcenter',
    'text_wrap': True,
    'bg_color': '#777777',
    'num_format': '0%',
    })

    merge_format_2 = spreadbook.add_format({
    'border': 1,
    'align': 'center',
    'valign': 'vcenter',
    'text_wrap': True,
    'bg_color': '#cccccc',
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
    'bg_color': '#777777',
    'color': 'navy',
    })

    hyperlink_format_2 = spreadbook.add_format({
    'border': 1,
    'align': 'center',
    'valign': 'vcenter',
    'text_wrap': True,
    'bg_color': '#cccccc',
    'color': 'blue',
    })

    number_format = spreadbook.add_format({'num_format': '0.00'})

    # open worksheet
    monthly_spreadsheet = spreadbook.add_worksheet(sheet_name)
    monthly_spreadsheet.freeze_panes(1, 2)
    monthly_spreadsheet.set_row(0, 15, title_format) # title format
    monthly_spreadsheet.write_row('A1', row_titles, title_format)
    monthly_spreadsheet.set_column(0, 1, 9) # A ~ B: set column width to 9 
    monthly_spreadsheet.set_column(2, TOTAL_MONTHS+1, 11, number_format) # C ~ : set column width to 11 and number format #.##

    ### merge data
    for filename in sorted(glob.glob("./{}/*.csv".format(MONTHLY_RAW_DATA_FOLDER))):
        (f_path, f_name) = os.path.split(filename)
        (f_short_name, f_extension) = os.path.splitext(f_name)

        with open('{}/{}.csv'.format(MONTHLY_RAW_DATA_FOLDER, f_short_name), 'rb') as file:
            reader = csv.reader(file)
            for row in reader:
                stock_name = u'%s(%s)'%(row[0].decode('utf-8').strip(" "),f_short_name)
                hyperlink = u"=HYPERLINK(\"https://tw.stock.yahoo.com/d/s/company_{}.html\",\"{}\")".format(f_short_name, stock_name)
                break
        # format: excel table style & an alternative color for next ID
        stock_count_row = stock_count *5
        color_format = spreadbook.add_format()
        if (stock_count % 2) == 0:
            color_format.set_bg_color('#777777')
            color_format.set_border(1) 
            monthly_spreadsheet.set_row((stock_count_row)+1, 21, merge_format_1)  #
            monthly_spreadsheet.set_row((stock_count_row)+2, 21, merge_format_1)  # 
            monthly_spreadsheet.set_row((stock_count_row)+3, 21, merge_format_1)  #
            monthly_spreadsheet.set_row((stock_count_row)+4, 21, merge_format_1)  #
            monthly_spreadsheet.set_row((stock_count_row)+5, 21, merge_format_1)  #
            monthly_spreadsheet.merge_range((stock_count_row)+1,0,(stock_count_row)+5,0, hyperlink, hyperlink_format)
        else:
            color_format.set_bg_color('#cccccc')
            color_format.set_border(1) 
            monthly_spreadsheet.set_row((stock_count_row)+1, 21, merge_format_2)  #
            monthly_spreadsheet.set_row((stock_count_row)+2, 21, merge_format_2)  #
            monthly_spreadsheet.set_row((stock_count_row)+3, 21, merge_format_2)  #
            monthly_spreadsheet.set_row((stock_count_row)+4, 21, merge_format_2)  #
            monthly_spreadsheet.set_row((stock_count_row)+5, 21, merge_format_2)  #
            monthly_spreadsheet.merge_range((stock_count_row)+1,0,(stock_count_row)+5,0, hyperlink, hyperlink_format_2)

        monthly_spreadsheet.write((stock_count_row)+1, 1, u"營收(千)")
        monthly_spreadsheet.write((stock_count_row)+2, 1, u"MOM(%)")
        monthly_spreadsheet.write((stock_count_row)+3, 1, u"YOY(%)")
        monthly_spreadsheet.write((stock_count_row)+4, 1, u"{}MA".format(MA))
        monthly_spreadsheet.write((stock_count_row)+5, 1, u"MA-{}MA".format(MA))

        # merge data
        merge_data(monthly_spreadsheet, f_name, datetime_start, stock_count)

        # chart 
        chart_trend(spreadbook, monthly_spreadsheet, stock_count_row)

        # merge data
        stock_count += 1

        process(f_short_name, total_files, stock_count)

    print ''
    print 'export excel : %s \r'%(OUTPUT_CHART_FILE)

    spreadbook.close()

    print 'done!'

if __name__ == '__main__':
    main()
