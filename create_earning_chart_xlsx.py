#coding=utf-8
import xlsxwriter
import openpyxl
import glob
import csv
import os
import numpy as np

from os import mkdir
from os.path import isdir
from xlsxwriter.utility import xl_rowcol_to_cell
from datetime import datetime, timedelta

# use openpyxl to generate Stock Chart
from openpyxl.chart import (
    BarChart,
    StockChart,
    Reference,
    Series,
)
from openpyxl.chart.axis import ChartLines
from openpyxl.chart.updown_bars import UpDownBars
from copy import deepcopy


WorkingDirectory = os.getcwd()
TSE_RAW_DATA_FOLDER = 'tse_earning_raw_data'
OTC_RAW_DATA_FOLDER = 'otc_earning_raw_data'
OUTPUT_CHART_FILE = 'tse_otc_earning_chart.xlsx'
TOTAL_YEARS = 3

Q4 = np.array([0,0,0],np.int) 

column_titles = [u'營收', u'毛利', u'營利', u'EPS']

def formula(worksheet):

    global TOTAL_DAYS

    row_init = 0
    col_init = 0
    row_end = TOTAL_DAYS + 1 

    for row in range(5,row_end,1):
        #current row value
        col_B = xl_rowcol_to_cell(row, 1)
        col_C = xl_rowcol_to_cell(row, 2)
        col_D = xl_rowcol_to_cell(row, 3)
        col_E = xl_rowcol_to_cell(row, 4)
        col_G = xl_rowcol_to_cell(row, 6)
        col_H = xl_rowcol_to_cell(row, 7)
        col_I = xl_rowcol_to_cell(row, 8)

        # 5MA
        first_row = xl_rowcol_to_cell(row-4, 4) #col_E
        worksheet.write_formula(row, col_init+6, "AVERAGE(%s:%s)"%(first_row, col_E))

        if row >= 30:
            # 30MA
            first_row = xl_rowcol_to_cell(row-29, 4) #col_E
            worksheet.write_formula(row, col_init+7, "AVERAGE(%s:%s)"%(first_row, col_E))

            # 5-30MA
            worksheet.write_formula(row, col_init+8, "%s-%s"%(col_G, col_H))
 
def chart_def(spreadbook, worksheet, f_short_name):

    global TOTAL_DAYS

    chart = spreadbook.add_chart({'type': 'stock'})
    # Add a series for each of the High-Low-Close columns.
    chart.add_series({
        'categories': '=%s!$A$31:$A$%s'%(f_short_name,TOTAL_DAYS+1), # Date
        'values': '=%s!$B$31:$B$%s'%(f_short_name,TOTAL_DAYS+1),     # Open 
    })

    chart.add_series({
        'categories': '=%s!$A$31:$A$%s'%(f_short_name,TOTAL_DAYS+1), # Date
        'values': '=%s!$C$31:$C$%s'%(f_short_name,TOTAL_DAYS+1),     # High
    })

    chart.add_series({
        'categories': '=%s!$A$31:$A$%s'%(f_short_name,TOTAL_DAYS+1), # Date
        'values': '=%s!$D$31:$D$%s'%(f_short_name,TOTAL_DAYS+1), # Low
    })

    chart.add_series({
        'categories': '=%s!$A$31:$A$%s'%(f_short_name, TOTAL_DAYS+1),# Date
        'values': '=%s!$E$31:$E$%s'%(f_short_name,TOTAL_DAYS+1),     # Close
    })


    chart.set_title({'name': u'TWSE 波段動能分析'})
    chart.set_x_axis({'name': 'Date'})
    chart.set_y_axis({'name': 'Price'})
    chart.set_size({'width': 720, 'height': 480})
    chart.set_legend({'none': True})

    chart.set_up_down_bars({
        'down': {
            'fill':   {'color': '#00B050'},
            'border': {'color': 'black'}
        },
        'up': {
            'fill':   {'color': 'red'},
            'border': {'color': 'black'},
        },
    })

    # Create threshold-line as the secondary chart.
    bar_chart = spreadbook.add_chart({'type': 'column'})

    bar_chart.add_series({
#        'name':       '=pressure',
        'categories': '=%s!$A$31:$A$%s'%(f_short_name, TOTAL_DAYS+1),   # Date
        'values':     '=%s!$I$31:$I$%s'%(f_short_name,TOTAL_DAYS+1),    # 5-30MA
        'line':       {'color': '#9494e0'},
        'y2_axis':    True,
    })

    #line_chart.set_size({'width': 720, 'height': 360})

    chart.combine(bar_chart)
    worksheet.insert_chart('C%s'%(TOTAL_DAYS+2), chart) #display chart at specific location

def chart_openpyxl(worksheet):

    global TOTAL_DAYS

    print 'generate Chart...'
    # create open-high-low-close chart 
    stock = StockChart()
    dates = Reference(worksheet, min_col=1, min_row=31, max_row=TOTAL_DAYS+1)
    data = Reference(worksheet, min_col=2, max_col=5, min_row=31, max_row=TOTAL_DAYS+1)
    stock.add_data(data, titles_from_data= True)
    stock.set_categories(dates)
    for s in stock.series:
        s.graphicalProperties.line.noFill = True 
    stock.hiLowLines = ChartLines()
    stock.upDownBars = UpDownBars()

    # Excel is broken and needs a cache of values in order to display hiLoLines :-/
    from openpyxl.chart.data_source import NumData, NumVal
    pts = [NumVal(idx=i) for i in range(len(data) - 1)]
    cache = NumData(pt=pts)

    # add dummy cache
    stock.series[-1].val.numRef.numCache = cache

    # create 5-30MA chart
    bar = BarChart()
    data = Reference(worksheet, min_col=9, min_row=31, max_row=TOTAL_DAYS+1)
    bar.add_data(data, titles_from_data=False)
    bar.set_categories(dates)
    
    # merge K-Line & 5-30MA chart
    stock_tmp = deepcopy(bar)
    bar_tmp = deepcopy(stock)
    stock_tmp.title = "TWSE Mometum Chart"
    stock_tmp.height = 15
    stock_tmp.width = 25
    stock_tmp.y_axis.axId = 20 
    stock_tmp.z_axis = bar_tmp.y_axis
    stock_tmp.y_axis.crosses = "max"
    stock_tmp.legend = None                  # hidden series label name 
    bar_tmp.y_axis.majorGridlines = None
    bar_tmp.y_axis.title = "Price"
    stock_tmp += bar_tmp

    worksheet.add_chart(stock_tmp, "A{}".format(TOTAL_DAYS+2))

def merge_data(worksheet, f_name, start_year, start_row):

    global TSE_RAW_DATA_FOLDER
    global OTC_RAW_DATA_FOLDER
    global TOTAL_YEARS
    global Q4

    Q4 = 0
    col_init = 2  # start to fill data at column 2

    count = 0
    start_to_fetch = False
    csv_rows_tmp = []

    (f_short_name, f_extension) = os.path.splitext(f_name)
    if (worksheet.get_name() == 'TSE'):
        spamReader = csv.reader(open('{}/{}'.format(TSE_RAW_DATA_FOLDER, f_name), 'rb'), delimiter=',',quotechar='"')
    else:
        spamReader = csv.reader(open('{}/{}'.format(OTC_RAW_DATA_FOLDER, f_name), 'rb'), delimiter=',',quotechar='"')


    ## read data to memory and filter out invalid data
    for row in spamReader:
        count += 1
        if count > 2:   #row 1 and 2 are headers, ignore it
            csv_rows_tmp.append(row)

    ## write data to worksheet
    for row in csv_rows_tmp:
        if (row[0] == "%s Q1"%(start_year)) or (start_to_fetch == True):
            start_to_fetch = True 

            # RAW data for Q4 means a whole year, modify to Q4 = Q4 - (Q1 ~ Q3) 
            if (row[0] == "%s Q4"%(start_year)):
                if (row[1] != 'n/a'):
                    tmp = np.array([int(row[1]),int(row[3]),int(row[5])])
                    Q4 = tmp - Q4 
                    #print Q4
                    row[1] = Q4[0] # 營收
                    row[3] = Q4[1] # 毛利
                    row[5] = Q4[2] # 營益
                start_year +=1
                Q4 = 0
            else:
                if (row[1] == 'n/a'):
                    #tmp = np.zeros((3,), dtype=np.int)
                    tmp = 0
                else:
                    tmp = np.array([int(row[1]),int(row[3]),int(row[5])])
                Q4 = Q4 + tmp 

            # write data to worksheet
            if (row[1] != 'n/a'):
                worksheet.write((start_row<<2)+1, col_init, row[1]) # 營收
                worksheet.write((start_row<<2)+2, col_init, row[3]) # 毛利
                worksheet.write((start_row<<2)+3, col_init, row[5]) # 營益
                worksheet.write((start_row<<2)+4, col_init, '%1.2f'%(float(row[5])/(float(row[9])/10))) # 營益的EPS
                col_init +=1
            else:
                worksheet.write((start_row<<2)+1, col_init, 'n/a')
                worksheet.write((start_row<<2)+2, col_init, 'n/a')
                worksheet.write((start_row<<2)+3, col_init, 'n/a')
                worksheet.write((start_row<<2)+4, col_init, 'n/a')
                col_init +=1

def process(market, name, total, count):
    if(market == 'TSE'):
        print "merge TSE data [%6s]..."%(name),
    else:
        print "merge OTC data [%6s]..."%(name),

    print "%3d"%(count*100/total) + "%\r",

def main():
    global column_titles
    global OUTPUT_CHART_FILE 
    global TOTAL_YEARS
    stock_count  = 0

    print '***** Create TSE/OTC Earning Chart *****'
    WorkingDirectory = os.getcwd()
    today = datetime.today()

    # total files in both TSE/OTC
    tse_files = sum(os.path.isfile(os.path.join('{}/'.format(TSE_RAW_DATA_FOLDER), f)) 
                    for f in os.listdir('{}/'.format(TSE_RAW_DATA_FOLDER)))
    otc_files = sum(os.path.isfile(os.path.join('{}/'.format(OTC_RAW_DATA_FOLDER), f)) 
                    for f in os.listdir('{}/'.format(OTC_RAW_DATA_FOLDER)))

    ## create earning data worksheet
    # column lists
    row_titles = [u'Stock', u'Item']
    for year in range(TOTAL_YEARS-1,-1,-1):
        row_titles.append('%s Q1'%(today.year-year))
        row_titles.append('%s Q2'%(today.year-year))
        row_titles.append('%s Q3'%(today.year-year))
        row_titles.append('%s Q4'%(today.year-year))

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

    merge_format_2 = spreadbook.add_format({
    'border': 1,
    'align': 'center',
    'valign': 'vcenter',
    'text_wrap': True,
    'bg_color': '#cccccc',
    })

    number_format=spreadbook.add_format()
    number_format.set_num_format('0.00')

    tse_spreadsheet = spreadbook.add_worksheet('TSE')
    tse_spreadsheet.freeze_panes(1, 0)
    tse_spreadsheet.set_row(0, 15, title_format) # title format
    tse_spreadsheet.write_row('A1', row_titles, title_format)
    tse_spreadsheet.set_column(0, 1, 9) # A ~ B: set column width to 11 
    tse_spreadsheet.set_column(2, TOTAL_YEARS*4+1, 11, number_format) # C ~ : set column width to 11 and number format #.##

    otc_spreadsheet = spreadbook.add_worksheet('OTC')
    otc_spreadsheet.freeze_panes(1, 0)
    otc_spreadsheet.set_row(0, 15, title_format) # title format
    otc_spreadsheet.write_row('A1', row_titles, title_format)
    otc_spreadsheet.set_column(0, 1, 9) # A ~ B: set column width to 11 
    otc_spreadsheet.set_column(2, TOTAL_YEARS*4+1, 11, number_format) # C ~ : set column width to 11 and number format #.##

    # formula
    #formula(spreadsheet)

    # chart - xlswriter
    #chart_def(spreadbook, spreadsheet, RAW_DATA_FILE)

    ### merge data
    ## this is for TSE data merge
    for filename in glob.glob("./{}/*.csv".format(TSE_RAW_DATA_FOLDER)):
        (f_path, f_name) = os.path.split(filename)
        (f_short_name, f_extension) = os.path.splitext(f_name)

 
        with open('{}/{}.csv'.format(TSE_RAW_DATA_FOLDER, f_short_name), 'rb') as file:
            reader = csv.reader(file)
            for row in reader:
                stock_name = u'%s(%s)'%(row[0].decode('utf-8').strip(" "),f_short_name)
                break
        # alternative color for stock ID
        color_format = spreadbook.add_format()
        if (stock_count % 2) == 0:
            color_format.set_bg_color('#777777')
            color_format.set_border(1) 
            tse_spreadsheet.set_row((stock_count<<2)+1, 15, merge_format_1)  #
            tse_spreadsheet.set_row((stock_count<<2)+2, 15, merge_format_1)  #
            tse_spreadsheet.set_row((stock_count<<2)+3, 15, merge_format_1)  #
            tse_spreadsheet.set_row((stock_count<<2)+4, 15, merge_format_1)  #
            tse_spreadsheet.merge_range((stock_count<<2)+1,0,(stock_count<<2)+4,0,stock_name, merge_format_1)
        else:
            color_format.set_bg_color('#cccccc')
            color_format.set_border(1) 
            tse_spreadsheet.set_row((stock_count<<2)+1, 15, merge_format_2)  #
            tse_spreadsheet.set_row((stock_count<<2)+2, 15, merge_format_2)  #
            tse_spreadsheet.set_row((stock_count<<2)+3, 15, merge_format_2)  #
            tse_spreadsheet.set_row((stock_count<<2)+4, 15, merge_format_2)  #
            tse_spreadsheet.merge_range((stock_count<<2)+1,0,(stock_count<<2)+4,0,stock_name, merge_format_2)

        tse_spreadsheet.write((stock_count<<2)+1, 1, u"營收")
        tse_spreadsheet.write((stock_count<<2)+2, 1, u"毛利")
        tse_spreadsheet.write((stock_count<<2)+3, 1, u"營益")
        tse_spreadsheet.write((stock_count<<2)+4, 1, u"EPS")

        merge_data(tse_spreadsheet, f_name, today.year-TOTAL_YEARS+1, stock_count)
        stock_count += 1

        process('TSE', f_short_name, tse_files, stock_count)

    print ''

    ## this is for OTC data merge
    stock_count = 0
    for filename in glob.glob("./{}/*.csv".format(OTC_RAW_DATA_FOLDER)):
        (f_path, f_name) = os.path.split(filename)
        (f_short_name, f_extension) = os.path.splitext(f_name)

 
        with open('{}/{}.csv'.format(OTC_RAW_DATA_FOLDER, f_short_name), 'rb') as file:
            reader = csv.reader(file)
            for row in reader:
                stock_name = u'%s(%s)'%(row[0].decode('utf-8').strip(" "),f_short_name)
                break
        # alternative color for stock ID
        color_format = spreadbook.add_format()
        if (stock_count % 2) == 0:
            color_format.set_bg_color('#777777')
            color_format.set_border(1) 
            otc_spreadsheet.set_row((stock_count<<2)+1, 15, merge_format_1)  #
            otc_spreadsheet.set_row((stock_count<<2)+2, 15, merge_format_1)  #
            otc_spreadsheet.set_row((stock_count<<2)+3, 15, merge_format_1)  #
            otc_spreadsheet.set_row((stock_count<<2)+4, 15, merge_format_1)  #
            otc_spreadsheet.merge_range((stock_count<<2)+1,0,(stock_count<<2)+4,0,stock_name, merge_format_1)
        else:
            color_format.set_bg_color('#cccccc')
            color_format.set_border(1) 
            otc_spreadsheet.set_row((stock_count<<2)+1, 15, merge_format_2)  #
            otc_spreadsheet.set_row((stock_count<<2)+2, 15, merge_format_2)  #
            otc_spreadsheet.set_row((stock_count<<2)+3, 15, merge_format_2)  #
            otc_spreadsheet.set_row((stock_count<<2)+4, 15, merge_format_2)  #
            otc_spreadsheet.merge_range((stock_count<<2)+1,0,(stock_count<<2)+4,0,stock_name, merge_format_2)

        otc_spreadsheet.write((stock_count<<2)+1, 1, u"營收")
        otc_spreadsheet.write((stock_count<<2)+2, 1, u"毛利")
        otc_spreadsheet.write((stock_count<<2)+3, 1, u"營益")
        otc_spreadsheet.write((stock_count<<2)+4, 1, u"EPS")

        merge_data(otc_spreadsheet, f_name, today.year-TOTAL_YEARS+1, stock_count)
        stock_count += 1



        process('OTC', f_short_name, otc_files, stock_count)
  
    print ''

    spreadbook.close()

    # chart - openpyxl
    #wb = openpyxl.load_workbook('./{}'.format(FILE_NAME))
    #ws = wb[RAW_DATA_FILE]
    #chart_openpyxl(ws)
    #wb.save(FILE_NAME)
    
    print 'done \r',

if __name__ == '__main__':
    main()
