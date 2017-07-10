#coding=utf-8
import xlsxwriter
import openpyxl
import glob
import csv
import os

from os import mkdir
from os.path import isdir
from xlsxwriter.utility import xl_rowcol_to_cell

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
FOLDER_SAVE = 'tse_pressure_data'
RAW_DATA_FILE = 'TWSE_IDX'   #this is raw data file name => (RAW_DATA_FILE).csv
FILE_NAME = 'twse_momentum_table.xlsx'
TOTAL_DAYS = 270 

titles = [u'時間', u'開盤', u'最高', u'最低', u'收盤', u'', u'5MA', u'30MA', u'5MA-30MA']

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

def merge_data(worksheet):

    global RAW_DATA_FILE
    global TOTAL_DAYS

    last_row = 0
    csv_row_count = 0
    row_init = 1 
    csv_rows_tmp = []

    spamReader = csv.reader(open('{}.csv'.format(RAW_DATA_FILE), 'rb'), delimiter=',',quotechar='"')
    # read data to memory and filter out invalid data
    count = 0
    for row in spamReader:
        count += 1
        if count > 2:   #column 1 and 2 are headers, ignore it
            if (row[1] != '0') and (row[2] != '--'):
                csv_rows_tmp.append(row)
                last_row += 1

    if last_row < TOTAL_DAYS:
        last_row = TOTAL_DAYS

    for row in csv_rows_tmp:
        csv_row_count +=1
        if csv_row_count > (last_row - TOTAL_DAYS):
            try:
                new_open  = float(row[1])
                new_high  = float(row[2])
                new_low   = float(row[3])
                new_close = float(row[4])
            except ValueError:
                new_open  = row[1]
                new_high  = row[2]
                new_low   = row[3]
                new_close = row[4]
            pass    

            worksheet.write(row_init, 0, row[0])
            worksheet.write(row_init, 1, new_open)
            worksheet.write(row_init, 2, new_high)
            worksheet.write(row_init, 3, new_low)
            worksheet.write(row_init, 4, new_close)
            row_init +=1

            print "copy data @ " + "%10s...\r"%row[0],

    print ''

def process(name, total, count):
    print "create table [%6s]..."%(name),
    print "%3d"%(count*100/total) + "%\r",

def main():
    global titles
    global RAW_DATA_FILE

    print '***** Create TWSE Momentum Chart *****'
    WorkingDirectory = os.getcwd()

    #create stocks worksheet

#        print '%6s'%(f_short_name) + " creating table...",

    # create uniform worksheets
    spreadbook = xlsxwriter.Workbook('./{}'.format(FILE_NAME))
    spread_format=spreadbook.add_format()
    spread_format.set_border(1)
    spread_format.set_bg_color('#cccccc')
    spread_format.set_align('center')
    number_format=spreadbook.add_format()
    number_format.set_num_format('0.00')

    spreadsheet = spreadbook.add_worksheet(RAW_DATA_FILE)
    spreadsheet.write_row('A1', titles, spread_format)
    spreadsheet.set_column(0, 0, 15) # A1: set column width to 15 
    spreadsheet.set_column(1, 4, 11) # B ~ E: set column width to 11 
    spreadsheet.set_column(6, 8, 11, number_format) # G ~ I: set column width to 11 and number format #.##

    # formula
    formula(spreadsheet)

    # chart - xlswriter
    #chart_def(spreadbook, spreadsheet, RAW_DATA_FILE)

    # merge data
    merge_data(spreadsheet)

    spreadbook.close()

    # chart - openpyxl
    wb = openpyxl.load_workbook('./{}'.format(FILE_NAME))
    ws = wb[RAW_DATA_FILE]
    chart_openpyxl(ws)
    wb.save(FILE_NAME)
    
    print 'done \r',
if __name__ == '__main__':
    main()
