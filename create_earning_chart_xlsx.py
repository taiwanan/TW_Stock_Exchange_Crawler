#coding=utf-8
import xlsxwriter
import glob
import csv
import os
import numpy as np

from os import mkdir
from os.path import isdir
from xlsxwriter.utility import xl_rowcol_to_cell
from datetime import datetime, timedelta

WorkingDirectory = os.getcwd()
TSE_RAW_DATA_FOLDER = 'tse_earning_raw_data'
OTC_RAW_DATA_FOLDER = 'otc_earning_raw_data'
OUTPUT_CHART_FILE = 'tse_otc_earning_chart.xlsx'
TOTAL_YEARS = 4 

Q4 = np.array([0,0,0],np.int) 

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
 
def chart_trend(spreadbook, worksheet, start_year, stock_count_row):

    # chart
    line_chart = spreadbook.add_chart({'type': 'line'})
    bar_chart = spreadbook.add_chart({'type': 'column'})
    line_yoy_chart = spreadbook.add_chart({'type': 'line'})
    bar_yoy_p_chart = spreadbook.add_chart({'type': 'column'})
    bar_yoy_op_chart = spreadbook.add_chart({'type': 'column'})
    bar_yoy_eps_chart = spreadbook.add_chart({'type': 'column'})
    # row to present
    revenue    = stock_count_row + 2 # 營收
    profit     = stock_count_row + 3 # 毛利
    profit_m   = stock_count_row + 4 # 毛利率
    o_profit   = stock_count_row + 5 # 營益
    o_profit_m = stock_count_row + 6 # 營益率
    eps        = stock_count_row + 7 # eps

    last_col_chr = chr(ord('C')+(TOTAL_YEARS*4-1))  # character for the last column

    ## data item
    # 營收
    line_chart.add_series({
        'name':       '=%s!$B2'%('TSE'),
        'categories': '=%s!$C$1:$%c$1'%('TSE', last_col_chr),
        'values':     '=%s!$C$%s:$%c$%s'%('TSE',revenue, last_col_chr ,revenue),
        'marker':     {'type': 'diamond', 'size': 3},
    })
    # 毛利
    line_chart.add_series({
        'name':       '=%s!$B3'%('TSE'),
        'categories': '=%s!$C$1:$%c$1'%('TSE', last_col_chr),
        'values':     '=%s!$C$%s:$%c$%s'%('TSE',profit, last_col_chr, profit),
    })
    # 營益 
    line_chart.add_series({
        'name':       '=%s!$B5'%('TSE'),
        'categories': '=%s!$C$1:$%c$1'%('TSE', last_col_chr),
        'values':     '=%s!$C$%s:$%c$%s'%('TSE',o_profit, last_col_chr, o_profit),
    })

    # EPS
    bar_chart.add_series({
        'name':       '=%s!$B7'%('TSE'),
        'categories': '=%s!$C$1:$%c$1'%('TSE', last_col_chr),
        'values':     '=%s!$C$%s:$%c$%s'%('TSE',eps , last_col_chr, eps),
        'y2_axis':    True,
    })

    '''    
    #毛利率
    bar_chart.add_series({
        'name':       '=%s!$B4'%('TSE'),
        'categories': '=%s!$C$1:$%c$1'%('TSE', last_col_chr),
        'values':     '=%s!$C$%s:$%c$%s'%('TSE',profit_m, last_col_chr, profit_m),
        'y2_axis':    True,
    })
    #營益率
    bar_chart.add_series({
        'name':       '=%s!$B6'%('TSE'),
        'categories': '=%s!$C$1:$%c$1'%('TSE', last_col_chr),
        'values':     '=%s!$C$%s:$%c$%s'%('TSE',o_profit_m, last_col_chr, o_profit_m),
        'y2_axis':    True,
    })
    '''    

    for q in range(1,5):
        # 毛利率(YOY)
        bar_yoy_p_chart.add_series({
            'name':       '%s'%(start_year+q-1),
            'categories': '={"Q1","Q2","Q3","Q4"}',
            'values':     '=%s!$%c$%s:$%c$%s'%('TSE',chr(ord('C')+(q-1)*4), profit_m, chr(ord('C')+(q*4-1)), profit_m),
        })
        # 營益率(YOY)
        bar_yoy_op_chart.add_series({
            'name':       '%s'%(start_year+q-1),
            'categories': '={"Q1","Q2","Q3","Q4"}',
            'values':     '=%s!$%c$%s:$%c$%s'%('TSE',chr(ord('C')+(q-1)*4), o_profit_m, chr(ord('C')+(q*4-1)), o_profit_m),
        })
        # eps (YOY)
        bar_yoy_eps_chart.add_series({
            'name':       '%s'%(start_year+q-1),
            'categories': '={"Q1","Q2","Q3","Q4"}',
            'values':     '=%s!$%c$%s:$%c$%s'%('TSE',chr(ord('C')+(q-1)*4), eps, chr(ord('C')+(q*4-1)), eps),
        })


    ## chart characteristic 
    #line_chart.set_title({'name': u'損益表 '})
    line_chart.set_size({'width': 530, 'height': 145})
    line_chart.set_legend({'position': 'right'}) #圖例註釋
    line_chart.set_y_axis({ 
        'name': u'(百萬)',
        'display_units': 'thousands', 
        'display_units_visible': False, 
    })
    line_chart.set_x_axis({
       #'name': 'date',
        'interval_unit': 4,
    })
    bar_chart.set_style(13)
    bar_yoy_p_chart.set_y_axis({'num_format': '0%'})
    bar_yoy_p_chart.set_size({'width': 360, 'height': 145})
    bar_yoy_p_chart.set_title({'name': u'毛利率', 'overlay': True,})
    bar_yoy_op_chart.set_y_axis({'num_format': '0%'})
    bar_yoy_op_chart.set_size({'width': 360, 'height': 145})
    bar_yoy_op_chart.set_title({'name': u'營益率', 'overlay': True,})
    bar_yoy_eps_chart.set_y_axis({'num_format': '0.00'})
    bar_yoy_eps_chart.set_size({'width': 360, 'height': 145})
    bar_yoy_eps_chart.set_title({'name': 'EPS', 'overlay': True,})

    line_chart.combine(bar_chart)
    worksheet.insert_chart('T%s'%(stock_count_row+2), line_chart) #display chart at specific location
    worksheet.insert_chart('AC%s'%(stock_count_row+2), bar_yoy_p_chart) #display chart at specific location
    worksheet.insert_chart('AI%s'%(stock_count_row+2), bar_yoy_op_chart) #display chart at specific location
    worksheet.insert_chart('AO%s'%(stock_count_row+2), bar_yoy_eps_chart) #display chart at specific location

def merge_data(worksheet, worksheet_yoy, f_name, start_year, start_row):

    global TSE_RAW_DATA_FOLDER
    global OTC_RAW_DATA_FOLDER
    global TOTAL_YEARS
    global Q4

    Q4 = 0
    col_init = 2  # start to fill data at column 2
    col_init_yoy = 2  
    col_init_yoy_num = 2  

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
    count = 0   # reset count
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
                worksheet.write_number((start_row*6)+1, col_init, int(row[1])) # 營收
                worksheet.write_number((start_row*6)+2, col_init, int(row[3])) # 毛利
                if (row[1] == '0' or row[1] == 0):
                    worksheet.write_number((start_row*6)+3, col_init, 0) # 毛利率
                    worksheet.write_number((start_row*6)+5, col_init, 0) # 營益率
                else:
                    worksheet.write((start_row*6)+3, col_init, float(row[3])/float(row[1])) # 毛利率
                    worksheet.write((start_row*6)+5, col_init, float(row[5])/float(row[1])) # 營益率
                worksheet.write_number((start_row*6)+4, col_init, int(row[5])) # 營益
                worksheet.write_number((start_row*6)+6, col_init, float('%1.2f'%(float(row[5])/(float(row[9])/10)))) # 營益的EPS
                col_init +=1
            else:
                worksheet.write((start_row*6)+1, col_init, 'n/a')
                worksheet.write((start_row*6)+2, col_init, 'n/a')
                worksheet.write((start_row*6)+3, col_init, 'n/a')
                worksheet.write((start_row*6)+4, col_init, 'n/a')
                worksheet.write((start_row*6)+5, col_init, 'n/a')
                worksheet.write((start_row*6)+6, col_init, 'n/a')
                col_init +=1

            # write data to worksheet_yoy
            if (row[1] != 'n/a'):
                worksheet_yoy.write_number((start_row*6)+1, col_init_yoy_num, int(row[1])) # 營收
                worksheet_yoy.write_number((start_row*6)+2, col_init_yoy_num, int(row[3])) # 毛利
                if (row[1] == '0' or row[1] == 0):
                    worksheet_yoy.write_number((start_row*6)+3, col_init_yoy_num, 0) # 毛利率
                    worksheet_yoy.write_number((start_row*6)+5, col_init_yoy_num, 0) # 營益率
                else:
                    worksheet_yoy.write((start_row*6)+3, col_init_yoy_num, float(row[3])/float(row[1])) # 毛利率
                    worksheet_yoy.write((start_row*6)+5, col_init_yoy_num, float(row[5])/float(row[1])) # 營益率
                worksheet_yoy.write_number((start_row*6)+4, col_init_yoy_num, int(row[5])) # 營益
                worksheet_yoy.write((start_row*6)+6, col_init_yoy_num, float('%1.2f'%(float(row[5])/(float(row[9])/10)))) # 營益的EPS
                count+=1
                col_init_yoy_num = col_init_yoy + 4 * count 

            else:
                worksheet_yoy.write((start_row*6)+1, col_init_yoy_num, 'n/a')
                worksheet_yoy.write((start_row*6)+2, col_init_yoy_num, 'n/a')
                worksheet_yoy.write((start_row*6)+3, col_init_yoy_num, 'n/a')
                worksheet_yoy.write((start_row*6)+4, col_init_yoy_num, 'n/a')
                worksheet_yoy.write((start_row*6)+5, col_init_yoy_num, 'n/a')
                worksheet_yoy.write((start_row*6)+6, col_init_yoy_num, 'n/a')
                count+=1
                col_init_yoy_num = col_init_yoy + 4 * count 

            if (count >= TOTAL_YEARS):
                count = 0
                col_init_yoy+=1
                col_init_yoy_num = col_init_yoy + 4 * count 

def process(market, name, total, count):
    if(market == 'TSE'):
        print "merge TSE data [%6s]..."%(name),
    else:
        print "merge OTC data [%6s]..."%(name),

    print "%3d"%(count*100/total) + "%\r",

def main():
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

    ### create worksheet style for earning data
    # column lists
    row_titles = [u'Stock', u'Item']
    for year in range(TOTAL_YEARS-1,-1,-1):
        row_titles.append('%s Q1'%(today.year-year))
        row_titles.append('%s Q2'%(today.year-year))
        row_titles.append('%s Q3'%(today.year-year))
        row_titles.append('%s Q4'%(today.year-year))

    row_titles_yoy = [u'Stock', u'Item']
    for quarter in range(1,5):
        for year in range(TOTAL_YEARS-1,-1,-1):
            row_titles_yoy.append('%s Q%s'%(today.year-year, quarter))

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
    tse_spreadsheet = spreadbook.add_worksheet('TSE')
    tse_spreadsheet.freeze_panes(1, 0)
    tse_spreadsheet.set_row(0, 15, title_format) # title format
    tse_spreadsheet.write_row('A1', row_titles, title_format)
    tse_spreadsheet.set_column(0, 1, 9) # A ~ B: set column width to 9 
    tse_spreadsheet.set_column(2, TOTAL_YEARS*4+1, 11, number_format) # C ~ : set column width to 11 and number format #.##

    tse_yoy_spreadsheet = spreadbook.add_worksheet('TSE-YOY')
    tse_yoy_spreadsheet.freeze_panes(1, 0)
    tse_yoy_spreadsheet.set_row(0, 15, title_format) # title format
    tse_yoy_spreadsheet.write_row('A1', row_titles_yoy, title_format)
    tse_yoy_spreadsheet.set_column(0, 1, 9) # A ~ B: set column width to 9
    tse_yoy_spreadsheet.set_column(2, TOTAL_YEARS*4+1, 11, number_format) # C ~ : set column width to 11 and number format #.##

    otc_spreadsheet = spreadbook.add_worksheet('OTC')
    otc_spreadsheet.freeze_panes(1, 0)
    otc_spreadsheet.set_row(0, 15, title_format) # title format
    otc_spreadsheet.write_row('A1', row_titles, title_format)
    otc_spreadsheet.set_column(0, 1, 9) # A ~ B: set column width to 9
    otc_spreadsheet.set_column(2, TOTAL_YEARS*4+1, 11, number_format) # C ~ : set column width to 11 and number format #.##

    otc_yoy_spreadsheet = spreadbook.add_worksheet('OTC-YOY')
    otc_yoy_spreadsheet.freeze_panes(1, 0)
    otc_yoy_spreadsheet.set_row(0, 15, title_format) # title format
    otc_yoy_spreadsheet.write_row('A1', row_titles_yoy, title_format)
    otc_yoy_spreadsheet.set_column(0, 1, 9) # A ~ B: set column width to 9
    otc_yoy_spreadsheet.set_column(2, TOTAL_YEARS*4+1, 11, number_format) # C ~ : set column width to 11 and number format #.##

    ### merge data
    ## this is for TSE data
    for filename in sorted(glob.glob("./{}/*.csv".format(TSE_RAW_DATA_FOLDER))):
        (f_path, f_name) = os.path.split(filename)
        (f_short_name, f_extension) = os.path.splitext(f_name)

        with open('{}/{}.csv'.format(TSE_RAW_DATA_FOLDER, f_short_name), 'rb') as file:
            reader = csv.reader(file)
            for row in reader:
                stock_name = u'%s(%s)'%(row[0].decode('utf-8').strip(" "),f_short_name)
                hyperlink = u"=HYPERLINK(\"https://tw.stock.yahoo.com/d/s/company_{}.html\",\"{}\")".format(f_short_name, stock_name)
                break
        # format: excel table style & an alternative color for next ID
        stock_count_row = stock_count * 6
        color_format = spreadbook.add_format()
        if (stock_count % 2) == 0:
            color_format.set_bg_color('#777777')
            color_format.set_border(1) 
            # normal TSE
            tse_spreadsheet.set_row((stock_count_row)+1, 18, merge_format_1)  #
            tse_spreadsheet.set_row((stock_count_row)+2, 18, merge_format_1)  #
            tse_spreadsheet.set_row((stock_count_row)+3, 18, merge_format_1a) #
            tse_spreadsheet.set_row((stock_count_row)+4, 18, merge_format_1)  #
            tse_spreadsheet.set_row((stock_count_row)+5, 18, merge_format_1a) #
            tse_spreadsheet.set_row((stock_count_row)+6, 18, merge_format_1)  #
            tse_spreadsheet.merge_range((stock_count_row)+1,0,(stock_count_row)+6,0, hyperlink, hyperlink_format)
            # TSE for YOY 
            tse_yoy_spreadsheet.set_row((stock_count_row)+1, 18, merge_format_1)  #
            tse_yoy_spreadsheet.set_row((stock_count_row)+2, 18, merge_format_1)  #
            tse_yoy_spreadsheet.set_row((stock_count_row)+3, 18, merge_format_1a) #
            tse_yoy_spreadsheet.set_row((stock_count_row)+4, 18, merge_format_1)  #
            tse_yoy_spreadsheet.set_row((stock_count_row)+5, 18, merge_format_1a) #
            tse_yoy_spreadsheet.set_row((stock_count_row)+6, 18, merge_format_1)  #
            tse_yoy_spreadsheet.merge_range((stock_count_row)+1,0,(stock_count_row)+6,0, hyperlink, hyperlink_format)
        else:
            color_format.set_bg_color('#cccccc')
            color_format.set_border(1) 
            # normal TSE
            tse_spreadsheet.set_row((stock_count_row)+1, 18, merge_format_2)  #
            tse_spreadsheet.set_row((stock_count_row)+2, 18, merge_format_2)  #
            tse_spreadsheet.set_row((stock_count_row)+3, 18, merge_format_2a) #
            tse_spreadsheet.set_row((stock_count_row)+4, 18, merge_format_2)  #
            tse_spreadsheet.set_row((stock_count_row)+5, 18, merge_format_2a) #
            tse_spreadsheet.set_row((stock_count_row)+6, 18, merge_format_2)  #
            tse_spreadsheet.merge_range((stock_count_row)+1,0,(stock_count_row)+6,0, hyperlink, hyperlink_format_2)
            # TSE for YOY
            tse_yoy_spreadsheet.set_row((stock_count_row)+1, 18, merge_format_2)  #
            tse_yoy_spreadsheet.set_row((stock_count_row)+2, 18, merge_format_2)  #
            tse_yoy_spreadsheet.set_row((stock_count_row)+3, 18, merge_format_2a) #
            tse_yoy_spreadsheet.set_row((stock_count_row)+4, 18, merge_format_2)  #
            tse_yoy_spreadsheet.set_row((stock_count_row)+5, 18, merge_format_2a) #
            tse_yoy_spreadsheet.set_row((stock_count_row)+6, 18, merge_format_2)  #
            tse_yoy_spreadsheet.merge_range((stock_count_row)+1,0,(stock_count_row)+6,0, hyperlink, hyperlink_format_2)

        tse_spreadsheet.write((stock_count_row)+1, 1, u"營收")
        tse_spreadsheet.write((stock_count_row)+2, 1, u"毛利")
        tse_spreadsheet.write((stock_count_row)+3, 1, u"毛利率")
        tse_spreadsheet.write((stock_count_row)+4, 1, u"營益")
        tse_spreadsheet.write((stock_count_row)+5, 1, u"營益率")
        tse_spreadsheet.write((stock_count_row)+6, 1, u"EPS")
        tse_yoy_spreadsheet.write((stock_count_row)+1, 1, u"營收")
        tse_yoy_spreadsheet.write((stock_count_row)+2, 1, u"毛利")
        tse_yoy_spreadsheet.write((stock_count_row)+3, 1, u"毛利率")
        tse_yoy_spreadsheet.write((stock_count_row)+4, 1, u"營益")
        tse_yoy_spreadsheet.write((stock_count_row)+5, 1, u"營益率")
        tse_yoy_spreadsheet.write((stock_count_row)+6, 1, u"EPS")

        # chart 
        chart_trend(spreadbook, tse_spreadsheet, today.year-TOTAL_YEARS+1, stock_count_row)

        # merge data
        merge_data(tse_spreadsheet, tse_yoy_spreadsheet, f_name, today.year-TOTAL_YEARS+1, stock_count)
        stock_count += 1

        process('TSE', f_short_name, tse_files, stock_count)

    print ''

    ## this is for OTC data
    stock_count = 0
    for filename in sorted(glob.glob("./{}/*.csv".format(OTC_RAW_DATA_FOLDER))):
        (f_path, f_name) = os.path.split(filename)
        (f_short_name, f_extension) = os.path.splitext(f_name)
 
        with open('{}/{}.csv'.format(OTC_RAW_DATA_FOLDER, f_short_name), 'rb') as file:
            reader = csv.reader(file)
            for row in reader:
                stock_name = u'%s(%s)'%(row[0].decode('utf-8').strip(" "),f_short_name)
                hyperlink = u"=HYPERLINK(\"https://tw.stock.yahoo.com/d/s/company_{}.html\",\"{}\")".format(f_short_name, stock_name)
                break
        # alternative color for stock ID
        stock_count_row = stock_count * 6
        color_format = spreadbook.add_format()
        if (stock_count % 2) == 0:
            color_format.set_bg_color('#777777')
            color_format.set_border(1) 
            otc_spreadsheet.set_row((stock_count_row)+1, 18, merge_format_1)  #
            otc_spreadsheet.set_row((stock_count_row)+2, 18, merge_format_1)  #
            otc_spreadsheet.set_row((stock_count_row)+3, 18, merge_format_1a) #
            otc_spreadsheet.set_row((stock_count_row)+4, 18, merge_format_1)  #
            otc_spreadsheet.set_row((stock_count_row)+5, 18, merge_format_1a) #
            otc_spreadsheet.set_row((stock_count_row)+6, 18, merge_format_1)  #
            otc_spreadsheet.merge_range((stock_count_row)+1,0,(stock_count_row)+6,0, hyperlink, hyperlink_format)
            # OTC for YOY 
            otc_yoy_spreadsheet.set_row((stock_count_row)+1, 18, merge_format_1)  #
            otc_yoy_spreadsheet.set_row((stock_count_row)+2, 18, merge_format_1)  #
            otc_yoy_spreadsheet.set_row((stock_count_row)+3, 18, merge_format_1a) #
            otc_yoy_spreadsheet.set_row((stock_count_row)+4, 18, merge_format_1)  #
            otc_yoy_spreadsheet.set_row((stock_count_row)+5, 18, merge_format_1a) #
            otc_yoy_spreadsheet.set_row((stock_count_row)+6, 18, merge_format_1)  #
            otc_yoy_spreadsheet.merge_range((stock_count_row)+1,0,(stock_count_row)+6,0, hyperlink, hyperlink_format)
        else:
            color_format.set_bg_color('#cccccc')
            color_format.set_border(1) 
            otc_spreadsheet.set_row((stock_count_row)+1, 18, merge_format_2)  #
            otc_spreadsheet.set_row((stock_count_row)+2, 18, merge_format_2)  #
            otc_spreadsheet.set_row((stock_count_row)+3, 18, merge_format_2a) #
            otc_spreadsheet.set_row((stock_count_row)+4, 18, merge_format_2)  #
            otc_spreadsheet.set_row((stock_count_row)+5, 18, merge_format_2a) #
            otc_spreadsheet.set_row((stock_count_row)+6, 18, merge_format_2)  #
            otc_spreadsheet.merge_range((stock_count_row)+1,0,(stock_count_row)+6,0, hyperlink, hyperlink_format_2)
            # OTC for YOY 
            otc_yoy_spreadsheet.set_row((stock_count_row)+1, 18, merge_format_2)  #
            otc_yoy_spreadsheet.set_row((stock_count_row)+2, 18, merge_format_2)  #
            otc_yoy_spreadsheet.set_row((stock_count_row)+3, 18, merge_format_2a) #
            otc_yoy_spreadsheet.set_row((stock_count_row)+4, 18, merge_format_2)  #
            otc_yoy_spreadsheet.set_row((stock_count_row)+5, 18, merge_format_2a) #
            otc_yoy_spreadsheet.set_row((stock_count_row)+6, 18, merge_format_2)  #
            otc_yoy_spreadsheet.merge_range((stock_count_row)+1,0,(stock_count_row)+6,0, hyperlink, hyperlink_format_2)

        otc_spreadsheet.write((stock_count_row)+1, 1, u"營收")
        otc_spreadsheet.write((stock_count_row)+2, 1, u"毛利")
        otc_spreadsheet.write((stock_count_row)+3, 1, u"毛利率")
        otc_spreadsheet.write((stock_count_row)+4, 1, u"營益")
        otc_spreadsheet.write((stock_count_row)+5, 1, u"營益率")
        otc_spreadsheet.write((stock_count_row)+6, 1, u"EPS")
        otc_yoy_spreadsheet.write((stock_count_row)+1, 1, u"營收")
        otc_yoy_spreadsheet.write((stock_count_row)+2, 1, u"毛利")
        otc_yoy_spreadsheet.write((stock_count_row)+3, 1, u"毛利率")
        otc_yoy_spreadsheet.write((stock_count_row)+4, 1, u"營益")
        otc_yoy_spreadsheet.write((stock_count_row)+5, 1, u"營益率")
        otc_yoy_spreadsheet.write((stock_count_row)+6, 1, u"EPS")

        # chart 
        chart_trend(spreadbook, otc_spreadsheet, today.year-TOTAL_YEARS+1, stock_count_row)

        # merge data
        merge_data(otc_spreadsheet, otc_yoy_spreadsheet, f_name, today.year-TOTAL_YEARS+1, stock_count)
        stock_count += 1

        process('OTC', f_short_name, otc_files, stock_count)
  
    print ''
    print 'export excel : %s \r'%(OUTPUT_CHART_FILE)

    spreadbook.close()

    print 'done \r',

if __name__ == '__main__':
    main()
