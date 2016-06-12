import xlsxwriter
import os
import datetime
import json
import logging
from decimal import Decimal

def write_to_xlsx(filename, items, qty_dict, metadata) :
    '''
    writes the file to xlsx using specified template and returns the directory where
    the file is stored
    '''

    filename = metadata['filename_prefix'] + metadata['auction'] + '_' + str(metadata['date']) + '.xlsx'
    output_dir = os.path.dirname(os.path.dirname(os.path.abspath(__file__))) + "/output/"
    output_path = output_dir + filename

    date_of_report = datetime.datetime.strptime(str(metadata['date']), "%Y%m%d")
    date_in_full = date_of_report.strftime('%d %B %Y')

    workbook = xlsxwriter.Workbook(output_path)
    worksheet = workbook.add_worksheet()

    # data row starts at 
    row = 26

    # blue
    lineColor = '#367DA2'

    # bold formats
    bold = workbook.add_format()
    bold = workbook.add_format({'bold': True})

    # Title Format and Value
    titleFormat = workbook.add_format() 

    titleFormat.set_font_name('Helvetica Neue UltraLight')
    titleFormat.set_font_size(28)
    titleFormat.set_bottom(2)
    titleFormat.set_bottom_color(lineColor)

    # merge cells for title
    worksheet.merge_range('A1:M3', 'Merged Range', titleFormat)
    worksheet.write('A1', 'Daily Auction Analysis', titleFormat)

    # Description Format and Value
    descriptionFormat = workbook.add_format()

    descriptionFormat.set_font_name('Helvetica Neue Light')
    descriptionFormat.set_font_size(10)
    descriptionFormat.set_text_wrap(True)
    descriptionFormat.set_align('top')

    timePeriod = '15:10 and 15:15'
    if (metadata['period'] == 'Open') :
        timePeriod = '08:00 and 09:00'

    symbol = 'NK225/15H.OS'
    if (metadata['symbol'] == 'Nikkei Mini') :
        symbol = 'NK225M/15H.OS'
    if (metadata['symbol'] == 'Topix') :
        symbol = 'TOPIX/15H.OS'

    worksheet.merge_range('A4:F7', 'Merged Range', descriptionFormat)
    worksheet.write_rich_string('A4', '\nThis analysis is for ', bold, date_in_full, ' orders of ', bold, symbol + ' that enter the auction period between ', bold, timePeriod + ' with size of ' + str(metadata['qty']) + ' lots and bigger', descriptionFormat)

    # DateSymbolTitle Format and Value
    dateSymbolTitleFormat = workbook.add_format()

    dateSymbolTitleFormat.set_font_name('Helvetica Neue')
    dateSymbolTitleFormat.set_bold(True)
    dateSymbolTitleFormat.set_font_size(10)
    dateSymbolTitleFormat.set_text_wrap(True)
    dateSymbolTitleFormat.set_align('top')
    dateSymbolTitleFormat.set_top(1)
    dateSymbolTitleFormat.set_right(1)
    dateSymbolTitleFormat.set_bottom(1)
    dateSymbolTitleFormat.set_top_color(lineColor)
    dateSymbolTitleFormat.set_right_color(lineColor)
    dateSymbolTitleFormat.set_bottom_color(lineColor)

    worksheet.set_row(4, 25)
    worksheet.set_row(5, 25)
    worksheet.set_row(6, 25)

    worksheet.merge_range('I5:J5', 'Merged Range', dateSymbolTitleFormat)
    worksheet.merge_range('I6:J6', 'Merged Range', dateSymbolTitleFormat)
    worksheet.merge_range('I7:J7', 'Merged Range', dateSymbolTitleFormat)

    worksheet.write('I5', 'Date:', dateSymbolTitleFormat)
    worksheet.write('I6', 'Symbol:', dateSymbolTitleFormat)
    worksheet.write('I7', 'Min lot size:', dateSymbolTitleFormat)

    fillColor = '#FDF7E4'

    # DateSymbolValue Format and Value
    dateSymbolValueFormatProps = {
        'font_name': 'Helvetica Neue Light',
        'bold': True,
        'font_size': 10,
        'text_wrap': True,
        'valign': 'top',
        'align': 'right',
        'top': 1,
        'left': 1,
        'bottom': 1, 
        'bg_color': fillColor,
        'top_color': lineColor,
        'left_color': lineColor,
        'bottom_color': lineColor
    }

    dateValueFormatProps = dict.copy(dateSymbolValueFormatProps)
    dateValueFormatProps['num_format'] = 'd mmm yy'

    minLotValueFormatProps = dict.copy(dateSymbolValueFormatProps)
    minLotValueFormatProps['num_format'] = '#,###'

    dateSymbolValueFormat = workbook.add_format(dateSymbolValueFormatProps)
    dateValueFormat = workbook.add_format(dateValueFormatProps)
    minLotValueFormat = workbook.add_format(minLotValueFormatProps)

    worksheet.merge_range('K5:M5', date_of_report, dateValueFormat)
    worksheet.merge_range('K6:M6', symbol, dateSymbolValueFormat)
    worksheet.merge_range('K7:M7', metadata['qty'], minLotValueFormat)

    summaryTitleFormat = workbook.add_format()

    summaryTitleFormat.set_font_name('Helvetica Neue Medium')
    summaryTitleFormat.set_font_size(12)
    summaryTitleFormat.set_align('vcenter')
    summaryTitleFormat.set_bottom(1)
    summaryTitleFormat.set_bottom_color(lineColor)

    worksheet.set_row(9, 25)

    worksheet.merge_range('A10:M10', 'Summary', summaryTitleFormat)

    summaryStatTitleFormat = workbook.add_format()

    summaryStatTitleFormat.set_font_name('Helvetica Neue')
    summaryStatTitleFormat.set_bold()
    summaryStatTitleFormat.set_align('top')
    summaryStatTitleFormat.set_font_size(10)
    summaryStatTitleFormat.set_bottom(4)
    summaryStatTitleFormat.set_right(1)
    summaryStatTitleFormat.set_right_color(lineColor)

    worksheet.merge_range('A11:D11', 'Total orders listed below', summaryStatTitleFormat)
    worksheet.merge_range('A12:D12', 'Matched orders listed below', summaryStatTitleFormat)
    worksheet.merge_range('A13:D13', 'Percentage of matched orders', summaryStatTitleFormat)
    worksheet.merge_range('A14:D14', 'Matched at', summaryStatTitleFormat)

    lastRowTitleFormat = workbook.add_format()
    lastRowTitleFormat.set_font_name('Helvetica Neue')
    lastRowTitleFormat.set_bold()
    lastRowTitleFormat.set_align('top')
    lastRowTitleFormat.set_font_size(10)
    lastRowTitleFormat.set_right(1)
    lastRowTitleFormat.set_right_color(lineColor)
    lastRowTitleFormat.set_bottom(1)
    lastRowTitleFormat.set_bottom_color(lineColor)
    worksheet.merge_range('A15:D15', 'Matched lots', lastRowTitleFormat)

    worksheet.set_row(10, 20)
    worksheet.set_row(11, 20)
    worksheet.set_row(12, 20)
    worksheet.set_row(13, 20)
    worksheet.set_row(14, 20)
    worksheet.set_row(15, 20)

    summaryStatValueProps = {
        'font_name': 'Helvetica Neue Light', 
        'align': 'top',
        'font_size': 10,
        'bottom': 4
    }

    summaryStatValueFormat = workbook.add_format(summaryStatValueProps)

    summaryStatValuePropsWithPercentage = dict.copy(summaryStatValueProps)
    summaryStatValuePropsWithPercentage['num_format'] = '0%'
    summaryStatValueFormatWithPercentage = workbook.add_format(summaryStatValuePropsWithPercentage)

    itemCount = len(items)
    lastRow = row + itemCount - 1

    summaryStatValuePropsWithDecimal = dict.copy(summaryStatValueProps)
    summaryStatValuePropsWithDecimal['num_format'] = '#,##0.00'
    summaryStatValueFormatWithDecimal = workbook.add_format(summaryStatValuePropsWithDecimal)

    worksheet.merge_range('E14:F14', qty_dict['match_price'], summaryStatValueFormatWithDecimal)

    lastRowValueFormat = workbook.add_format()
    lastRowValueFormat.set_font_name('Helvetica Neue Light')
    lastRowValueFormat.set_align('top')
    lastRowValueFormat.set_font_size(10)
    lastRowValueFormat.set_num_format('#,###')
    lastRowValueFormat.set_bottom(1)
    lastRowValueFormat.set_bottom_color(lineColor)
    worksheet.merge_range('E15:F15', qty_dict['match_qty'], lastRowValueFormat)

    worksheet.merge_range('G11:M11', 'Number of orders enter auction period and shown below', summaryStatValueFormat)
    worksheet.merge_range('G12:M12', 'Number of orders which are matched and shown below', summaryStatValueFormat)
    worksheet.merge_range('G13:M13', 'Percentage of total orders which are matched', summaryStatValueFormat)
    worksheet.merge_range('G14:M14', 'Price matched at', summaryStatValueFormat)
    worksheet.merge_range('G15:M15', 'Total lots matched', lastRowValueFormat)

    marketStatsTableTitleFormatProps = {
        'font_name': 'Helvetica Neue Medium',
        'font_size': 12,
        'text_wrap': True,
        'valign': 'vcenter'
    }

    marketStatsTableTitleFormat = workbook.add_format(marketStatsTableTitleFormatProps)

    worksheet.merge_range('A18:M18', 'MARKET STATS', marketStatsTableTitleFormat)

    worksheet.set_row(17, 30)

    marketStatsTableHeaderFormatProps = {
        'font_name': 'Helvetica Neue',
        'font_size': 10,
        'text_wrap': True,
        'bold': True,
        'valign': 'vcenter',
        'bg_color': '#F0F5F8',
        'font_color': '#367DA2',
        'align': 'center'
    }

    marketStatsTableHeaderFormat = workbook.add_format(marketStatsTableHeaderFormatProps)

    worksheet.merge_range('A19:C19', '', marketStatsTableHeaderFormat)
    worksheet.merge_range('D19:H19', 'MARKET', marketStatsTableHeaderFormat)
    worksheet.merge_range('I19:M19', 'LIMIT', marketStatsTableHeaderFormat)

    worksheet.set_row(18, 20)

    marketStatsFirstColumnProps = {
        'font_name': 'Helvetica Neue',
        'bold': True,
        'valign': 'top',
        'font_size': 10,
        'top': 4,
        'right': 1,
        'right_color': lineColor
    }

    marketStatsFirstColumnFormat = workbook.add_format(marketStatsFirstColumnProps)

    worksheet.merge_range('A20:C20', 'BUY', marketStatsFirstColumnFormat)
    worksheet.merge_range('A21:C21', 'SELL', marketStatsFirstColumnFormat)
    worksheet.merge_range('A22:C22', 'TOTAL', marketStatsFirstColumnFormat)

    marketStatsValueColumnProps = {
        'font_name': 'Helvetica Neue',
        'valign': 'top',
        'font_size': 10,
        'top': 4,
        'num_format': '#,###'
    }

    marketStatsValueColumnFormat = workbook.add_format(marketStatsValueColumnProps)

    worksheet.merge_range('D20:H20', qty_dict['buy_mkt_qty'], marketStatsValueColumnFormat)
    worksheet.merge_range('D21:H21', qty_dict['sell_mkt_qty'], marketStatsValueColumnFormat)
    worksheet.merge_range('D22:H22', qty_dict['buy_mkt_qty'] + qty_dict['sell_mkt_qty'], marketStatsValueColumnFormat)

    worksheet.merge_range('I20:M20', qty_dict['buy_limit_qty'], marketStatsValueColumnFormat)
    worksheet.merge_range('I21:M21', qty_dict['sell_limit_qty'], marketStatsValueColumnFormat)
    worksheet.merge_range('I22:M22', qty_dict['buy_limit_qty'] + qty_dict['sell_limit_qty'], marketStatsValueColumnFormat)

    worksheet.set_row(19, 23)
    worksheet.set_row(20, 23)
    worksheet.set_row(21, 23)

    ordersTableHeaderProps = {
        'font_name': 'Helvetica Neue',
        'font_size': 10,
        'text_wrap': True,
        'bold': True,
        'bg_color': lineColor,
        'font_color': '#FFFFFF',
        'align': 'center'
    }

    ordersTableHeaderFormat = workbook.add_format(ordersTableHeaderProps)

    worksheet.merge_range('A25:C25', 'ORDERS', ordersTableHeaderFormat)
    worksheet.merge_range('D25:E25', 'SIDE', ordersTableHeaderFormat)
    worksheet.merge_range('F25:G25', 'PRICE', ordersTableHeaderFormat)
    worksheet.merge_range('H25:I25', 'SIZE', ordersTableHeaderFormat)
    worksheet.merge_range('J25:K25', 'TIME IN', ordersTableHeaderFormat)
    worksheet.merge_range('L25:M25', 'TIME OUT', ordersTableHeaderFormat)

    worksheet.set_row(24, 23)

    ordersIDColumnProps = {
        'font_name': 'Helvetica Neue',
        'font_size': 10,
        'text_wrap': True,
        'bold': True,
        'top': 1,
        'bottom': 1,
        'right': 1,
        'left': 1
    }

    ordersIDColumnFormat = workbook.add_format(ordersIDColumnProps)

    ordersValueColumnProps = {
        'font_name': 'Helvetica Neue Light',
        'font_size': 10,
        'text_wrap': True,
        'top': 1,
        'bottom': 1,
        'right': 1,
        'left': 1,
        'align': 'center'
    }

    ordersValueColumnFormat = workbook.add_format(ordersValueColumnProps)

    orderPriceColumnProps = dict.copy(ordersValueColumnProps)
    orderPriceColumnProps['num_format'] = '#,##0.00'

    ordersPriceColumnFormat = workbook.add_format(orderPriceColumnProps)

    orderSizeColumnProps = dict.copy(ordersValueColumnProps)
    orderSizeColumnProps['num_format'] = '#,###'

    ordersSizeColumnFormat = workbook.add_format(orderSizeColumnProps)

    orderTimeColumnProps = dict.copy(ordersValueColumnProps)
    orderTimeColumnProps['right'] = 4
    orderTimeColumnProps['bottom'] = 4
    orderTimeColumnProps['top'] = 0

    ordersTimeColumnFormat = workbook.add_format(orderTimeColumnProps)


    matchedRows = 0
    for lineitem in items :
        cells = "A{!s}:C{!s}".format(row, row)
        worksheet.merge_range(cells, lineitem['id'], ordersIDColumnFormat)
        worksheet.merge_range("D{!s}:E{!s}".format(row, row), lineitem['side'], ordersValueColumnFormat)
        worksheet.merge_range("F{!s}:G{!s}".format(row, row), lineitem['price'], ordersPriceColumnFormat)
        worksheet.merge_range("H{!s}:I{!s}".format(row, row), lineitem['size'], ordersSizeColumnFormat)
        worksheet.merge_range("J{!s}:K{!s}".format(row, row), lineitem['time-in'], ordersTimeColumnFormat)
        worksheet.merge_range("L{!s}:M{!s}".format(row, row), lineitem['time-out'], ordersTimeColumnFormat)
        if (lineitem['time-out'] == 'MATCHED') :
            matchedRows += 1
        row += 1

    worksheet.merge_range('E11:F11', itemCount, summaryStatValueFormat)
    worksheet.merge_range('E12:F12', matchedRows, summaryStatValueFormat)
    average = 0.0
    if matchedRows > 0 :
        average = Decimal(matchedRows) / Decimal(itemCount)
    print average
    worksheet.merge_range('E13:F13', average, summaryStatValueFormatWithPercentage)

    greyColorFill = '#EFEFEF'

    printArea = "A1:M{!s}".format(lastRow)

    worksheet.print_area(printArea) 
    worksheet.set_landscape()
    worksheet.set_paper(12) # A3
    workbook.close()

    return output_dir
