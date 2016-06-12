# import urllib.request
# from urllib.error import HTTPError
import json
import environs
import os
from utils import *
from pprint import pprint
from decimal import Decimal
from datetime import date, timedelta
from makexlsx import *
from subprocess import call

from check_days import *

#OFFICE_CMD = "libreoffice"
OFFICE_CMD = "soffice"

client = handshake('http://dailyauctions.tilde.sg/login/', None, {'username':os.environ.get('login_username', 'Not Set'), 'password':os.environ.get('login_password', 'Not Set')}, None)

yesterday = getLastValidDay()

print "we are going to do auction reports for the date:" + yesterday.strftime('%Y %b %d, %a ')

yesterday_date = yesterday.strftime('%Y-%m-%d')
yesterday_date_in_iso = yesterday.strftime('%Y%m%d')

fileparams = {
	"company": "Grasshopper",
	"symbol":"NK225M",
	"period": "T_close",
	"filename": "{company}_{symbol_short}_{period}_{date}",
	"qty" : "50",
	"match_price": "",
	"date": yesterday_date
}

metadata = {"auction": "", "date": "", "qty": 0}

# params = {'auction':fileparams['symbol'] + '_' + fileparams['period'], 'qty':fileparams['qty'], 'roll':'n'}
# items = get_normal_items_from_url(client, params, fileparams)
# write_to_file(fileparams['filename']'.txt', items)
# json.dump(d, open("text.txt",'w'))


params = {
	'auction':'NK225_T_close', 
	'qty':'80', 
	'date':yesterday_date_in_iso,
	'symbol': 'Nikkei Big',
	'period': 'Close',
	'filename_prefix': ''
}
qty_dict = {}
items = get_grasshopper_items_from_url(client, params, qty_dict, metadata)

print 'Nikkei 80 Big Close'
out_path = write_to_xlsx('NikkeiBigClose80.xlsx', items, qty_dict, metadata)

## write a function to convert the xlsx into pdf using libreoffice
call([OFFICE_CMD, "--headless", "--invisible", "--convert-to", "pdf", out_path +
      "NK225_T_close_" + yesterday_date_in_iso + ".xlsx", "--outdir", out_path])

params = {
	'auction':'NK225_T_open', 
	'qty':'80', 
	'date':yesterday_date_in_iso,
	'symbol': 'Nikkei Big',
	'period': 'Open',
	'filename_prefix': ''
}
qty_dict = {}
items = get_grasshopper_items_from_url(client, params, qty_dict, metadata)

print 'Nikkei 80 Big Open'
write_to_xlsx('NikkeiBigOpen80.xlsx', items, qty_dict, metadata)

## write a function to convert the xlsx into pdf using libreoffice
call(["libreoffice", "--headless", "--invisible", "--convert-to", "pdf", "/src/Apps/DailyAuctionsReportTool/output/NK225_T_open_" + yesterday_date_in_iso + ".xlsx", "--outdir", "/src/Apps/DailyAuctionsReportTool/output"])


params = {
	'auction':'NK225M_T_close', 
	'qty':'250', 
	'date':yesterday_date_in_iso,
	'symbol': 'Nikkei Mini',
	'period': 'Close',
	'filename_prefix': ''
}
qty_dict = {}
items = get_normal_items_from_url(client, params, qty_dict, metadata)

print 'Nikkei 250 Mini Close'
write_to_xlsx('NikkeiMiniClose250.xlsx', items, qty_dict, metadata)

## write a function to convert the xlsx into pdf using libreoffice
call(["libreoffice", "--headless", "--invisible", "--convert-to", "pdf", "/src/Apps/DailyAuctionsReportTool/output/NK225M_T_close_" + yesterday_date_in_iso + ".xlsx", "--outdir", "/src/Apps/DailyAuctionsReportTool/output"])

params = {
	'auction':'NK225M_T_open', 
	'qty':'250', 
	'date':yesterday_date_in_iso,
	'symbol': 'Nikkei Mini',
	'period': 'Open',
	'filename_prefix': ''
}
qty_dict = {}
items = get_normal_items_from_url(client, params, qty_dict, metadata)

print 'Nikkei 250 Mini Open'
write_to_xlsx('NikkeiMiniOpen250.xlsx', items, qty_dict, metadata)

## write a function to convert the xlsx into pdf using libreoffice
call(["libreoffice", "--headless", "--invisible", "--convert-to", "pdf", "/src/Apps/DailyAuctionsReportTool/output/NK225M_T_open_" + yesterday_date_in_iso + ".xlsx", "--outdir", "/src/Apps/DailyAuctionsReportTool/output"])

params = {
	'auction':'TPX_T_close', 
	'qty':'50', 
	'date':yesterday_date_in_iso,
	'symbol': 'Topix',
	'period': 'Close',
	'filename_prefix': ''
}
qty_dict = {}
items = get_normal_items_from_url(client, params, qty_dict, metadata)

print 'Topx 50 Close'
write_to_xlsx('TopixClose50.xlsx', items, qty_dict, metadata)

## write a function to convert the xlsx into pdf using libreoffice
call(["libreoffice", "--headless", "--invisible", "--convert-to", "pdf", "/src/Apps/DailyAuctionsReportTool/output/TPX_T_close_" + yesterday_date_in_iso + ".xlsx", "--outdir", "/src/Apps/DailyAuctionsReportTool/output"])

params = {
	'auction':'TPX_T_open', 
	'qty':'50', 
	'date':yesterday_date_in_iso,
	'symbol': 'Topix',
	'period': 'Open',
	'filename_prefix': ''
}
qty_dict = {}
items = get_normal_items_from_url(client, params, qty_dict, metadata)

print 'Topx 50 Open'
write_to_xlsx('TopixOpen50.xlsx', items, qty_dict, metadata)

## write a function to convert the xlsx into pdf using libreoffice
call(["libreoffice", "--headless", "--invisible", "--convert-to", "pdf", "/src/Apps/DailyAuctionsReportTool/output/TPX_T_open_" + yesterday_date_in_iso + ".xlsx", "--outdir", "/src/Apps/DailyAuctionsReportTool/output"])

params = {
	'auction':'NK225_T_close', 
	'qty':'40', 
	'date':yesterday_date_in_iso,
	'symbol': 'Nikkei Big',
	'period': 'Close', 
	'filename_prefix': 'Derwin40_'
}
qty_dict = {}
items = get_normal_items_from_url(client, params, qty_dict, metadata)

print 'Derwin Nikkei 40 Close'
write_to_xlsx('DerwinNikkei40.xlsx', items, qty_dict, metadata)

## write a function to convert the xlsx into pdf using libreoffice
call(["libreoffice", "--headless", "--invisible", "--convert-to", "pdf", "/src/Apps/DailyAuctionsReportTool/output/Derwin40_NK225_T_close_" + yesterday_date_in_iso + ".xlsx", "--outdir", "/src/Apps/DailyAuctionsReportTool/output"])

params = {
	'auction':'NK225_T_open', 
	'qty':'40', 
	'date':yesterday_date_in_iso,
	'symbol': 'Nikkei Big',
	'period': 'Open', 
	'filename_prefix': 'Derwin40_'
}
qty_dict = {}
items = get_normal_items_from_url(client, params, qty_dict, metadata)

print 'Derwin Nikkei 40 Open'
write_to_xlsx('DerwinNikkei40.xlsx', items, qty_dict, metadata)

## write a function to convert the xlsx into pdf using libreoffice
call(["libreoffice", "--headless", "--invisible", "--convert-to", "pdf", "/src/Apps/DailyAuctionsReportTool/output/Derwin40_NK225_T_open_" + yesterday_date_in_iso + ".xlsx", "--outdir", "/src/Apps/DailyAuctionsReportTool/output"])
