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

# OFFICE_CMD = "libreoffice"
OFFICE_CMD = "soffice"

def do_auction_reports():
    yesterday = getLastValidDay()
    print "we are going to do auction reports for the date:" + yesterday.strftime('%Y %b %d, %a ')
    yesterday_date = yesterday.strftime('%Y-%m-%d')
    yesterday_date_in_iso = yesterday.strftime('%Y%m%d')

    # List of Dictionary to hold info about pdfs to create
    # xls_name, pdf_name, params, items_getter
    file_tasks = [
        {'xls_name'    : 'NikkeiBigClose80.xlsx',
         'pdf_name'    : 'NK225_T_close_' + yesterday_date_in_iso + ".xlsx",
         'params'      : {
                            'auction':'NK225_T_close',
                            'qty':'80',
                            'date':yesterday_date_in_iso,
                            'symbol': 'Nikkei Big',
                            'period': 'Close',
                            'filename_prefix': ''
                         },
         'item_getter' : get_grasshopper_items_from_url,
        },
        {'xls_name'    : 'NikkeiBigOpen80.xlsx',
         'pdf_name'    : 'NK225_T_open_' + yesterday_date_in_iso + ".xlsx",
         'params'      : {
                            'auction': 'NK225_T_open',
                            'qty': '80',
                            'date':yesterday_date_in_iso,
                            'symbol': 'Nikkei Big',
                            'period': 'Open',
                            'filename_prefix': ''
                         },
         'item_getter' : get_grasshopper_items_from_url,
        },
        {'xls_name'    : 'NikkeiMiniClose250.xlsx',
         'pdf_name'    : 'NK225M_T_close_' + yesterday_date_in_iso + ".xlsx",
         'params'      : {
                            'auction':'NK225M_T_close',
                            'qty':'250',
                            'date':yesterday_date_in_iso,
                            'symbol': 'Nikkei Mini',
                            'period': 'Close',
                            'filename_prefix': ''
                         },
         'item_getter' : get_normal_items_from_url,
        },
        {'xls_name'    : 'NikkeiMiniOpen250.xlsx',
         'pdf_name'    : 'NK225M_T_open_' + yesterday_date_in_iso + ".xlsx",
         'params'      : {
                            'auction':'NK225M_T_open',
                            'qty':'250',
                            'date':yesterday_date_in_iso,
                            'symbol': 'Nikkei Mini',
                            'period': 'Open',
                            'filename_prefix': ''
                         },
         'item_getter' : get_normal_items_from_url,
        },
        {'xls_name'    : 'TopixClose50.xlsx',
         'pdf_name'    : 'TPX_T_close_' + yesterday_date_in_iso + ".xlsx",
         'params'      : {
                            'auction':'TPX_T_close',
                            'qty':'50',
                            'date':yesterday_date_in_iso,
                            'symbol': 'Topix',
                            'period': 'Close',
                            'filename_prefix': ''
                         },
         'item_getter' : get_normal_items_from_url,
        },
        {'xls_name'    : 'TopixOpen50.xlsx',
         'pdf_name'    : 'TPX_T_open_' + yesterday_date_in_iso + ".xlsx",
         'params'      : {
                            'auction':'TPX_T_open',
                            'qty':'50',
                            'date':yesterday_date_in_iso,
                            'symbol': 'Topix',
                            'period': 'Open',
                            'filename_prefix': ''
                         },
         'item_getter' : get_normal_items_from_url,
        },
        {'xls_name'    : 'DerwinNikkei40.xlsx',
         'pdf_name'    : 'Derwin40_NK225_T_close_' + yesterday_date_in_iso + ".xlsx",
         'params'      : {
                            'auction': 'NK225_T_close',
                            'qty':'40',
                            'date':yesterday_date_in_iso,
                            'symbol': 'Nikkei Big',
                            'period': 'Close',
                            'filename_prefix': 'Derwin40_'
                         },
         'item_getter' : get_normal_items_from_url,
        },
        {'xls_name'    : 'DerwinNikkei40.xlsx',
         'pdf_name'    : 'Derwin40_NK225_T_open_' + yesterday_date_in_iso + ".xlsx",
         'params'      : {
                            'auction': 'NK225_T_open',
                            'qty':'40',
                            'date':yesterday_date_in_iso,
                            'symbol': 'Nikkei Big',
                            'period': 'Open',
                            'filename_prefix': 'Derwin40_'
                         },
         'item_getter' : get_normal_items_from_url,
        },
    ]

    client = handshake('http://dailyauctions.tilde.sg/login/', None, {'username':os.environ.get('login_username', 'Not Set'), 'password':os.environ.get('login_password', 'Not Set')}, None)

    metadata = {"auction": "", "date": "", "qty": 0}

    for task in file_tasks:
        qty_dict = {}
        items = task['item_getter'](client, task['params'], qty_dict, metadata)

        print(task['xls_name'])
        out_path = write_to_xlsx(task['xls_name'], items, qty_dict, metadata)

        ## write a function to convert the xlsx into pdf using libreoffice
        call([OFFICE_CMD, "--headless", "--invisible", "--convert-to", "pdf", out_path +
              task['pdf_name'], "--outdir", out_path])

    '''
    Looks like below is not used anywhere

    fileparams = {
        "company": "Grasshopper",
        "symbol":"NK225M",
        "period": "T_close",
        "filename": "{company}_{symbol_short}_{period}_{date}",
        "qty" : "50",
        "match_price": "",
        "date": yesterday_date
    }


    # params = {'auction':fileparams['symbol'] + '_' + fileparams['period'], 'qty':fileparams['qty'], 'roll':'n'}
    # items = get_normal_items_from_url(client, params, fileparams)
    # write_to_file(fileparams['filename']'.txt', items)
    # json.dump(d, open("text.txt",'w'))
    '''

if __name__ == "__main__":
    do_auction_reports()
