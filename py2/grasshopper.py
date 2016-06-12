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
from check_days import *

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
# params = {'auction':fileparams['symbol'] + '_' + fileparams['period'], 'qty':fileparams['qty'], 'roll':'n'}
# items = get_normal_items_from_url(client, params, fileparams)
# write_to_file(fileparams['filename']'.txt', items)
# json.dump(d, open("text.txt",'w'))



params = {'auction':'NK225_T_close', 'qty':'40', 'roll':'n'}
qty_dict = {}
items = get_normal_items_from_url(client, params, qty_dict, yesterday_date_in_iso)
print 'Derwin special Nikkei 80 Big Close'
write_to_file('DerwinNikkeiBigClose40.txt', items)



params = {'auction':'NK225_T_open', 'qty':'40', 'roll':'n'}
qty_dict = {}
items = get_normal_items_from_url(client, params, qty_dict, yesterday_date_in_iso)
print 'Derwin special Nikkei 80 Big Open'
write_to_file('DerwinNikkeiBigOpen40.txt', items)

params = {'auction':'NK225_T_close', 'qty':'80', 'roll':'n'}
qty_dict = {}
items = get_grasshopper_items_from_url(client, params, qty_dict, yesterday_date_in_iso)
print 'Nikkei 80 Big Close'
write_to_file('NikkeiBigClose80.txt', items)



params = {'auction':'NK225_T_open', 'qty':'80', 'roll':'n'}
qty_dict = {}
items = get_grasshopper_items_from_url(client, params, qty_dict, yesterday_date_in_iso)
print 'Nikkei 80 Big Open'
write_to_file('NikkeiBigOpen80.txt', items)



params = {'auction':'NK225M_T_close', 'qty':'250', 'roll':'n'}
qty_dict = {}
items = get_normal_items_from_url(client, params,  qty_dict, yesterday_date_in_iso)
print 'Nikkei 250 Mini Close'
write_to_file('NikkeiMiniClose250.txt', items)

params = {'auction':'NK225M_T_open', 'qty':'250', 'roll':'n'}
qty_dict = {}
items = get_normal_items_from_url(client, params,  qty_dict, yesterday_date_in_iso)
print 'Nikkei 250 Mini Open'
write_to_file('NikkeiMiniOpen250.txt', items)


params = {'auction':'TPX_T_close', 'qty':'50', 'roll':'n'}
qty_dict = {}
items = get_normal_items_from_url(client, params, qty_dict, yesterday_date_in_iso)
print 'TPX 50 close'
write_to_file('TopixClose50.txt', items)

params = {'auction':'TPX_T_open', 'qty':'50', 'roll':'n'}
qty_dict = {}
items = get_normal_items_from_url(client, params, qty_dict, yesterday_date_in_iso)
print 'TPX 50 open'
write_to_file('TopixOpen50.txt', items)