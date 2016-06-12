import json
import logging
import requests
import urlparse
import os
import ast
from decimal import Decimal

def handshake(host, client, params, proxies, verify=False):
    log = logging.getLogger(__name__)
    if not client:
        client = requests.session()
    url = urlparse.urlsplit(host)
    host = '%s://%s' % (url.scheme, url.netloc)
    try:
        log.info('handshake params=%r proxies=%r', params, proxies)
        r = client.post('%s/handshake/' % host, data=params, proxies=proxies, verify=verify)
        r.raise_for_status()
    except requests.exceptions.HTTPError, e:
        log.info('handshake post failed error=%s reason=%s', r.reason, str(e))
        r = client.get('%s/handshake/' % host, params=params, proxies=proxies, verify=verify)
        if r.ok:
            r = client.get('%s/login' % host, proxies=proxies, verify=verify)
            csrftoken = client.cookies['csrftoken']
            params['csrfmiddlewaretoken'] = csrftoken
            log.info('got token :) %s', csrftoken)
    return client


def legit_for_grasshopper(line_item, match_price) :
    upper_limit = match_price + 500
    lower_limit = match_price - 500
    if line_item['price'] == 'M':
        return True
    try:
        price = Decimal(line_item['price'])
    except:
        return True
    if (price > upper_limit 
    and line_item['side'] == 'S' 
    and line_item['size'] == 100):
        return False
    if (price < lower_limit 
    and line_item['side'] == 'B'
    and line_item['size'] == 100):
        return False
    return True

def get_grasshopper_items(all_items) :
    selllimitqty = all_items.pop(-1)
    sell_limit_qty = int(selllimitqty['sell-limit-qty'])

    sellmktqty = all_items.pop(-1)
    sell_mkt_qty = int(sellmktqty['sell-mkt-qty'])

    buylimitqty = all_items.pop(-1)
    buy_limit_qty = int(buylimitqty['buy-limit-qty'])

    buymktqty = all_items.pop(-1)
    buy_mkt_qty = int(buymktqty['buy-mkt-qty'])

    matchlimitqty = all_items.pop(-1)
    matchmktqty = all_items.pop(-1)

    matchqty = all_items.pop(-1)
    match_qty = int(matchqty['match-qty'])

    match_price_line_item = all_items.pop(-1)
    match_price = Decimal(match_price_line_item['match-price'])
    
    grasshopper_items = []
    for lineitem in all_items :
        upper_limit = match_price + 500
        lower_limit = match_price - 500

        if legit_for_grasshopper(lineitem, match_price) :
            grasshopper_items.append(lineitem)
    return {'grasshopper_items':grasshopper_items, 'match_price':match_price, 'match_qty': match_qty, 'sell_limit_qty':sell_limit_qty, 'buy_limit_qty':buy_limit_qty, 'sell_mkt_qty':sell_mkt_qty, 'buy_mkt_qty':buy_mkt_qty}

def get_grasshopper_items_from_url(client, params, qty_dict, metadata) :    
    response = client.get('http://dailyauctions.tilde.sg/auction/json', params=params)

    if response.ok:
        data = json.loads( response.text )
        lineitems = data['data']
        result = get_grasshopper_items(lineitems)
        grasshopper_items = result['grasshopper_items']

        qty_dict['match_price'] = result['match_price']
        qty_dict['match_qty'] = result['match_qty']
        
        qty_dict['sell_limit_qty'] = result['sell_limit_qty']
        qty_dict['buy_limit_qty'] = result['buy_limit_qty']
        
        qty_dict['sell_mkt_qty'] = result['sell_mkt_qty']
        qty_dict['buy_mkt_qty'] = result['buy_mkt_qty']        
        
        # normal grasshopper.py needs this to be commented out
        for key, value in params.iteritems():
            metadata[key] = value

        metadata['date'] = data['metadata']['date']
        metadata['qty'] = data['metadata']['qty']
        metadata['auction'] = data['metadata']['auction']
        # end of the special metadata we need

        return grasshopper_items

    return None

def get_normal_items_from_url(client, params, qty_dict, metadata) :    
    response = client.get('http://dailyauctions.tilde.sg/auction/json', params=params)

    if response.ok:
        data = json.loads( response.text )
        lineitems = data['data']
        result = get_normal_items(lineitems)

        grasshopper_items = result['grasshopper_items']

        qty_dict['match_price'] = result['match_price']
        qty_dict['match_qty'] = result['match_qty']
        
        qty_dict['sell_limit_qty'] = result['sell_limit_qty']
        qty_dict['buy_limit_qty'] = result['buy_limit_qty']
        
        qty_dict['sell_mkt_qty'] = result['sell_mkt_qty']
        qty_dict['buy_mkt_qty'] = result['buy_mkt_qty']        
        

        # normal grasshopper.py needs this to be commented out
        for key, value in params.iteritems():
            metadata[key] = value

        metadata['date'] = data['metadata']['date']
        metadata['qty'] = data['metadata']['qty']
        metadata['auction'] = data['metadata']['auction']
        # end of the special metadata we need

        return grasshopper_items

    return None

def get_normal_items(all_items) :
    selllimitqty = all_items.pop(-1)
    sell_limit_qty = int(selllimitqty['sell-limit-qty'])

    sellmktqty = all_items.pop(-1)
    sell_mkt_qty = int(sellmktqty['sell-mkt-qty'])

    buylimitqty = all_items.pop(-1)
    buy_limit_qty = int(buylimitqty['buy-limit-qty'])

    buymktqty = all_items.pop(-1)
    buy_mkt_qty = int(buymktqty['buy-mkt-qty'])

    matchlimitqty = all_items.pop(-1)
    matchmktqty = all_items.pop(-1)

    matchqty = all_items.pop(-1)
    match_qty = int(matchqty['match-qty'])

    match_price_line_item = all_items.pop(-1)
    match_price = Decimal(match_price_line_item['match-price'])
    
    grasshopper_items = []
    for lineitem in all_items :
        grasshopper_items.append(lineitem)
        

    return {'grasshopper_items':grasshopper_items, 'match_price':match_price, 'match_qty': match_qty, 'sell_limit_qty':sell_limit_qty, 'buy_limit_qty':buy_limit_qty, 'sell_mkt_qty':sell_mkt_qty, 'buy_mkt_qty':buy_mkt_qty}

def write_to_file(filename, items) :
    path = os.path.dirname(os.path.dirname(os.path.abspath(__file__))) + "/output/"
    # print path
    f = open(path + filename, 'w')
    for lineitem in items :
        line = "{!s}\t{!s}\t{!s}\t{!s}\t{!s}\t{!s}\n".format(lineitem['id'], lineitem['side'], lineitem['price'], lineitem['size'], lineitem['time-in'], lineitem['time-out'])

        f.write(line)