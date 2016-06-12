from datetime import date, timedelta
import pandas as pd
# BDay is business day, not birthday...
from pandas.tseries.offsets import BDay

twenty16ExchangeHolidays = [];

## reference from http://www.jpx.co.jp/english/corporate/calendar/
# March 20 Vernal Equinox
twenty16ExchangeHolidays.append(pd.datetime(2016, 3, 20))
# March 21 Holiday
twenty16ExchangeHolidays.append(pd.datetime(2016, 3, 21))
# April 29 Showa Day
twenty16ExchangeHolidays.append(pd.datetime(2016, 4, 29))
# May 3 
twenty16ExchangeHolidays.append(pd.datetime(2016, 5, 3))
# May 4 
twenty16ExchangeHolidays.append(pd.datetime(2016, 5, 4))
# May 5 
twenty16ExchangeHolidays.append(pd.datetime(2016, 5, 5))
# July 18
twenty16ExchangeHolidays.append(pd.datetime(2016, 7, 18))
# August 11
twenty16ExchangeHolidays.append(pd.datetime(2016, 8, 11))
# Septe 19
twenty16ExchangeHolidays.append(pd.datetime(2016, 9, 19))
# Sept 22
twenty16ExchangeHolidays.append(pd.datetime(2016, 9, 22))
# Oct 10
twenty16ExchangeHolidays.append(pd.datetime(2016, 10, 10))
# Nov 3
twenty16ExchangeHolidays.append(pd.datetime(2016, 11, 3))
# Nov 23
twenty16ExchangeHolidays.append(pd.datetime(2016, 11, 23))
# Dec 23
twenty16ExchangeHolidays.append(pd.datetime(2016, 12, 23))
# Dec 31
twenty16ExchangeHolidays.append(pd.datetime(2016, 12, 31))

def checkExchangeHoliday(holidays, date):
	for day in holidays :
		if date == day :
			return True
	return False

def getLastValidDay() :
	# pd.datetime is an alias for datetime.datetime
	today = pd.to_datetime('today')
	lastBusinessDay = today - BDay(1)

	validDay = False

	while (validDay == False):
	
		if (checkExchangeHoliday(twenty16ExchangeHolidays, lastBusinessDay)) :
			lastBusinessDay = lastBusinessDay - BDay(1)
		else :
			validDay = True


	return lastBusinessDay


# lastValidDayInString = lastValidDay.strftime('%Y-%m-%d')
# print lastValidDayInString