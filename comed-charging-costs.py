#!/usr/local/bin/python3


#
# Import all necessary libraries
#

import argparse
import json
import time
from datetime import datetime
from datetime import date
from datetime import timedelta


#
# Define some global constants
#

VERSION= '0.0.3'

SECONDS_PER_MINUTE= 60
MINUTES_PER_HOUR= 60
HOURS_PER_DAY= 24
MINUTES_PER_DAY= MINUTES_PER_HOUR * HOURS_PER_DAY
SECONDS_PER_DAY= MINUTES_PER_DAY * SECONDS_PER_MINUTE
MINUTE_ALIGNMENT= 5

CENTS_PER_DOLLAR= 100


# 5 hours should charge a car from 30% to 80% (240VAC / 40A)
DEFAULT_DAILY_CHARGE_MINUTES= 300

# 01:00 or 60 minutes after midnight
DEFAULT_DAILY_SCHEDULE_START= 60

# 8 hours daily window of opportunity to charge
DEFAULT_DAILY_CHARGE_WINDOW= 480

# 2 cents per kWh
DEFAULT_MAX_PRICE= 2

# in minutes
DEFAULT_LAG= 10


KEY_TIME= 'millisUTC'
KEY_PRICE= 'price'


#
# Define our functions
#

# Collect all expected and detected arguments from the command line
#
def GetArguments():
  argumentParser= argparse.ArgumentParser()
    
  argumentParser.add_argument('-p', '--pricesFilePath',
    dest='pricesFilePath', required=True, action='store',
    help='Path to a file containing historical ComEd prices')
    
    
  argumentParser.add_argument('-s', '--schedule',
    dest='scheduleStart', required=False, action='store',
    help='Time when daily scheduled session starts')

  argumentParser.add_argument('--span', '--scheduleSpan',
    dest='scheduleSpan', required=False, action='store',
    help='Time span of daily scheduled session')
    
    
  argumentParser.add_argument('-i', '--ideal', dest='ideal', required=False,
    action='store_true', default=False,
    help='Analyze ideal charging scenario')
  
  argumentParser.add_argument('--idealPrice', '--idealMaxPrice',
    dest='idealMaxPrice', type=float, required=False, action='store',
    help='Maximum price (in cents) for ideal charging')

  argumentParser.add_argument('--idealSpan',
    dest='idealSpan', required=False, action='store',
    help='Combined time span of ideal daily sessions')
    
    
  argumentParser.add_argument('-t', '--triggered',
    dest='triggeredStart', required=False, action='store',
    help='Time when daily triggered session starts')
  
  argumentParser.add_argument('--triggeredPrice', '--triggeredMaxPrice',
    dest='triggeredMaxPrice', type=float, required=False, action='store',
    help='Maximum price (in cents) for ideal charging')

  argumentParser.add_argument('--triggeredSpan',
    dest='triggeredSpan', required=False, action='store',
    help='Combined time span of triggered daily sessions')

  argumentParser.add_argument('--triggeredWindow',
    dest='triggeredWindow', required=False, action='store',
    help='Continuos window of time when a car is connected to power')

  argumentParser.add_argument('--triggeredLag',
    dest='triggeredLag', required=False, action='store',
    help='Lag between price time and start or end time of a triggered session')


  diagnostics= argumentParser.add_mutually_exclusive_group()
  diagnostics.add_argument('-d', '--debug', dest='debug', required=False,
    action='store_true', default=False, help='Activate verbose diagnostics')
  diagnostics.add_argument('-q', '--quiet', dest='quiet', required=False,
    action='store_true', default=False, help='Suppress non-critical messages')

  argumentParser.add_argument('--version', action='version',
    version='%(prog)s '+VERSION)


  options= argumentParser.parse_args()
  
  if (options.ideal == False
      and (options.idealMaxPrice != None or options.idealSpan != None)):
        options.ideal= True
  
  if (options.triggeredStart == None
      and (options.triggeredMaxPrice != None or options.triggeredSpan != None
      or options.triggeredLag != None)):
        if options.scheduleStart != None:
          options.triggeredStart= options.scheduleStart
        else:
          options.triggeredStart= 0
    
  
  return options


# Read historical prices from a JSON file
#
def GetPrices(pricesFilePath):
  with open(pricesFilePath, 'r') as pricesFileObject:
    priceTuples= json.load(pricesFileObject)
    
  # convert array of 2-tuples into a dictionary
  prices= {}
  for price in priceTuples:
    prices[int(price[KEY_TIME]) // 1000]= price[KEY_PRICE]

  return prices
  
  
# Compute charging cost based on a static schedule
#
def ComputeCostOfStaticSchedule(prices, options):
  options.scheduleStart= ConvertToMinutes(options.scheduleStart)
  offset= options.scheduleStart % MINUTE_ALIGNMENT
  if offset != 0:
    # align to the next five-minute interval
    options.scheduleStart+= (MINUTE_ALIGNMENT - offset)
  
  if options.scheduleSpan == None:
    options.scheduleSpan= DEFAULT_DAILY_CHARGE_MINUTES
  else:
    options.scheduleSpan= ConvertToMinutes(options.scheduleSpan)

    
  if not options.quiet:
    print()
    if options.scheduleSpan <= MINUTES_PER_DAY:
      print('Static schedule starts at '
        + ConvertToTimeString(options.scheduleStart)
        + ' and runs until '
        + ConvertToTimeString(options.scheduleStart + options.scheduleSpan))
    else:
      print('Static schedule span {:.2f} hours'.format(
        options.scheduleSpan / MINUTES_PER_HOUR)
        + ' exceeds 24 hours -- this schedule cannot realistically work!')
      return DEFAULT_MAX_PRICE
    print()
    
  
  timeStamps= sorted(prices)
  firstDate= date.fromtimestamp(timeStamps[0])
  lastDate= date.fromtimestamp(timeStamps[-1])
  
  if not options.quiet:
    print('Analyzing prices from {} [{:d}] to {} [{:d}]:'.format(
      firstDate, int(timeStamps[0]), lastDate, int(timeStamps[-1])))
    
  chargeDate= firstDate
  oneDay= timedelta(days=1)
  days= 0
  counter= 0
  accumulator= 0
  
  while chargeDate <= lastDate:
    chargeDateAsSeconds= time.mktime(chargeDate.timetuple())
    chargeSecond= chargeDateAsSeconds + options.scheduleStart * SECONDS_PER_MINUTE
    chargeSecondStop= chargeSecond + options.scheduleSpan * SECONDS_PER_MINUTE
    
    if options.debug:
      print()
      print('\t Charging from [{}] to [{}]'.format(
        datetime.fromtimestamp(chargeSecond), datetime.fromtimestamp(chargeSecondStop)))
      
    intervalDaily= 0
    while chargeSecond < chargeSecondStop:
      intervalDaily+= 1
      
      if chargeSecond in prices:
        counter+= 1
        price= float(prices[chargeSecond])
        accumulator+= price
        if options.debug:
          print('\t\t {:>3d}: {:>5.1f} cents at [{}]'.format(
            intervalDaily, price, datetime.fromtimestamp(chargeSecond)))
        
      chargeSecond+= MINUTE_ALIGNMENT * SECONDS_PER_MINUTE
      
    chargeDate+= oneDay
    days+= 1
    
  averagePrice= accumulator / counter
  
  if not options.quiet:
    if options.debug:
      print()
      print()
    print('\t Average price of {:.3f} cents over {:d} day{} [{:d} interval{}]'.format(
      averagePrice, days, PluralS(days), counter, PluralS(counter)))
    
  return averagePrice
  
  
# Compute charging cost based on ideal charging with a price ceiling
#
def ComputeCostOfIdealCharging(prices, options):
  if options.idealMaxPrice == None:
    if options.waterlinePrice != None:
      options.idealMaxPrice= options.waterlinePrice
    else:
      options.idealMaxPrice= DEFAULT_MAX_PRICE
  
  if options.idealSpan != None:
    options.idealSpan= ConvertToMinutes(options.idealSpan)
  elif options.scheduleSpan != None:
    options.idealSpan= ConvertToMinutes(options.scheduleSpan)
  else:
      options.idealSpan= DEFAULT_DAILY_CHARGE_MINUTES
    
  if not options.quiet:
    print()
    print()
    print('Ideal charging target {:d} minute{}'.format(
      options.idealSpan, PluralS(options.idealSpan))
      + ' per day costing {:.3f} cent{} or less'.format(
      options.idealMaxPrice, PluralS(options.idealMaxPrice)))
    print()
    
  
  timeStamps= sorted(prices)
  firstDate= date.fromtimestamp(timeStamps[0])
  lastDate= date.fromtimestamp(timeStamps[-1])

  if not options.quiet:
    print('Analyzing prices from {} [{:d}] to {} [{:d}]:'.format(
      firstDate, int(timeStamps[0]), lastDate, int(timeStamps[-1])))

  chargeDate= firstDate
  oneDay= timedelta(days=1)
  deficitIntervals= 0
  days= deficitDays= 0
  intervalsTotal= 0
  costTotal= 0
  
  while chargeDate <= lastDate:
    chargeDateAsSeconds= time.mktime(chargeDate.timetuple())
    chargeSecond= 0
    intervals= 0
    dailyPrices= {}
    
    if options.debug:
      print()
      print('\t Charge date= {} from [{:d}] to [{:d}]'.format(
        chargeDate, int(chargeDateAsSeconds + chargeSecond),
        int(chargeDateAsSeconds + SECONDS_PER_DAY)))

    # extract daily prices which satisfy our maximum price
    while chargeSecond < SECONDS_PER_DAY:
      intervalKey= int(chargeDateAsSeconds + chargeSecond)
      
      if intervalKey in prices:
        if float(prices[intervalKey]) <= options.idealMaxPrice:
          if prices[intervalKey] in dailyPrices:
            dailyPrices[prices[intervalKey]]+= 1
          else:
            dailyPrices[prices[intervalKey]]= 1
        
      chargeSecond+= MINUTE_ALIGNMENT * SECONDS_PER_MINUTE
      
    # accumulate intervals and prices until "charged"
    intervals= options.idealSpan // MINUTE_ALIGNMENT + deficitIntervals
    sortedPrices= sorted(dailyPrices)
    sortedPricesIndex= 0
    while intervals > 0 and len(sortedPrices) > sortedPricesIndex:
      lowestPrice= sortedPrices[sortedPricesIndex]
      intervalsAtThisPrice= min(dailyPrices[lowestPrice], intervals)
      intervalsTotal+= intervalsAtThisPrice
      costTotal+= intervalsAtThisPrice * float(lowestPrice)
            
      intervals-= intervalsAtThisPrice
      sortedPricesIndex+= 1
      
      if options.debug:
        print(('\t\t Price= {:.3f}; intervals= {:d}; '
          + 'prices remaining= {:d}').format(
          float(lowestPrice), intervalsAtThisPrice,
          int(len(sortedPrices) - sortedPricesIndex)))

    deficitIntervals= max(intervals, 0)
    if deficitIntervals > 0:
      deficitDays+= 1
    
    if options.debug:
      print('\t\t Deficit intervals: {:d}'.format(deficitIntervals))
        
    chargeDate+= oneDay
    days+= 1


  if not options.quiet:
    if options.debug:
      print()
      print()
      
    averagePrice= costTotal / intervalsTotal
    print('\t Average price of {:.3f} cent{} over {:d} day{} [{:d} interval{}]'.format(
      averagePrice, PluralS(averagePrice), days, PluralS(days),
      intervalsTotal, PluralS(intervalsTotal)))
    if deficitDays > 0:
      print('\t Could not fully charge on {:d} day{} during this period'.format(
        deficitDays, PluralS(deficitDays)))
    
  return True
  
  
# Compute charging cost based on triggered charging with a price ceiling and lag
#
def ComputeCostOfTriggeredCharging(prices, options):
  options.triggeredStart= ConvertToMinutes(options.triggeredStart)
  offset= options.triggeredStart % MINUTE_ALIGNMENT
  if offset != 0:
    # align to the next five-minute interval
    options.triggeredStart+= (MINUTE_ALIGNMENT - offset)
    
  if options.triggeredMaxPrice == None:
    if options.waterlinePrice != None:
      options.triggeredMaxPrice= options.waterlinePrice
    else:
      options.triggeredMaxPrice= DEFAULT_MAX_PRICE
  
  if options.triggeredSpan != None:
    options.triggeredSpan= ConvertToMinutes(options.triggeredSpan)
  elif options.scheduleSpan != None:
    options.triggeredSpan= ConvertToMinutes(options.scheduleSpan)
  else:
    options.triggeredSpan= DEFAULT_DAILY_CHARGE_MINUTES
  
  if options.triggeredWindow != None:
    options.triggeredWindow= ConvertToMinutes(options.triggeredWindow)
  elif options.scheduleSpan != None:
    options.triggeredWindow= ConvertToMinutes(options.scheduleSpan * 2)
  else:
    options.triggeredWindow= DEFAULT_DAILY_CHARGE_WINDOW
  
  if options.triggeredLag != None:
    options.triggeredLag= ConvertToMinutes(options.triggeredLag)
    offset= options.triggeredLag % MINUTE_ALIGNMENT
    if offset != 0:
      # align to the next five-minute interval
      options.triggeredLag+= (MINUTE_ALIGNMENT - offset)
  else:
    options.triggeredLag= DEFAULT_LAG
    
  if not options.quiet:
    print()
    print()
    if options.triggeredSpan <= MINUTES_PER_DAY:
      print('Triggered charging starts at '
        + ConvertToTimeString(options.triggeredStart)
        + ' and runs until '
        + ConvertToTimeString(options.triggeredStart + options.triggeredWindow))
      print(' active for {:d} minute{} when price is {:.3f} cent{} or less'.format(
        options.triggeredSpan, PluralS(options.triggeredSpan),
        options.triggeredMaxPrice, PluralS(options.triggeredMaxPrice))
        + ' and lagging {:d} minute{} behind priced time'.format(
        options.triggeredLag, PluralS(options.triggeredLag))
        )
    print()
    
  
  timeStamps= sorted(prices)
  firstDate= date.fromtimestamp(timeStamps[0])
  lastDate= date.fromtimestamp(timeStamps[-1])

  if not options.quiet:
    print('Analyzing prices from {} [{:d}] to {} [{:d}]:'.format(
      firstDate, int(timeStamps[0]), lastDate, int(timeStamps[-1])))

  chargeDate= firstDate
  oneDay= timedelta(days=1)
  deficitIntervals= 0
  days= deficitDays= 0
  intervalsTotal= 0
  costTotal= 0
  
  while chargeDate <= lastDate:
    chargeDateAsSeconds= time.mktime(chargeDate.timetuple())
    chargeSecond= chargeDateAsSeconds + options.triggeredStart * SECONDS_PER_MINUTE
    chargeSecondStop= chargeSecond + options.triggeredWindow * SECONDS_PER_MINUTE
    
    if options.debug:
      print()
      print('\t Charging from [{}] to [{}]'.format(
        datetime.fromtimestamp(chargeSecond), datetime.fromtimestamp(chargeSecondStop)))
      
    intervalDaily= 0
    stopLag= startLag= -MINUTE_ALIGNMENT
    charging= False
    intervals= options.triggeredSpan // MINUTE_ALIGNMENT + deficitIntervals
    while chargeSecond < chargeSecondStop:
      intervalDaily+= 1
      if options.debug:
        diagnosticMessage= '\t\t {:>3d}:'.format(intervalDaily)
      
      if chargeSecond in prices:
        price= float(prices[chargeSecond])
        
        if options.debug:
          diagnosticMessage+= ' [{}]'.format(datetime.fromtimestamp(chargeSecond))
          
          
        # should we stop charging?
        if stopLag == 0:
          charging= False
        elif price > options.triggeredMaxPrice and stopLag < 0:
          stopLag= options.triggeredLag
          
        # should we start charging?
        if startLag == 0:
          charging= True
        elif price <= options.triggeredMaxPrice and startLag < 0:
          startLag= options.triggeredLag
          
        if charging:
          intervals-= 1
          intervalsTotal+= 1
          costTotal+= price
          if options.debug:
            diagnosticMessage+= ' Charging at {:.1f} cents'.format(price)
            if stopLag >= MINUTE_ALIGNMENT:
              diagnosticMessage+= ', but stopping in {:d} minute{}'.format(
                stopLag, PluralS(stopLag))
        else:
          if options.debug:
            diagnosticMessage+= ' Not charging at {:.1f} cents'.format(price)
            if startLag >= MINUTE_ALIGNMENT:
              diagnosticMessage+= ', but starting in {:d} minute{}'.format(
                startLag, PluralS(startLag))
      
      chargeSecond+= MINUTE_ALIGNMENT * SECONDS_PER_MINUTE
      stopLag-= MINUTE_ALIGNMENT
      startLag-= MINUTE_ALIGNMENT
      
      if options.debug:
        print(diagnosticMessage)
        
    # tally for charging session
    deficitIntervals= max(intervals, 0)
    if deficitIntervals > 0:
      deficitDays+= 1

    if options.debug:
      print('\t\t Deficit intervals: {:d}'.format(deficitIntervals))
        
    chargeDate+= oneDay
    days+= 1


  if not options.quiet:
    if options.debug:
      print()
      print()
      
    averagePrice= costTotal / intervalsTotal
    print('\t Average price of {:.3f} cent{} over {:d} day{} [{:d} interval{}]'.format(
      averagePrice, PluralS(averagePrice), days, PluralS(days),
      intervalsTotal, PluralS(intervalsTotal)))
    if deficitDays > 0:
      print('\t Could not fully charge on {:d} day{} during this period'.format(
        deficitDays, PluralS(deficitDays)))
    
  return True
  
  
# Convert a string or a numeric value to an integer representing minutes
#
def ConvertToMinutes(timeValue):
  if not isinstance(timeValue, (float, str, int)):
    raise AssertioError("Time values may only be of type 'float', 'int', or 'str'")
  
  # just treat an integer value as minutes
  if isinstance(timeValue, int):
    return timeValue
    
  # convert a float to an int and return as minutes
  if isinstance(timeValue, float):
    return int(timeValue)
    
  # extract time from a string which may contain ":", "h[ours]", "m[inutes]", "am," or "pm"
  if isinstance(timeValue, str):
    # find and process relative minutes or hours
    position= timeValue.find('h')
    if position > 0:
      return int(timeValue[:position]) * MINUTES_PER_HOUR
      
    position= timeValue.find('m')
    if position > 0:
      # 'm' has no value meaning in further processing -- just drop it
      timeValue= timeValue[:position]
      
     
    # find and process absolute or mixed hours and minutes
    offset= 0
    hoursMultiplier= 1
    
    position= timeValue.find('a')
    if position > 0:
      timeValue= timeValue[:position]
      hoursMultiplier= MINUTES_PER_HOUR
            
    position= timeValue.find('p')
    if position > 0:
      timeValue= timeValue[:position]
      offset= HOURS_PER_DAY * MINUTES_PER_HOUR / 2
      hoursMultiplier= MINUTES_PER_HOUR
      
    position= timeValue.find(':')
    if position > 0:
      values= timeValue.split(':')
      return int(int(values[0]) * MINUTES_PER_HOUR + int(values[1]) + offset)
    else:
      return int(int(timeValue) * hoursMultiplier + offset)
        
             
  return DEFAULT_DAILY_SCHEDULE_START
    

# Convert minutes to time after midnight
#
def ConvertToTimeString(minutes):
  if minutes > MINUTES_PER_DAY:
    days= minutes // MINUTES_PER_DAY
    minutes-= days * MINUTES_PER_DAY
    if days == 1:
      dayString= ' the next day'
    else:
      dayString= ' {:d} day{} later'.format(days, PluralS(days))
  else:
    dayString= ''

  timeString= '{:02d}:{:02d}'.format(
    minutes // MINUTES_PER_HOUR, minutes % MINUTES_PER_HOUR)
    
  return timeString + dayString


    

# Should there be an 's' at the end?
#
def PluralS(number):
  if float(number) == 1:
    return ''
  else:
    return 's'
  

# Main entry point
#
def main():
  try:
    options= GetArguments()
    prices= GetPrices(options.pricesFilePath)
    
    if options.scheduleStart is not None:
      options.waterlinePrice= ComputeCostOfStaticSchedule(prices, options)
    else:
      options.waterlinePrice= DEFAULT_MAX_PRICE
      
    if options.ideal:
      ComputeCostOfIdealCharging(prices, options)
      
    if options.triggeredStart is not None:
      ComputeCostOfTriggeredCharging(prices, options)
      
  
  except Exception as error:
    print(type(error))
    print(error.args[0])
    for counter in range(1, len(error.args)):
      print('\t' + str(error.args[counter]))
      
    if options.debug:
      print('[Debug] Traceback:')
      errorType, errorValue, errorTraceback= sys.exc_info()
      traceback.print_exception(errorType, errorValue, errorTraceback, limit=2, file=sys.stderr)
    

  else:
    if not options.quiet:
      print('')
    
    if options.debug:
      print('[Debug] All done!')
      print()


#
# Execute if we were run as a program
#

if __name__ == '__main__':
  main()
