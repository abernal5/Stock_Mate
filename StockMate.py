# -*- coding: utf-8 -*-
"""
Created on Sat Nov 14 00:34:43 2020

@author: Alonso
"""

import yfinance as yf
import xlrd
import xlwt
import PySimpleGUI as sg
import sys
import time


class Stock(object):
    '''
    This class saves all the details about a selected stock.

    **Parameters**

        info: *dict*
            The dictionary holding information about the stock ticker.
        analyst: *DataFrame, boolean*
            The dataframe holding information about analyst opinions.
            True if data is being read from existing xls.
        xlflag: *boolean*
            True if reading from an existing Full excel sheet
    '''

    def __init__(self, info, analyst, xlflag):
        '''
        Initializes function.
        '''
        if xlflag is True:
            # The name of the company the stock belongs to.
            self.name = info.get('longName')
            # Price stock (to the day)
            self.price = info.get('bid')
            # Dividend rate per stock per year
            self.dividend = info.get('dividendRate')
            # Avg price over fifty days
            self.avgprice = info.get('fiftyDayAverage')
            # Fifty Two Week Highs and Lows. Good for risk analysis.
            self.fivetwohigh = info.get('fiftyTwoWeekHigh')
            self.fivetwolow = info.get('fiftyTwoWeekLow')
            # Peg ratio > 1 shows overvaluation. < 1 shows undervaluation.
            self.peg = info.get('pegRatio')
            # PB Ratio < 1 is great. PB Ratio < 3 is serviceable. Tangible.
            self.pb = info.get('priceToBook')
            # Profit Margins. Must compare inside industry. Higher = Better
            self.profitmargin = info.get('profitMargins')
            # Industry/Sector
            self.sector = info.get('sector')
            # Stock ticker symbol.
            self.symbol = info.get('symbol')
            # Two hundred day average. A good indicator of risk.
            self.twohun_avg = info.get('twoHundredDayAverage')
            # Company website for user to peruse.
            self.website = info.get('website')
            # Analyst general sentiment (self.sentiment)
            sentiments = analyst['Action'].value_counts().index
            mainsentiment = sentiments[0]
            auxsentiment = sentiments[1]
            if mainsentiment == 'up':
                self.sentiment = 'good'
            elif mainsentiment == 'down':
                self.sentiment = 'bad'
            elif mainsentiment == 'main' or mainsentiment == 'init':
                if auxsentiment == 'main' or auxsentiment == 'init':
                    self.sentiment = 'stable'
                elif auxsentiment == 'up:':
                    self.sentiment = 'warm'
                elif auxsentiment == 'down':
                    self.sentiment = 'cold'
                else:
                    self.sentiment = 'stable'
            else:
                self.sentiment = 'stable'
            # Analyst grades (self.grade)
            grades = analyst['To Grade'].value_counts().index
            valuegrade = analyst['To Grade'].value_counts()
            good = 0
            neutral = 0
            bad = 0
            i = 0
            for x in grades:
                if x in {'Buy', 'Overweight', 'Outperform', 'Strong Buy',
                         'Long-Term Buy'}:
                    good += valuegrade[i]
                if x in {'Neutral', 'Hold', 'Market Perform', 'Equal-Weight',
                         'Sector Perform', 'Perform'}:
                    neutral += valuegrade[i]
                if x in {'Sell', 'Underperform', 'Underweight'}:
                    bad += valuegrade[i]
                i += 1
            if good >= neutral and good >= bad:
                self.grade = 'good'
            elif neutral >= good and neutral >= bad:
                self.grade = 'neutral'
            else:
                self.grade = 'bad'
            # Final value
            self.value = 0
            # Amount purchased
            self.purchase = 0
        else:
            # The same as above but directly from an excel file
            self.symbol = info.get('symbol')
            self.name = info.get('name')
            self.price = info.get('price')
            self.dividend = info.get('dividend')
            self.avgprice = info.get('avgprice')
            self.fivetwohigh = info.get('fivetwohigh')
            self.fivetwolow = info.get('fivetwolow')
            self.peg = info.get('peg')
            self.pb = info.get('pb')
            self.profitmargin = info.get('profitmargin')
            self.sector = info.get('sector')
            self.twohun_avg = info.get('twohun_avg')
            self.website = info.get('website')
            self.sentiment = info.get('sentiment')
            self.grade = info.get('grade')
            self.value = 0
            self.purchase = 0


def stock_list_initializer(excel_name, xlflag):
    '''
    This function initializes the stock objects to be examined.

    **Parameters**

        excel_name: *str*
            The excel file to be read in to create Stock objects.
            Can be either full or not.
        xlflag: *boolean*
            True if reading from an existing Full excel sheet

    ** Returns**

        stock_list: *list of Stock objects*
            The created stock objects to be examined.
    '''
    # Reads in excel file
    wb = xlrd.open_workbook(excel_name)
    sheet = wb.sheet_by_index(0)
    ticker_list = []
    for i in range(sheet.nrows):
        ticker_list.append(sheet.cell_value(i, 0))
    stock_list = []
    # Gets rid of excel header
    ticker_list.pop(0)
    # Triggered if a full excel file must be created.
    if xlflag is True:
        for i in ticker_list:
            # Due to constant API calls, I need to further slow this section
            # down. If not, the serves mistake me for a DDOS attack.
            time.sleep(1)
            try:
                data = yf.Ticker(i)
                info = data.info
            except yf.HTTPError:
                break
            analyst = data.recommendations
            stock_list.append(Stock(info, analyst, xlflag))
        # Create full excel sheet for future use
        workbook = xlwt.Workbook()
        worksheet = workbook.add_sheet('Full Worksheet')
        worksheet.write(0, 0, "Symbol")
        worksheet.write(0, 1, "Name")
        worksheet.write(0, 2, "Price")
        worksheet.write(0, 3, "Dividend")
        worksheet.write(0, 4, "Avg. Price")
        worksheet.write(0, 5, "Fifty-Two Week High")
        worksheet.write(0, 6, "Fifty-Two Week Low")
        worksheet.write(0, 7, "PEG Ratio")
        worksheet.write(0, 8, "PB Ratio")
        worksheet.write(0, 9, "Profit Margin")
        worksheet.write(0, 10, "Sector")
        worksheet.write(0, 11, "Two-Hundred Day Average")
        worksheet.write(0, 12, "Website")
        worksheet.write(0, 13, "Sentiment")
        worksheet.write(0, 14, "Grade")
        row = 1
        for j in stock_list:
            worksheet.write(row, 0, j.symbol)
            worksheet.write(row, 1, j.name)
            worksheet.write(row, 2, j.price)
            worksheet.write(row, 3, j.dividend)
            worksheet.write(row, 4, j.avgprice)
            worksheet.write(row, 5, j.fivetwohigh)
            worksheet.write(row, 6, j.fivetwolow)
            worksheet.write(row, 7, j.peg)
            worksheet.write(row, 8, j.pb)
            worksheet.write(row, 9, j.profitmargin)
            worksheet.write(row, 10, j.sector)
            worksheet.write(row, 11, j.twohun_avg)
            worksheet.write(row, 12, j.website)
            worksheet.write(row, 13, j.sentiment)
            worksheet.write(row, 14, j.grade)
            row += 1
        if excel_name[-4:] == ".xls":
            workbook.save(excel_name[:-4] + "_full.xls")
        else:
            workbook.save(excel_name[:-5] + "_full.xls")
    else:
        # If an existing xls file exists, I can skip over all the API calls.
        # This method is much faster, although won't have up-to-date data.
        analyst = True
        for i in range(1, sheet.nrows):
            info = {
                "symbol": sheet.cell_value(i, 0),
                "name": sheet.cell_value(i, 1),
                "price": sheet.cell_value(i, 2),
                "dividend": sheet.cell_value(i, 3),
                "avgprice": sheet.cell_value(i, 4),
                "fivetwohigh": sheet.cell_value(i, 5),
                "fivetwolow": sheet.cell_value(i, 6),
                "peg": sheet.cell_value(i, 7),
                "pb": sheet.cell_value(i, 8),
                "profitmargin": sheet.cell_value(i, 9),
                "sector": sheet.cell_value(i, 10),
                "twohun_avg": sheet.cell_value(i, 11),
                "website": sheet.cell_value(i, 12),
                "sentiment": sheet.cell_value(i, 13),
                "grade": sheet.cell_value(i, 14)
                }
            stock_list.append(Stock(info, analyst, xlflag))
    return stock_list


def mate_calculator(values, stock_list):
    '''
    This function calculates the value of each stock provided.
    It does so by sifting through user preferences and eliminating stocks.
    The remaining stocks are given values based on mathematical models
    and user input responses. Stocks are then purchased until the available
    budget is expended.

    **Parameters**

        values: *dict*
            These are the user-submitted responses.
        stock_list: *list of Stock objects*
            This is the list of stock objects to be examined.

    ** Returns**

        stock_mates: *list of tuples*
            This list of tuples states: which stock ticker to buy,
            how many of that stock to buy, and the company website in case
            the user wants to peruse it. This information is then added to
            an xls (excel) file.
    '''
    # Narrowing based on stability
    result = []
    for k in stock_list:
        if values.get(0) is True and k.sentiment != 'stable':
            result.append(k)
        if values.get(1) is True and k.sentiment in {'warm',
                                                     'stable', 'cold'}:
            result.append(k)
        if values.get(2) is True and k.sentiment == 'stable':
            result.append(k)
        else:
            result.append(k)
    stock_list.clear()
    # Narrowing based on timeframe. Looking for stocks that change drastically
    # to accomodate those with smaller timeframes. Should still be stabilized
    # from above.
    for n in result:
        change = abs(n.fivetwohigh - n.fivetwolow) / n.fivetwohigh
        if values.get(4) is True and change > 0.3:
            stock_list.append(n)
        if values.get(5) is True and change > 0.15:
            stock_list.append(n)
        else:
            stock_list.append(n)
    result.clear()
    # Giving weights to different criteria
    for o in stock_list:
        if values.get(9) is True:
            o.grade = 'neutral'
    # Setting budget and liquidity
    budget = int(values.get(28))
    if values.get(11) is True:
        budget = budget * 0.9
    if values.get(12) is True:
        budget = budget * 0.8
    for p in stock_list:
        if values.get(14) is True and p.dividend > 4:
            p.value += 1
        if values.get(15) is True and p.dividend > 4:
            p.value += 2
    # Giving weight to sector preferences
    for q in stock_list:
        if values.get(17) is True and p.sector == 'Basic Materials':
            q.value += 2
        if values.get(18) is True and p.sector == 'Communication Services':
            q.value += 2
        if values.get(19) is True and p.sector == 'Consumer Cyclical':
            q.value += 2
        if values.get(20) is True and p.sector == 'Consumer Defensive':
            q.value += 2
        if values.get(21) is True and p.sector == 'Energy':
            q.value += 2
        if values.get(22) is True and p.sector == 'Financial Services':
            q.value += 2
        if values.get(23) is True and p.sector == 'Healthcare':
            q.value += 2
        if values.get(24) is True and p.sector == 'Industrials':
            q.value += 2
        if values.get(25) is True and p.sector == 'Real Estate':
            q.value += 2
        if values.get(26) is True and p.sector == 'Technology':
            q.value += 2
        if values.get(27) is True and p.sector == 'Utilities':
            q.value += 2
    # Giving weight from analysts and from financial data
    for r in stock_list:
        if r.grade == 'good':
            r.value += 2
        if r.grade == 'bad':
            r.value -= 2
        if r.pb == '':
            r.value = r.value
        else:
            if float(r.pb) < 3:
                r.value += 1
        if r.peg == '':
            r.value = r.value
        else:
            if float(r.peg) < 4 and float(r.peg) > 0.5:
                r.value += 1
    # Spending my way through the budget
    mark = 1
    while budget >= 1 and mark != 0:
        mark = 0
        for s in stock_list:
            if s.value > 3 and budget >= float(s.price):
                s.purchase += 2
                budget = budget - float(s.price) * 2
                mark += 1
            if s.value > 1 and budget >= float(s.price):
                s.purchase += 1
                budget = budget - float(s.price)
                mark += 1
    stock_mates = []
    # Creating my purchasing information and appending it to resulting list.
    for t in stock_list:
        if t.purchase > 0:
            stock_pair = [t.name, t.purchase, t.website]
            stock_mates.append(stock_pair)
    return stock_mates


if __name__ == "__main__":
    # First pop-up just checks to see if full excel is available.
    layout = [[sg.Text('Welcome!\nDo you need to create a new full '
                       'excel sheet?\n'
                       '(Please be aware, this takes a very long time.)')],
              [sg.Button('Yes'), sg.Button('No')]]

    window = sg.Window('StockMate', layout)
    while True:
        event, values = window.Read()
        if event is None or event == sg.WIN_CLOSED:
            sys.exit()
            break
        if event == 'Yes':
            xlflag = True
            break
        if event == 'No':
            xlflag = False
            break
    window.Close()

    # Second pup-up asks for excel file name, full or otherwise.
    layout = [[sg.Text('Please specify the excel sheet name to use (limited or'
                       ' full).\nInclude the extension.')],
              [sg.Input(), sg.Button('Ok')]]

    window = sg.Window('StockMate', layout)
    while True:
        event, values = window.Read()
        if event is None or event == sg.WIN_CLOSED:
            sys.exit()
            break
        if event == 'Ok':
            break
    window.Close()

    # This is all the information needed to initialize the Stock object list.
    excel_name = values[0]
    stock_list = stock_list_initializer(excel_name, xlflag)

    # Collect user preference data.
    layout = [[sg.Text('All Ready to Start!\nPlease answer the following'
                       ' questions to get a sense of your preferences')],
              [sg.Frame(layout=[
                  [sg.Radio('All on Black (High Risk)', "Risk",
                            default=True),
                   sg.Radio('King of Jacks (Medium Risk)', "Risk"),
                   sg.Radio('Planning for Retirement (Low Risk)', "Risk"),
                   sg.Radio('My What? (N/A)', "Risk")]],
                  title='What is your risk tolerance?')],
              [sg.Frame(layout=[
                  [sg.Radio('I am a Day Trader (< 3 months)', "Lifespan",
                            default=True),
                   sg.Radio("Let's See How It Goes (3-6 months)", "Lifespan"),
                   sg.Radio('My Savings Account (> 6 months)', "Lifespan"),
                   sg.Radio('My What? (N/A)', "Lifespan")]],
                  title='What is your expected investment timeframe?')],
              [sg.Frame(layout=[
                  [sg.Radio("I Trust Them (Yes)", "Analyst",
                            default=True),
                   sg.Radio("I Don't Trust Them (No)", "Analyst")]],
                  title='Should I consider analyst opinion?')],
              [sg.Frame(layout=[
                  [sg.Radio('Liquidity Is My Bank Account (None)', "Liq",
                            default=True),
                   sg.Radio('Keep Some For Good Opportunities (10%)', "Liq"),
                   sg.Radio('I Need To Be Ready To Steer (20%)', "Risk")]],
                  title='How much liquidity should I allow for?')],
              [sg.Frame(layout=[
                  [sg.Radio('I Can Keep Them or Leave Them (Not Much)', "Div",
                            default=True),
                   sg.Radio('Alway Good To Have (Neutral)', "Div"),
                   sg.Radio('Dividends Are My Salary (Very Much)', "Div"),
                   sg.Radio('I Was Never Good at Geometry (N/A)', "Div")]],
                  title='How important are dividends for you?')],
              [sg.Frame(layout=[
                  [sg.Checkbox('Basic Materials'),
                   sg.Checkbox('Communication Services'),
                   sg.Checkbox('Consumer Cyclical'),
                   sg.Checkbox('Consumer Defensive'),
                   sg.Checkbox('Energy'),
                   sg.Checkbox('Financial Services'),
                   sg.Checkbox('Healthcare'),
                   sg.Checkbox('Industrials'),
                   sg.Checkbox('Real Estate'),
                   sg.Checkbox('Technology'),
                   sg.Checkbox('Utilities')]],
                  title='Please select any sectors of interest')],
              [sg.Text("What is your budget? (USD, whole numbers, no commas)"),
               sg.Input(), sg.Button('Next')]]

    window = sg.Window('StockMate', layout)
    while True:
        event, values = window.Read()
        if event is None or event == sg.WIN_CLOSED:
            sys.exit()
            break
        if event == 'Next':
            break
    window.Close()

    # Pass these preferences through the calculator
    stock_choices = mate_calculator(values, stock_list)

    # Create excel with the calculated match information.
    workbook = xlwt.Workbook()
    worksheet = workbook.add_sheet('Your_Mates')
    worksheet.write(0, 0, "Stock Ticker")
    worksheet.write(0, 1, "Amount Suggested")
    worksheet.write(0, 2, "Company Website")
    row = 1
    for m in stock_choices:
        worksheet.write(row, 0, m[0])
        worksheet.write(row, 1, m[1])
        worksheet.write(row, 2, m[2])
        row += 1
    workbook.save("Stock_Mate_Choices.xls")

    # A good-bye message.
    layout = [[sg.Text('Thank you for using Stock Mate!\n'
                       'You will find your results in Stock_Mate_Choices.xls\n'
                       'You will also find your full xls sheet, for faster'
                       ' run time next time you use our service, as:\n'
                       '[Initial Excel Name]_full.xls')],
              [sg.Button('Done')]]

    window = sg.Window('StockMate', layout)
    while True:
        event, values = window.Read()
        if event is None or event == sg.WIN_CLOSED:
            sys.exit()
            break
        if event == 'Done':
            break
    window.Close()
