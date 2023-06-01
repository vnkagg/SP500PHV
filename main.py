# Dependencies
# SSL certificate verification
# pip libraries : pandas, requests, bs4, yfinance, openpyxl

import yfinance as yf
import pandas as pd
import requests
from bs4 import BeautifulSoup

# List of S&P500 index's constituents from Wikipedia
# ---------------------------------------------------------------------------------------------------#
# data_frame = pd.read_html("https://en.wikipedia.org/wiki/List_of_S%26P_500_companies")[0]
# Get the ticker symbols from the DataFrame
# tickers = data_frame["Symbol"].tolist()

#                                            Note
# I HAVE COMMENTED OUT THE ABOVE CODE TO FETCH THE SYMBOLS OF THE ASSETS OF THE S&P500 INDEX AUTOMATICALLY
# BECAUSE THERE CAN BE POTENTIAL ISSUES WITH SSL CERTIFICATION ON DIFFERENT MACHINES AND SERVERS
# THESE ISSUES CAN BE A BIT LENGTHY TO DEBUG SO FOR NOW THE TICKERS ARE COPIED MANUALLY
# THE AUTOMATION CAN BE DONE BY GIVING IT SOME TIME
# ---------------------------------------------------------------------------------------------------#

tickers = ['MMM','AOS','ABT','ABBV','ACN','ATVI','ADM','ADBE','ADP','AAP','AES',
'AFL','A','APD','AKAM','ALK','ALB','ARE','ALGN','ALLE','LNT','ALL','GOOGL','GOOG',
'MO','AMZN','AMCR','AMD','AEE','AAL','AEP','AXP','AIG','AMT','AWK','AMP','ABC','AME',
'AMGN','APH','ADI','ANSS','AON','APA','AAPL','AMAT','APTV','ACGL','ANET','AJG','AIZ','T',
'ATO','ADSK','AZO','AVB','AVY','AXON','BKR','BALL','BAC','BBWI','BAX','BDX','WRB','BRK.B',
'BBY','BIO','TECH','BIIB','BLK','BK','BA','BKNG','BWA','BXP','BSX','BMY','AVGO','BR','BRO',
'BF.B','BG','CHRW','CDNS','CZR','CPT','CPB','COF','CAH','KMX','CCL','CARR','CTLT','CAT','CBOE',
'CBRE','CDW','CE','CNC','CNP','CDAY','CF','CRL','SCHW','CHTR','CVX','CMG','CB','CHD','CI','CINF',
'CTAS','CSCO','C','CFG','CLX','CME','CMS','KO','CTSH','CL','CMCSA','CMA','CAG','COP','ED','STZ',
'CEG','COO','CPRT','GLW','CTVA','CSGP','COST','CTRA','CCI','CSX','CMI','CVS','DHI','DHR','DRI',
'DVA','DE','DAL','XRAY','DVN','DXCM','FANG','DLR','DFS','DISH','DIS','DG','DLTR','D','DPZ','DOV',
'DOW','DTE','DUK','DD','DXC','EMN','ETN','EBAY','ECL','EIX','EW','EA','ELV','LLY','EMR','ENPH','ETR',
'EOG','EPAM','EQT','EFX','EQIX','EQR','ESS','EL','ETSY','RE','EVRG','ES','EXC','EXPE','EXPD','EXR',
'XOM','FFIV','FDS','FICO','FAST','FRT','FDX','FITB','FSLR','FE','FIS','FISV','FLT','FMC','F','FTNT',
'FTV','FOXA','FOX','BEN','FCX','GRMN','IT','GEHC','GEN','GNRC','GD','GE','GIS','GM','GPC','GILD','GL',
'GPN','GS','HAL','HIG','HAS','HCA','PEAK','HSIC','HSY','HES','HPE','HLT','HOLX','HD','HON','HRL','HST',
'HWM','HPQ','HUM','HBAN','HII','IBM','IEX','IDXX','ITW','ILMN','INCY','IR','PODD','INTC','ICE','IFF',
'IP','IPG','INTU','ISRG','IVZ','INVH','IQV','IRM','JBHT','JKHY','J','JNJ','JCI','JPM','JNPR','K','KDP',
'KEY','KEYS','KMB','KIM','KMI','KLAC','KHC','KR','LHX','LH','LRCX','LW','LVS','LDOS','LEN','LNC','LIN',
'LYV','LKQ','LMT','L','LOW','LYB','MTB','MRO','MPC','MKTX','MAR','MMC','MLM','MAS','MA','MTCH','MKC',
'MCD','MCK','MDT','MRK','META','MET','MTD','MGM','MCHP','MU','MSFT','MAA','MRNA','MHK','MOH','TAP',
'MDLZ','MPWR','MNST','MCO','MS','MOS','MSI','MSCI','NDAQ','NTAP','NFLX','NWL','NEM','NWSA','NWS','NEE',
'NKE','NI','NDSN','NSC','NTRS','NOC','NCLH','NRG','NUE','NVDA','NVR','NXPI','ORLY','OXY','ODFL','OMC',
'ON','OKE','ORCL','OGN','OTIS','PCAR','PKG','PARA','PH','PAYX','PAYC','PYPL','PNR','PEP','PFE','PCG',
'PM','PSX','PNW','PXD','PNC','POOL','PPG','PPL','PFG','PG','PGR','PLD','PRU','PEG','PTC','PSA','PHM',
'QRVO','PWR','QCOM','DGX','RL','RJF','RTX','O','REG','REGN','RF','RSG','RMD','RVTY','RHI','ROK','ROL',
'ROP','ROST','RCL','SPGI','CRM','SBAC','SLB','STX','SEE','SRE','NOW','SHW','SPG','SWKS','SJM','SNA',
'SEDG','SO','LUV','SWK','SBUX','STT','STLD','STE','SYK','SYF','SNPS','SYY','TMUS','TROW','TTWO','TPR',
'TRGP','TGT','TEL','TDY','TFX','TER','TSLA','TXN','TXT','TMO','TJX','TSCO','TT','TDG','TRV','TRMB','TFC',
'TYL','TSN','USB','UDR','ULTA','UNP','UAL','UPS','URI','UNH','UHS','VLO','VTR','VRSN','VRSK','VZ','VRTX',
'VFC','VTRS','VICI','V','VMC','WAB','WBA','WMT','WBD','WM','WAT','WEC','WFC','WELL','WST','WDC','WRK','WY',
'WHR','WMB','WTW','GWW','WYNN','XEL','XYL','YUM','ZBRA','ZBH','ZION','ZTS']

# Stocks of interest as per the problem statement - apple, amazon, google
target_tickers = ['AAPL', 'AMZN', 'GOOGL']

# Storing the total market capitalization of the entire S&P500 index (^GSPC)
marketCap_gspc = 0 # initialising, will be calculated later

# storing all the required metrics for the stocks of interests in a dictionary
# the elements in the data dictionary can be automated by dedicating a class for it.
# since the required metrics are very few, i'm going with the manual approach here
data = {
    "AAPL" : {
        'marketCap' : None,
        's&pweight' : None,
        'operationMargin' : None,
        'lastClose' : None,
        'ytdPerformance' : None,
        'revenue' : None,
        'netIncome' : None
    },
    "AMZN" : {
        'marketCap' : None,
        's&pweight' : None,
        'operationMargin' : None,
        'lastClose' : None,
        'ytdPerformance' : None,
        'revenue' : None,
        'netIncome' : None
    },
    "GOOGL" : {
        'marketCap' : None,
        's&pweight' : None,
        'operationMargin' : None,
        'lastClose' : None,
        'ytdPerformance' : None,
        'revenue' : None,
        'netIncome' : None
    }
}
# 1. starting off by calculating the net market capitalisation of s&p500 index
for ticker in tickers:
    try:
        marketCap = yf.Ticker(ticker).info['marketCap']
        # this marketCap represents the marketCap of the currently iterated stock
        marketCap_gspc += marketCap
        # updating the marketCaps of the targets to calculate the S&P weights ahead
        if(ticker in target_tickers):
            data[ticker]['marketCap'] = marketCap
    except:
        # we can also print the error received here, for now i'll just skip
        continue

# 2. Doing the Main calculation for all the metrics
for ticker in target_tickers:
    #fetching all the information of our target ticker in a variable(dicitionary) called stock
    stock = yf.Ticker(ticker).info
    
    # 1 S&P500 weight of the stock
    marketCap = data[ticker]['marketCap']
    weight = ((marketCap)/(marketCap_gspc))*100
    # #updating the weight of this stock
    data[ticker]['s&pweight'] = weight

    # 2 Operating Margin
    operatingMargin = stock['operatingMargins']
    data[ticker]['operatingMargin'] = operatingMargin

    # 3  Last Close Date
    lastClose = stock['previousClose']
    data[ticker]['lastClose'] = lastClose

    # 4 

    # 5 Stocks' YTD performance
    previousClose = stock['previousClose']
    currentPrice = stock['currentPrice']
    ytd = ((currentPrice - previousClose)/(previousClose))*100
    data[ticker]['ytdPerformance'] = ytd

    # 6 Revenue
    totalRevenue = stock['totalRevenue']
    fiftyTwoWeekChange = stock['52WeekChange']
    # revenue = ((totalRevenue - fiftyTwoWeekChange)/(fiftyTwoWeekChange))*100
    revenue = totalRevenue
    data[ticker]['revenue'] = revenue

    # 7 Net Income
    netIncomeToCommon = stock['netIncomeToCommon']
    fiftyTwoWeekLow = stock['fiftyTwoWeekLow']
    netIncome = netIncomeToCommon
    data[ticker]['netIncome']=netIncome

    # 8

# Printing all the Stock Data on the screen
for stock in data:
    print(stock)
    print("S&P 500 weight of the stock : %f " % data[stock]['s&pweight'])
    print("Operatin Margin : %f" % data[stock]['operatingMargin'])
    print("Last Closing Price : %f " % data[stock]['lastClose'])
    print("YTD Performance : %f " % data[stock]['ytdPerformance'])
    print("Annual Revenue : %d " % data[stock]['revenue'])
    print("Net Annual Income : %d " % data[stock]['netIncome'])
    print()
    
# Declare the file where you want the output to be stored here
output_file = "output.xlsx"
# Convert your data to a pandas Data Frame
df = pd.DataFrame(data)
# Write the Data to the output file
df.to_excel(output_file)

print(f"Output Saved to {output_file}")

