import datetime as dt
import pandas as pd
from pandas_datareader import data as pdr
import yfinance as yf
import numpy as np
from scipy import stats
import matplotlib
matplotlib.use('Agg')
import matplotlib.pyplot as plt
from matplotlib.backends.backend_pdf import PdfPages

colors = {'red': '#ff207c', 'grey': '#42535b', 'blue': '#207cff', 'orange': '#ffa320', 'green': '#00ec8b'}
config_ticks = {'size': 14, 'color': colors['grey'], 'labelcolor': colors['grey']}
config_title = {'size': 18, 'color': colors['grey'], 'ha': 'left', 'va': 'baseline'}

tickers = []
tickers_name = []

yf.pdr_override()
start =dt.datetime(2019,12,1)
now = dt.datetime.now()
exportList= pd.DataFrame(columns=['Stock', "Name", "RS_Rating","Percentile", "50 Day MA", "150 Day Ma", "200 Day MA", "52 week High", "52 week Low"])

stockdata = pd.read_excel("Hong_Kong_Top_200_MarketCap_Stocks_List.xlsx", engine="openpyxl")
stockdetails = pd.DataFrame(columns=['Stock', "RS_rating", "moving_average_50", "moving_average_150", "moving_average_200", "low_of_52week", "high_of_52week","currentClose", "moving_average_200_20",'Name'])

for i in stockdata.index:
    if str(stockdata["Tickers"][i]) != "nan":
        tickers.append(str(stockdata["Tickers"][i]))
    if str(stockdata["Name"][i]) != "nan":
        tickers_name.append(str(stockdata["Name"][i]))

# Getting stocks data
for i in range(len(tickers)):
     try:
        df = pdr.get_data_yahoo(tickers[i], start, now)
        smaUsed = [50,150,200]
        for x in smaUsed:
            sma = x
            df["SMA_"+str(sma)]=round(df.iloc[:,4].rolling(window=sma).mean(),2)

        currentClose=df["Adj Close"][-1]
        moving_average_50 = df["SMA_50"][-1]
        moving_average_150 = df["SMA_150"][-1]
        moving_average_200 = df["SMA_200"][-1]
        low_of_52week = min(df["Adj Close"][-260:])
        high_of_52week = max(df["Adj Close"][-260:])

        try:
            RS_rating = ((((currentClose - df["Adj Close"][-63])/df["Adj Close"][-63]) * 0.4) + (((currentClose - df["Adj Close"][-126]) /df["Adj Close"][-126]) * 0.2) + (((currentClose - df["Adj Close"][-189]) / df["Adj Close"][-189]) * 0.2) + (((currentClose - df["Adj Close"][-252]) / df["Adj Close"][-252]) * 0.2)) * 100
        except:
            RS_rating = 0

        try:
            moving_average_200_20 = df["SMA_200"][-20]
        except:
            moving_average_200_20 = 0


        print("Getting data for " + tickers[i] + "...")

        stockdetails = stockdetails.append({'Stock':tickers[i], "RS_rating":RS_rating,"moving_average_50": moving_average_50, "moving_average_150": moving_average_150, "moving_average_200": moving_average_200, "low_of_52week": low_of_52week, "high_of_52week": high_of_52week, "currentClose": currentClose, "moving_average_200_20": moving_average_200_20,"Name": tickers_name[i]}, ignore_index=True)

     except:
        print("No data on " + tickers[i])






# for testing only
# stockdetails = pd.read_excel("stockdetails.xlsx", engine="openpyxl")

RS_rating_requirement = np.percentile(stockdetails["RS_rating"].tolist(), 70)
print(RS_rating_requirement)
print(stockdetails)

# checking conditions
for row in stockdetails.iterrows():
    # Condition 1: Current Price > 150 SMA and > 200 SMA
    if(row[1][7]>row[1][3]>row[1][4]):
        cond_1 = True
    else:
        cond_1 = False

    # Condition 2: 150 SMA and > 200 SMA
    if(row[1][3]>row[1][4]):
        cond_2 = True
    else:
        cond_2 = False

    # Condition 3 200 SMA trending up for at least 1 month (ideally 4-5 months)
    if(row[1][4]>row[1][8]):
        cond_3 = True
    else:
        cond_3 = False

    # Condition 4: 50 SMA> 150 SMA and 50 SMA> 200 SMA
    if(row[1][2]>row[1][3] and row[1][2]>row[1][4]):
        cond_4 = True
    else:
        cond_4 = False

    # Condition 5: Current Price > 50 SMA
    if(row[1][7]>row[1][2]):
        cond_5 = True
    else:
        cond_5 = False

    # Condition 6: Current Price is at least 30% above 52 week low
    if(row[1][7]>=(1.3*row[1][5])):
        cond_6 = True
    else:
        cond_6 = False

    # Condition 7: Current Price is within 25% of 52 week high
    if(row[1][7]>=(.75*row[1][6])):
        cond_7 = True
    else:
        cond_7 = False

    if(row[1][1] >= RS_rating_requirement):
        cond_8 = True
    else:
        cond_8 = False

    if (cond_1 and cond_2 and cond_3 and cond_4 and cond_5 and cond_6 and cond_7 and cond_8):
        percentile_score = stats.percentileofscore(stockdetails["RS_rating"].tolist(), row[1][1])
        exportList = exportList.append(
            {'Stock': row[1][0], "Name": row[1][9], "RS_Rating": row[1][1],"Percentile": percentile_score, "50 Day MA": row[1][2], "150 Day Ma": row[1][3],
             "200 Day MA": row[1][4], "52 week Low": row[1][5], "52 week High": row[1][6]},
            ignore_index=True)

exportList = exportList.sort_values("Percentile", ascending=False)


with PdfPages('Charts.pdf') as export_pdf:

    # Plotting the graphs
    for i in exportList.index:
        df = pdr.get_data_yahoo(str(exportList["Stock"][i]), start, now)
        smaUsed = [50, 150, 200]
        for x in smaUsed:
            sma = x
            df["SMA_" + str(sma)] = round(df.iloc[:, 4].rolling(window=sma).mean(), 2)
        df = df.tail(200)
        df.reset_index(inplace=True)
        print(df.dtypes)
        plt.rc('figure', figsize=(15,10))
        fig, axes = plt.subplots(2,1,
                                gridspec_kw=
                                {'height_ratios': [3, 1]})
        fig.tight_layout(pad=3)

        date = df["Date"]
        close = df["Adj Close"]
        vol = df["Volume"]

        plot_price = axes[0]
        plot_price.plot(date, close, color=colors['blue'],
                        linewidth=2, label='Price')

        plot_vol = axes[1]
        plot_vol.bar(date, vol, width=2, color='darkgrey')

        # fig.grid(True)
        fig.suptitle(str(exportList["Stock"][i]))
        export_pdf.savefig()
        plt.close()





writer = pd.ExcelWriter('exportList.xlsx')
exportList.to_excel(writer, index=False)
writer.save()
