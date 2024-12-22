import pandas as pd
import talib

# Load the market data into a pandas DataFrame for easier TA calculations
market_data_file = 'market_data.xlsx'
df = pd.read_excel(market_data_file)

# Ensure the first column is 'Date' and convert it to datetime if it's not already
date_column = df.iloc[:, 0]
if not pd.api.types.is_datetime64_any_dtype(date_column):
    date_column = pd.to_datetime(date_column)

# Convert the 'Date' column to date-only strings to remove any time component
date_column = date_column.dt.strftime('%Y-%m-%d')

# Drop 'Date' from the processing DataFrame
df = df.iloc[:, 1:]  # Assume the first column is 'Date'

# Identify tickers by grouping every 6 columns (Open, High, Low, Close, Adj_Close, Volume)
ticker_groups = [df.columns[i:i + 6] for i in range(0, len(df.columns), 6)]

# Dictionary to hold the technical analysis data for each ticker
ta_data = {'Date': date_column}  # Add the 'Date' column first

# Process each ticker's OHLCV columns
for group_index, group in enumerate(ticker_groups):
    if len(group) != 6:
        print(f"Skipping group {group}: expected 6 columns, found {len(group)}.")
        continue

    # Extract the ticker name from the first column (e.g., 'GSPC_Open')
    try:
        ticker_name = group[0].split('_')[0]  # Extract ticker name (e.g., 'GSPC')
    except IndexError:
        print(f"Skipping group {group}: error extracting ticker name.")
        continue

    # Extract OHLCV data for this ticker
    ticker_df = df[group].copy()
    ticker_df.columns = ['Open', 'High', 'Low', 'Close', 'Adj_Close', 'Volume']  # Standardize the column names

    # Dictionary to store indicators for the current ticker
    indicators = {}

    # -- TA Indicators for the current ticker --

    # 1. Overlap Studies
    indicators[f'{ticker_name}_SMA_14'] = talib.SMA(ticker_df['Close'], timeperiod=14)
    indicators[f'{ticker_name}_EMA_14'] = talib.EMA(ticker_df['Close'], timeperiod=14)
    indicators[f'{ticker_name}_WMA_14'] = talib.WMA(ticker_df['Close'], timeperiod=14)
    indicators[f'{ticker_name}_DEMA_14'] = talib.DEMA(ticker_df['Close'], timeperiod=14)
    indicators[f'{ticker_name}_TEMA_14'] = talib.TEMA(ticker_df['Close'], timeperiod=14)
    indicators[f'{ticker_name}_TRIMA_14'] = talib.TRIMA(ticker_df['Close'], timeperiod=14)
    indicators[f'{ticker_name}_KAMA_14'] = talib.KAMA(ticker_df['Close'], timeperiod=14)
    indicators[f'{ticker_name}_MAMA'] = talib.MAMA(ticker_df['Close'])[0]  # MAMA and FAMA
    indicators[f'{ticker_name}_FAMA'] = talib.MAMA(ticker_df['Close'])[1]
    indicators[f'{ticker_name}_T3_14'] = talib.T3(ticker_df['Close'], timeperiod=14)
    indicators[f'{ticker_name}_MIDPOINT_14'] = talib.MIDPOINT(ticker_df['Close'], timeperiod=14)
    indicators[f'{ticker_name}_MIDPRICE_14'] = talib.MIDPRICE(ticker_df['High'], ticker_df['Low'], timeperiod=14)
    indicators[f'{ticker_name}_BB_upper'], indicators[f'{ticker_name}_BB_middle'], indicators[f'{ticker_name}_BB_lower'] = talib.BBANDS(
        ticker_df['Close'], timeperiod=20, nbdevup=2, nbdevdn=2, matype=0)

    # 2. Momentum Indicators
    indicators[f'{ticker_name}_ADX_14'] = talib.ADX(ticker_df['High'], ticker_df['Low'], ticker_df['Close'], timeperiod=14)
    indicators[f'{ticker_name}_ADXR_14'] = talib.ADXR(ticker_df['High'], ticker_df['Low'], ticker_df['Close'], timeperiod=14)
    indicators[f'{ticker_name}_APO'] = talib.APO(ticker_df['Close'])
    indicators[f'{ticker_name}_AROON_down'], indicators[f'{ticker_name}_AROON_up'] = talib.AROON(
        ticker_df['High'], ticker_df['Low'], timeperiod=14)
    indicators[f'{ticker_name}_AROONOSC'] = talib.AROONOSC(ticker_df['High'], ticker_df['Low'], timeperiod=14)
    indicators[f'{ticker_name}_BOP'] = talib.BOP(ticker_df['Open'], ticker_df['High'], ticker_df['Low'], ticker_df['Close'])
    indicators[f'{ticker_name}_CCI_14'] = talib.CCI(ticker_df['High'], ticker_df['Low'], ticker_df['Close'], timeperiod=14)
    indicators[f'{ticker_name}_CMO_14'] = talib.CMO(ticker_df['Close'], timeperiod=14)
    indicators[f'{ticker_name}_DX_14'] = talib.DX(ticker_df['High'], ticker_df['Low'], ticker_df['Close'], timeperiod=14)
    indicators[f'{ticker_name}_MACD'], indicators[f'{ticker_name}_MACD_signal'], indicators[
        f'{ticker_name}_MACD_hist'] = talib.MACD(ticker_df['Close'], fastperiod=12, slowperiod=26, signalperiod=9)
    indicators[f'{ticker_name}_MFI_14'] = talib.MFI(ticker_df['High'], ticker_df['Low'], ticker_df['Close'],
                                                    ticker_df['Volume'], timeperiod=14)
    indicators[f'{ticker_name}_MOM_14'] = talib.MOM(ticker_df['Close'], timeperiod=14)
    indicators[f'{ticker_name}_ROC_10'] = talib.ROC(ticker_df['Close'], timeperiod=10)
    indicators[f'{ticker_name}_ROCP_10'] = talib.ROCP(ticker_df['Close'], timeperiod=10)
    indicators[f'{ticker_name}_RSI_14'] = talib.RSI(ticker_df['Close'], timeperiod=14)
    indicators[f'{ticker_name}_TRIX_14'] = talib.TRIX(ticker_df['Close'], timeperiod=14)
    indicators[f'{ticker_name}_ULTOSC'] = talib.ULTOSC(ticker_df['High'], ticker_df['Low'], ticker_df['Close'],
                                                       timeperiod1=7, timeperiod2=14, timeperiod3=28)
    indicators[f'{ticker_name}_WILLR_14'] = talib.WILLR(ticker_df['High'], ticker_df['Low'], ticker_df['Close'],
                                                        timeperiod=14)

    # 3. Volume Indicators
    indicators[f'{ticker_name}_AD'] = talib.AD(ticker_df['High'], ticker_df['Low'], ticker_df['Close'],
                                               ticker_df['Volume'])
    indicators[f'{ticker_name}_ADOSC'] = talib.ADOSC(ticker_df['High'], ticker_df['Low'], ticker_df['Close'],
                                                     ticker_df['Volume'], fastperiod=3, slowperiod=10)
    indicators[f'{ticker_name}_OBV'] = talib.OBV(ticker_df['Close'], ticker_df['Volume'])

    # 4. Volatility Indicators
    indicators[f'{ticker_name}_ATR_14'] = talib.ATR(ticker_df['High'], ticker_df['Low'], ticker_df['Close'],
                                                    timeperiod=14)
    indicators[f'{ticker_name}_NATR_14'] = talib.NATR(ticker_df['High'], ticker_df['Low'], ticker_df['Close'],
                                                      timeperiod=14)
    indicators[f'{ticker_name}_TRANGE'] = talib.TRANGE(ticker_df['High'], ticker_df['Low'], ticker_df['Close'])

    # 5. Cycle Indicators (Hilbert Transform)
    indicators[f'{ticker_name}_HT_DCPERIOD'] = talib.HT_DCPERIOD(ticker_df['Close'])
    indicators[f'{ticker_name}_HT_DCPHASE'] = talib.HT_DCPHASE(ticker_df['Close'])
    indicators[f'{ticker_name}_HT_PHASOR_inphase'], indicators[f'{ticker_name}_HT_PHASOR_quadrature'] = talib.HT_PHASOR(
        ticker_df['Close'])
    indicators[f'{ticker_name}_HT_TRENDMODE'] = talib.HT_TRENDMODE(ticker_df['Close'])
    indicators[f'{ticker_name}_HT_TRENDLINE'] = talib.HT_TRENDLINE(ticker_df['Close'])

    # 6. Price Transformations
    indicators[f'{ticker_name}_AVGPRICE'] = talib.AVGPRICE(ticker_df['Open'], ticker_df['High'],
                                                           ticker_df['Low'], ticker_df['Close'])
    indicators[f'{ticker_name}_MEDPRICE'] = talib.MEDPRICE(ticker_df['High'], ticker_df['Low'])
    indicators[f'{ticker_name}_TYPPRICE'] = talib.TYPPRICE(ticker_df['High'], ticker_df['Low'],
                                                           ticker_df['Close'])
    indicators[f'{ticker_name}_WCLPRICE'] = talib.WCLPRICE(ticker_df['High'], ticker_df['Low'],
                                                           ticker_df['Close'])

    # 7. Pattern Recognition Indicators (Candlestick Patterns)
    candlestick_patterns = {
        'CDL2CROWS': talib.CDL2CROWS,
        'CDL3BLACKCROWS': talib.CDL3BLACKCROWS,
        'CDL3INSIDE': talib.CDL3INSIDE,
        'CDL3LINESTRIKE': talib.CDL3LINESTRIKE,
        'CDL3OUTSIDE': talib.CDL3OUTSIDE,
        'CDL3STARSINSOUTH': talib.CDL3STARSINSOUTH,
        'CDL3WHITESOLDIERS': talib.CDL3WHITESOLDIERS,
        'CDLABANDONEDBABY': talib.CDLABANDONEDBABY,
        'CDLADVANCEBLOCK': talib.CDLADVANCEBLOCK,
        'CDLBELTHOLD': talib.CDLBELTHOLD,
        'CDLBREAKAWAY': talib.CDLBREAKAWAY,
        'CDLCLOSINGMARUBOZU': talib.CDLCLOSINGMARUBOZU,
        'CDLCONCEALBABYSWALL': talib.CDLCONCEALBABYSWALL,
        'CDLCOUNTERATTACK': talib.CDLCOUNTERATTACK,
        'CDLDARKCLOUDCOVER': talib.CDLDARKCLOUDCOVER,
        'CDLDOJI': talib.CDLDOJI,
        'CDLDOJISTAR': talib.CDLDOJISTAR,
        'CDLDRAGONFLYDOJI': talib.CDLDRAGONFLYDOJI,
        'CDLENGULFING': talib.CDLENGULFING,
        'CDLEVENINGDOJISTAR': talib.CDLEVENINGDOJISTAR,
        'CDLEVENINGSTAR': talib.CDLEVENINGSTAR,
        'CDLGAPSIDESIDEWHITE': talib.CDLGAPSIDESIDEWHITE,
        'CDLGRAVESTONEDOJI': talib.CDLGRAVESTONEDOJI,
        'CDLHAMMER': talib.CDLHAMMER,
        'CDLHANGINGMAN': talib.CDLHANGINGMAN,
        'CDLHARAMI': talib.CDLHARAMI,
        'CDLHARAMICROSS': talib.CDLHARAMICROSS,
        'CDLHIGHWAVE': talib.CDLHIGHWAVE,
        'CDLHIKKAKE': talib.CDLHIKKAKE,
        'CDLHIKKAKEMOD': talib.CDLHIKKAKEMOD,
        'CDLHOMINGPIGEON': talib.CDLHOMINGPIGEON,
        'CDLIDENTICAL3CROWS': talib.CDLIDENTICAL3CROWS,
        'CDLINNECK': talib.CDLINNECK,
        'CDLINVERTEDHAMMER': talib.CDLINVERTEDHAMMER,
        'CDLKICKING': talib.CDLKICKING,
        'CDLKICKINGBYLENGTH': talib.CDLKICKINGBYLENGTH,
        'CDLLADDERBOTTOM': talib.CDLLADDERBOTTOM,
        'CDLLONGLEGGEDDOJI': talib.CDLLONGLEGGEDDOJI,
        'CDLLONGLINE': talib.CDLLONGLINE,
        'CDLMARUBOZU': talib.CDLMARUBOZU,
        'CDLMATCHINGLOW': talib.CDLMATCHINGLOW,
        'CDLMATHOLD': talib.CDLMATHOLD,
        'CDLMORNINGDOJISTAR': talib.CDLMORNINGDOJISTAR,
        'CDLMORNINGSTAR': talib.CDLMORNINGSTAR,
        'CDLONNECK': talib.CDLONNECK,
        'CDLPIERCING': talib.CDLPIERCING,
        'CDLRICKSHAWMAN': talib.CDLRICKSHAWMAN,
        'CDLRISEFALL3METHODS': talib.CDLRISEFALL3METHODS,
        'CDLSEPARATINGLINES': talib.CDLSEPARATINGLINES,
        'CDLSHOOTINGSTAR': talib.CDLSHOOTINGSTAR,
        'CDLSHORTLINE': talib.CDLSHORTLINE,
        'CDLSPINNINGTOP': talib.CDLSPINNINGTOP,
        'CDLSTALLEDPATTERN': talib.CDLSTALLEDPATTERN,
        'CDLSTICKSANDWICH': talib.CDLSTICKSANDWICH,
        'CDLTAKURI': talib.CDLTAKURI,
        'CDLTASUKIGAP': talib.CDLTASUKIGAP,
        'CDLTHRUSTING': talib.CDLTHRUSTING,
        'CDLTRISTAR': talib.CDLTRISTAR,
        'CDLUNIQUE3RIVER': talib.CDLUNIQUE3RIVER,
        'CDLUPSIDEGAP2CROWS': talib.CDLUPSIDEGAP2CROWS,
        'CDLXSIDEGAP3METHODS': talib.CDLXSIDEGAP3METHODS
    }

    # Apply all candlestick pattern recognition indicators
    for pattern_name, pattern_func in candlestick_patterns.items():
        indicators[f'{ticker_name}_{pattern_name}'] = pattern_func(
            ticker_df['Open'], ticker_df['High'], ticker_df['Low'], ticker_df['Close'])

    # Add computed indicators to the result dataframe with correct column labels
    for key, value in indicators.items():
        ta_data[key] = value

# Combine the 'Date' column and the computed TA indicators into one DataFrame
combined_df = pd.DataFrame(ta_data)

# Optional: If you want to ensure that all datetime objects in the DataFrame are date-only strings
# This step is already handled by converting 'Date' to strings earlier

# Save the result to a new Excel file
output_file = 'ta_data.xlsx'
combined_df.to_excel(output_file, index=False)

print(f"Technical analysis data has been saved to {output_file}.")


