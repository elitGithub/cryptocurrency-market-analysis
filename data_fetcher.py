#
# File: data_fetcher.py
# Description: Module for fetching comprehensive historical data.
#

import pandas as pd
from datetime import datetime, timedelta
import time
import logging

def fetch_full_ohlcv(exchange, symbol: str, timeframe: str, since_days: int) -> list:
    """
    Fetches the complete historical OHLCV data for a symbol by handling pagination.

    Args:
        exchange: The initialized ccxt exchange object.
        symbol (str): The trading symbol (e.g., 'BTC/USDT').
        timeframe (str): The timeframe (e.g., '1d', '4h').
        since_days (int): The number of days of historical data to fetch.

    Returns:
        list: A list of lists containing the full OHLCV data.
    """
    if not exchange.has['fetchOHLCV']:
        logging.error(f"Exchange '{exchange.id}' does not support fetchOHLCV.")
        return []

    logging.info(f"Fetching full historical data for {symbol} on {exchange.id}...")

    # Calculate the start timestamp
    since_datetime = datetime.utcnow() - timedelta(days=since_days)
    since_timestamp = int(since_datetime.timestamp() * 1000)

    all_ohlcv = []
    limit = 1000 # Set to a high value, exchange will cap it if necessary

    while True:
        try:
            # Fetch a batch of data
            ohlcv = exchange.fetch_ohlcv(symbol, timeframe, since=since_timestamp, limit=limit)

            if not ohlcv:
                # No more data to fetch
                break

            all_ohlcv.extend(ohlcv)

            # Update the 'since' timestamp to the timestamp of the last candle fetched
            # This is the key to pagination
            since_timestamp = ohlcv[-1][0] + 1 # +1 to avoid fetching the last candle again

            logging.info(f"Fetched {len(ohlcv)} candles. Total: {len(all_ohlcv)}. Continuing from {datetime.fromtimestamp(since_timestamp/1000)}")

        except Exception as e:
            logging.error(f"An error occurred while fetching OHLCV data: {e}")
            break

    logging.info(f"Finished fetching data. Total candles retrieved: {len(all_ohlcv)}")
    return all_ohlcv

def ohlcv_to_dataframe(ohlcv: list) -> pd.DataFrame:
    """
    Converts a list of OHLCV data into a cleaned pandas DataFrame.

    Args:
        ohlcv (list): The list of lists from fetch_full_ohlcv.

    Returns:
        pd.DataFrame: A DataFrame with a datetime index and appropriate columns.
    """
    if not ohlcv:
        return pd.DataFrame()

    df = pd.DataFrame(ohlcv, columns=['timestamp', 'open', 'high', 'low', 'close', 'volume'])

    # Convert timestamp to datetime and set as index
    df['timestamp'] = pd.to_datetime(df['timestamp'], unit='ms')
    df.set_index('timestamp', inplace=True)

    # Ensure data types are correct
    df = df.astype(float)

    # Remove potential duplicate rows that can occur during pagination
    df = df[~df.index.duplicated(keep='first')]

    # Sort by index to ensure chronological order
    df.sort_index(inplace=True)

    return df
