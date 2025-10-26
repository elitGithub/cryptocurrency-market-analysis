"""Data Fetcher - Handles fetching and processing historical OHLCV data."""
import pandas as pd
from datetime import datetime, timedelta
import logging
from typing import List, Optional


class DataFetcher:
    """Fetches and processes cryptocurrency market data."""
    
    def __init__(self, exchange):
        self.exchange = exchange
        self.exchange_id = exchange.id if hasattr(exchange, 'id') else 'unknown'
    
    def fetch_ohlcv(self, symbol: str, timeframe: str, days: int) -> List:
        """Fetch complete historical OHLCV data with pagination."""
        if not self.exchange.has['fetchOHLCV']:
            logging.error(f"Exchange '{self.exchange_id}' doesn't support fetchOHLCV")
            return []
        
        logging.info(f"Fetching {days} days of {symbol} data from {self.exchange_id}...")
        
        since_datetime = datetime.utcnow() - timedelta(days=days)
        since_timestamp = int(since_datetime.timestamp() * 1000)
        
        all_ohlcv = []
        limit = 1000
        
        while True:
            try:
                ohlcv = self.exchange.fetch_ohlcv(
                    symbol, timeframe, 
                    since=since_timestamp, 
                    limit=limit
                )
                
                if not ohlcv:
                    break
                
                all_ohlcv.extend(ohlcv)
                since_timestamp = ohlcv[-1][0] + 1
                
                logging.info(f"Fetched {len(ohlcv)} candles (total: {len(all_ohlcv)})")
                
            except Exception as e:
                logging.error(f"Error fetching OHLCV data: {e}")
                break
        
        logging.info(f"✓ Fetched {len(all_ohlcv)} total candles")
        return all_ohlcv
    
    @staticmethod
    def to_dataframe(ohlcv: List) -> Optional[pd.DataFrame]:
        """Convert OHLCV list to pandas DataFrame."""
        if not ohlcv:
            logging.warning("Empty OHLCV data provided")
            return None
        
        df = pd.DataFrame(
            ohlcv, 
            columns=['timestamp', 'open', 'high', 'low', 'close', 'volume']
        )
        
        df['timestamp'] = pd.to_datetime(df['timestamp'], unit='ms')
        df.set_index('timestamp', inplace=True)
        df = df.astype(float)
        df = df[~df.index.duplicated(keep='first')]
        df.sort_index(inplace=True)
        
        return df
