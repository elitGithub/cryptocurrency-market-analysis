"""Market Scanner - Analyzes exchange markets and gathers metrics."""
import logging
from typing import Dict
import pandas as pd


class MarketScanner:
    """Scans and analyzes cryptocurrency exchange markets."""
    
    @staticmethod
    def analyze_exchange(exchange_id: str, exchange_instance) -> Dict:
        """Analyze an exchange's markets to gather key metrics."""
        results = {
            'name': exchange_id,
            'total_spot_pairs': 0,
            'usdt_quoted_pairs': 0,
            'supports_fetchOHLCV': False,
            'rate_limit_ms': exchange_instance.rateLimit
        }
        
        try:
            markets = exchange_instance.load_markets()
            
            spot_markets = {
                k: v for k, v in markets.items() 
                if v.get('spot', False) and v.get('active', True)
            }
            results['total_spot_pairs'] = len(spot_markets)
            
            usdt_pairs = [m for m in spot_markets.values() if m['quote'] == 'USDT']
            results['usdt_quoted_pairs'] = len(usdt_pairs)
            
            if exchange_instance.has.get('fetchOHLCV'):
                results['supports_fetchOHLCV'] = True
                
        except Exception as e:
            logging.error(f"Failed to analyze '{exchange_id}': {e}")
        
        return results
    
    @staticmethod
    def scan_all_exchanges(exchanges: Dict) -> pd.DataFrame:
        """Scan all exchanges and return results as DataFrame."""
        results = []
        
        for exchange_id, exchange_instance in exchanges.items():
            logging.info(f"Scanning {exchange_id}...")
            result = MarketScanner.analyze_exchange(exchange_id, exchange_instance)
            results.append(result)
        
        return pd.DataFrame(results)
