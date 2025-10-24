#
# File: market_scanner.py
# Description: Module for scanning and curating exchanges and markets.
#

import pandas as pd
from exchange_manager import ExchangeManager
import logging

def analyze_exchange_markets(exchange_id: str, exchange_instance) -> dict:
    """
    Analyzes an exchange's markets to gather key metrics.

    Args:
        exchange_id (str): The ID of the exchange.
        exchange_instance: The initialized ccxt exchange object.

    Returns:
        dict: A dictionary containing analysis results for the exchange.
    """
    results = {
        'name': exchange_id,
        'total_spot_pairs': 0,
        'usdt_quoted_pairs': 0,
        'supports_fetchOHLCV': False,
        'rate_limit_ms': exchange_instance.rateLimit
    }

    try:
        # Load all available markets from the exchange
        markets = exchange_instance.load_markets()

        # Filter for active spot markets
        spot_markets = {k: v for k, v in markets.items() if v.get('spot', False) and v.get('active', True)}
        results['total_spot_pairs'] = len(spot_markets)

        # Count pairs quoted in USDT
        usdt_pairs = [market for market in spot_markets.values() if market['quote'] == 'USDT']
        results['usdt_quoted_pairs'] = len(usdt_pairs)

        # Check for essential data fetching capability
        if exchange_instance.has.get('fetchOHLCV'):
            results['supports_fetchOHLCV'] = True

    except Exception as e:
        logging.error(f"Could not analyze markets for '{exchange_id}': {e}")
        # Return default results on error to avoid crashing the whole scan

    return results
