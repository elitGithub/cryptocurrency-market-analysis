"""
Exchange Manager - Handles cryptocurrency exchange connections and operations.
"""
import ccxt
import logging
from typing import Dict, Optional


class ExchangeManager:
    """Manages connections to multiple cryptocurrency exchanges."""
    
    def __init__(self, exchange_ids: list):
        """
        Initialize exchange manager with a list of exchange IDs.
        
        Args:
            exchange_ids: List of exchange identifiers (e.g., ['binance', 'kucoin'])
        """
        self.exchange_ids = exchange_ids
        self.exchanges: Dict[str, ccxt.Exchange] = {}
        self._initialize_exchanges()
    
    def _initialize_exchanges(self) -> None:
        """Initialize ccxt exchange objects with rate limiting enabled."""
        logging.info(f"Initializing {len(self.exchange_ids)} exchanges...")
        
        for exchange_id in self.exchange_ids:
            if exchange_id not in ccxt.exchanges:
                logging.warning(f"Exchange '{exchange_id}' not supported by ccxt. Skipping.")
                continue
            
            try:
                exchange_class = getattr(ccxt, exchange_id)
                config = {
                    'enableRateLimit': True,
                    'options': {'defaultType': 'spot'},
                }
                
                self.exchanges[exchange_id] = exchange_class(config)
                logging.info(f"✓ Initialized '{exchange_id}'")
                
            except Exception as e:
                logging.error(f"✗ Failed to initialize '{exchange_id}': {e}")
    
    def get_exchange(self, exchange_id: str) -> Optional[ccxt.Exchange]:
        """
        Get a specific exchange instance.
        
        Args:
            exchange_id: Exchange identifier
            
        Returns:
            Exchange instance or None if not found
        """
        return self.exchanges.get(exchange_id)
    
    def get_all_exchanges(self) -> Dict[str, ccxt.Exchange]:
        """Get all initialized exchange instances."""
        return self.exchanges
    
    def get_exchange_count(self) -> int:
        """Get count of successfully initialized exchanges."""
        return len(self.exchanges)