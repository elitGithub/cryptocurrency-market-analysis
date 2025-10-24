#
# File: exchange_manager.py
# Description: Foundational module for managing exchange connections.
#

import ccxt
import time
import logging

# Configure basic logging
logging.basicConfig(level=logging.INFO, format='%(asctime)s - %(levelname)s - %(message)s')

def list_available_exchanges():
    """Prints all exchange IDs supported by the ccxt library."""
    print("Supported Exchanges by ccxt:")
    print(", ".join(ccxt.exchanges))

class ExchangeManager:
    """
    A centralized class to manage connections to multiple cryptocurrency exchanges using ccxt.
    """
    def __init__(self, exchange_ids: list):
        """
        Initializes the ExchangeManager with a list of exchange IDs.

        Args:
            exchange_ids (list): A list of string identifiers for the exchanges to connect to.
        """
        self.exchange_ids = exchange_ids
        self.exchanges = {}
        self._initialize_exchanges()

    def _initialize_exchanges(self):
        """
        Private method to instantiate ccxt exchange objects for the provided IDs.
        It includes crucial configuration for rate limiting.
        """
        logging.info(f"Initializing connections for: {', '.join(self.exchange_ids)}")
        for exchange_id in self.exchange_ids:
            if exchange_id not in ccxt.exchanges:
                logging.warning(f"Exchange '{exchange_id}' is not supported by ccxt. Skipping.")
                continue
            try:
                # getattr(ccxt, exchange_id) is equivalent to ccxt.binance(), ccxt.kraken(), etc.
                exchange_class = getattr(ccxt, exchange_id)

                # Critical configuration: enable the built-in rate limiter.
                config = {
                    'enableRateLimit': True,
                    'options': {
                        'defaultType': 'spot',
                    },
                }

                self.exchanges[exchange_id] = exchange_class(config)
                logging.info(f"Successfully initialized connection for '{exchange_id}'.")
            except Exception as e:
                logging.error(f"Failed to initialize exchange '{exchange_id}': {e}")

    def get_exchange(self, exchange_id: str):
        """
        Retrieves an initialized exchange instance.

        Args:
            exchange_id (str): The ID of the exchange to retrieve.

        Returns:
            An initialized ccxt exchange object, or None if not found.
        """
        return self.exchanges.get(exchange_id)

    def get_all_exchanges(self) -> dict:
        """
        Returns the dictionary of all successfully initialized exchange instances.

        Returns:
            dict: A dictionary mapping exchange IDs to their ccxt objects.
        """
        return self.exchanges