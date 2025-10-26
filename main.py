#!/usr/bin/env python3
"""
Crypto Trading Bot - Main Entry Point
Analyzes cryptocurrency markets and generates comprehensive reports.
"""
import configparser
import logging
import sys
import os

# Add crypto_bot to path
sys.path.insert(0, os.path.dirname(__file__))

from crypto_bot import CryptoAnalyzer


def setup_logging():
    """Configure logging for the application."""
    logging.basicConfig(
        level=logging.INFO,
        format='%(asctime)s - %(levelname)s - %(message)s',
        handlers=[
            logging.StreamHandler(sys.stdout)
        ]
    )


def load_config():
    """Load configuration from config.ini file."""
    config = configparser.ConfigParser()

    if not config.read('config.ini'):
        logging.error("Failed to read config.ini. Make sure the file exists.")
        sys.exit(1)

    try:
        # Parse configuration
        target_exchanges = [
            e.strip() for e in config.get('exchanges', 'target_exchanges', fallback='').split(',')
            if e.strip()
        ]

        symbol = config.get('analysis', 'symbol', fallback='BTC/USDT')
        timeframe = config.get('analysis', 'timeframe', fallback='1d')
        history_days = config.getint('analysis', 'history_days', fallback=730)
        short_ma = config.getint('analysis', 'short_ma', fallback=50)
        long_ma = config.getint('analysis', 'long_ma', fallback=200)
        output_dir = config.get('paths', 'output_dir', fallback='output')

        if not target_exchanges:
            logging.error("No target_exchanges specified in config.ini")
            sys.exit(1)

        return {
            'exchange_ids': target_exchanges,
            'symbol': symbol,
            'timeframe': timeframe,
            'history_days': history_days,
            'short_ma': short_ma,
            'long_ma': long_ma,
            'output_dir': output_dir
        }

    except (configparser.NoSectionError, configparser.NoOptionError, ValueError) as e:
        logging.error(f"Configuration error in config.ini: {e}")
        sys.exit(1)


def main():
    """Main entry point."""
    setup_logging()

    print("\n" + "=" * 70)
    print("  CRYPTOCURRENCY TRADING BOT - MARKET ANALYSIS TOOL")
    print("=" * 70 + "\n")

    # Load configuration
    config = load_config()

    # Create analyzer
    analyzer = CryptoAnalyzer(**config)

    # Run analysis
    success = analyzer.run()

    # Exit with appropriate code
    sys.exit(0 if success else 1)


if __name__ == '__main__':
    main()