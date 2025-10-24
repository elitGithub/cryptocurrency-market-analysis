#
# File: demo_report_generator.py
# Description: Demo script to generate sample reports without needing live exchange data
#

import pandas as pd
import numpy as np
from datetime import datetime, timedelta
from report_generator import determine_trading_signal, generate_word_report, generate_powerpoint_report
import logging

logging.basicConfig(level=logging.INFO, format='%(asctime)s - %(levelname)s - %(message)s')

def generate_sample_data():
    """Generate sample OHLCV data with indicators for demonstration"""
    
    # Generate 400 days of sample data
    dates = pd.date_range(end=datetime.now(), periods=400, freq='D')
    
    # Simulate a trending price with some volatility
    base_price = 30000
    trend = np.linspace(0, 20000, 400)  # Upward trend
    volatility = np.random.normal(0, 1000, 400).cumsum()
    
    close_prices = base_price + trend + volatility
    
    # Generate OHLC data
    df = pd.DataFrame({
        'close': close_prices,
        'open': close_prices * (1 + np.random.uniform(-0.02, 0.02, 400)),
        'high': close_prices * (1 + np.random.uniform(0, 0.03, 400)),
        'low': close_prices * (1 + np.random.uniform(-0.03, 0, 400)),
        'volume': np.random.uniform(1000000, 5000000, 400)
    }, index=dates)
    
    # Calculate indicators
    df['SMA_short'] = df['close'].rolling(window=50).mean()
    df['SMA_long'] = df['close'].rolling(window=200).mean()
    df['RSI'] = 50 + np.random.uniform(-20, 20, 400)  # Simplified RSI
    
    # Bollinger Bands
    bb_window = 20
    rolling_mean = df['close'].rolling(window=bb_window).mean()
    rolling_std = df['close'].rolling(window=bb_window).std()
    df['BB_middle'] = rolling_mean
    df['BB_upper'] = rolling_mean + (rolling_std * 2)
    df['BB_lower'] = rolling_mean - (rolling_std * 2)
    
    # Drop NaN values
    df = df.dropna()
    
    return df

def generate_sample_exchange_results():
    """Generate sample exchange comparison data"""
    return [
        {
            'name': 'binance',
            'total_spot_pairs': 2147,
            'usdt_quoted_pairs': 412,
            'supports_fetchOHLCV': True,
            'rate_limit_ms': 50
        },
        {
            'name': 'kucoin',
            'total_spot_pairs': 1543,
            'usdt_quoted_pairs': 287,
            'supports_fetchOHLCV': True,
            'rate_limit_ms': 10
        },
        {
            'name': 'bybit',
            'total_spot_pairs': 891,
            'usdt_quoted_pairs': 203,
            'supports_fetchOHLCV': True,
            'rate_limit_ms': 20
        },
        {
            'name': 'gate',
            'total_spot_pairs': 2234,
            'usdt_quoted_pairs': 456,
            'supports_fetchOHLCV': True,
            'rate_limit_ms': 50
        }
    ]

def generate_sample_suggestions(signal):
    """Generate sample analysis suggestions"""
    suggestions = [
        f"{signal['action']} signal detected with {signal['confidence']} confidence",
        f"Current trend is {signal['trend']} based on moving average analysis",
        f"RSI indicator shows {signal['rsi']:.1f}, which is in the " + 
        ('overbought' if signal['rsi'] > 70 else 'oversold' if signal['rsi'] < 30 else 'neutral') + " zone"
    ]
    
    if 'BUY' in signal['action']:
        suggestions.append("BULLISH SIGNAL: Technical indicators suggest potential for upward price movement")
    elif 'SELL' in signal['action']:
        suggestions.append("BEARISH SIGNAL: Technical indicators suggest potential for downward price movement")
    else:
        suggestions.append("NEUTRAL SIGNAL: No strong directional bias detected in current market conditions")
    
    return suggestions

def main():
    """Generate demo reports"""
    
    print("=" * 70)
    print("CRYPTOCURRENCY TRADING BOT - DEMO REPORT GENERATOR")
    print("=" * 70)
    print("\nGenerating sample data and reports...")
    print("-" * 70)
    
    # Configuration
    symbol = "BTC/USDT"
    short_ma = 50
    long_ma = 200
    
    # Generate sample data
    logging.info("Creating sample market data...")
    df = generate_sample_data()
    exchange_results = generate_sample_exchange_results()
    
    # Determine trading signal
    latest_price = df.iloc[-1]['close']
    signal = determine_trading_signal([], df, latest_price)
    
    # Generate suggestions
    suggestions = generate_sample_suggestions(signal)
    
    # Display signal to console
    print(f"\n{'=' * 70}")
    print(f"TRADING SIGNAL: {signal['action']}")
    print(f"{'=' * 70}")
    print(f"Confidence:  {signal['confidence']}")
    print(f"Price:       ${signal['price']:.2f}")
    print(f"Trend:       {signal['trend']}")
    print(f"RSI:         {signal['rsi']:.1f}")
    print(f"\nReasoning:")
    print(f"  {signal['reasoning']}")
    print(f"{'=' * 70}\n")
    
    # Generate reports
    logging.info("Generating Word document...")
    generate_word_report(symbol, exchange_results, signal, suggestions, df, short_ma, long_ma)
    
    logging.info("Generating PowerPoint presentation...")
    generate_powerpoint_report(symbol, exchange_results, signal, suggestions, df, short_ma, long_ma)
    
    # Move files to outputs directory
    import shutil
    import os
    
    os.makedirs('/mnt/user-data/outputs', exist_ok=True)
    
    if os.path.exists('crypto_analysis_report.docx'):
        shutil.move('crypto_analysis_report.docx', 
                   '/mnt/user-data/outputs/crypto_analysis_report.docx')
        logging.info("Word document moved to outputs directory")
    
    if os.path.exists('crypto_analysis_presentation.pptx'):
        shutil.move('crypto_analysis_presentation.pptx',
                   '/mnt/user-data/outputs/crypto_analysis_presentation.pptx')
        logging.info("PowerPoint presentation moved to outputs directory")
    
    print("\n" + "=" * 70)
    print("✅ REPORTS GENERATED SUCCESSFULLY!")
    print("=" * 70)
    print("\nYour reports are ready:")
    print("  • PowerPoint Presentation")
    print("  • Word Document")
    print("\nEach report includes:")
    print("  ✓ Clear BUY/SELL/HOLD recommendation")
    print("  ✓ Exchange comparison analysis")
    print("  ✓ Technical analysis details")
    print("  ✓ Educational section explaining indicators")
    print("  ✓ Important disclaimer")
    print("=" * 70 + "\n")

if __name__ == '__main__':
    main()
