"""
Demo script to showcase report generation with sample data.
This script demonstrates the bot's capabilities without requiring exchange API access.
"""

import pandas as pd
import numpy as np
from datetime import datetime, timedelta
import logging
from report_generator import generate_comprehensive_report
import os

logging.basicConfig(level=logging.INFO, format='%(asctime)s - %(levelname)s - %(message)s')

def generate_sample_data():
    """Generate realistic sample cryptocurrency data for demonstration."""
    
    # Generate 730 days of sample data
    dates = pd.date_range(end=datetime.now(), periods=730, freq='D')
    
    # Start with a base price and add realistic movements
    base_price = 40000
    price_changes = np.random.randn(730) * 1000  # Random daily changes
    trend = np.linspace(0, 10000, 730)  # Upward trend
    prices = base_price + np.cumsum(price_changes) + trend
    
    # Create OHLCV data
    df = pd.DataFrame({
        'open': prices * (1 + np.random.randn(730) * 0.01),
        'high': prices * (1 + abs(np.random.randn(730)) * 0.02),
        'low': prices * (1 - abs(np.random.randn(730)) * 0.02),
        'close': prices,
        'volume': np.random.randint(1000000, 10000000, 730)
    }, index=dates)
    
    # Calculate indicators manually
    df['SMA_short'] = df['close'].rolling(window=50).mean()
    df['SMA_long'] = df['close'].rolling(window=200).mean()
    
    # RSI calculation
    delta = df['close'].diff()
    gain = (delta.where(delta > 0, 0)).rolling(window=14).mean()
    loss = (-delta.where(delta < 0, 0)).rolling(window=14).mean()
    rs = gain / loss
    df['RSI'] = 100 - (100 / (1 + rs))
    
    # Bollinger Bands
    bb_period = 20
    df['BB_middle'] = df['close'].rolling(window=bb_period).mean()
    bb_std = df['close'].rolling(window=bb_period).std()
    df['BB_upper'] = df['BB_middle'] + (bb_std * 2)
    df['BB_lower'] = df['BB_middle'] - (bb_std * 2)
    
    df.dropna(inplace=True)
    
    return df

def create_sample_chart(df, symbol):
    """Create a simple chart from sample data."""
    import matplotlib.pyplot as plt
    
    fig, (ax1, ax2) = plt.subplots(2, 1, figsize=(12, 8), gridspec_kw={'height_ratios': [3, 1]})
    
    # Plot price and moving averages
    ax1.plot(df.index, df['close'], label='Price', linewidth=1.5, color='black')
    ax1.plot(df.index, df['SMA_short'], label='50-Day MA', linewidth=1, color='blue')
    ax1.plot(df.index, df['SMA_long'], label='200-Day MA', linewidth=1, color='orange')
    ax1.fill_between(df.index, df['BB_lower'], df['BB_upper'], alpha=0.2, color='gray', label='Bollinger Bands')
    ax1.set_ylabel('Price (USDT)', fontsize=12)
    ax1.set_title(f'{symbol} Technical Analysis (Demo Data)', fontsize=14, fontweight='bold')
    ax1.legend(loc='upper left')
    ax1.grid(True, alpha=0.3)
    
    # Plot RSI
    ax2.plot(df.index, df['RSI'], color='purple', linewidth=1.5)
    ax2.axhline(y=70, color='r', linestyle='--', alpha=0.5, label='Overbought')
    ax2.axhline(y=30, color='g', linestyle='--', alpha=0.5, label='Oversold')
    ax2.fill_between(df.index, 0, df['RSI'], where=(df['RSI'] >= 70), alpha=0.3, color='red')
    ax2.fill_between(df.index, 0, df['RSI'], where=(df['RSI'] <= 30), alpha=0.3, color='green')
    ax2.set_ylabel('RSI', fontsize=12)
    ax2.set_xlabel('Date', fontsize=12)
    ax2.set_ylim([0, 100])
    ax2.legend(loc='upper left')
    ax2.grid(True, alpha=0.3)
    
    plt.tight_layout()
    chart_path = f'{symbol.replace("/", "-")}_chart.png'
    plt.savefig(chart_path, dpi=150, bbox_inches='tight')
    plt.close()
    
    logging.info(f"Demo chart saved as '{chart_path}'")
    return chart_path

def generate_sample_suggestions(df):
    """Generate sample trading suggestions based on indicators."""
    
    latest = df.iloc[-1]
    previous = df.iloc[-2]
    
    suggestions = []
    
    # Check for crossovers
    if latest['SMA_short'] > latest['SMA_long'] and previous['SMA_short'] <= previous['SMA_long']:
        suggestions.append(
            "BULLISH SIGNAL: A 'Golden Cross' may have occurred recently. "
            "The short-term moving average has crossed above the long-term moving average, "
            "suggesting a potential shift to an uptrend."
        )
    elif latest['SMA_short'] < latest['SMA_long'] and previous['SMA_short'] >= previous['SMA_long']:
        suggestions.append(
            "BEARISH SIGNAL: A 'Death Cross' may have occurred recently. "
            "The short-term moving average has crossed below the long-term moving average, "
            "suggesting a potential shift to a downtrend."
        )
    
    # Current trend
    if latest['SMA_short'] > latest['SMA_long']:
        suggestions.append(
            "CURRENT TREND: The asset appears to be in a long-term uptrend, "
            "as the short-term moving average is currently above the long-term average."
        )
    else:
        suggestions.append(
            "CURRENT TREND: The asset appears to be in a long-term downtrend, "
            "as the short-term moving average is currently below the long-term average."
        )
    
    # RSI analysis
    latest_rsi = latest['RSI']
    if latest_rsi > 70:
        suggestions.append(
            f"MOMENTUM WARNING: The asset is in 'overbought' territory (RSI = {latest_rsi:.2f}). "
            "This suggests the recent upward price movement may be losing momentum and could be due for a pullback."
        )
    elif latest_rsi < 30:
        suggestions.append(
            f"MOMENTUM OPPORTUNITY: The asset is in 'oversold' territory (RSI = {latest_rsi:.2f}). "
            "This suggests the recent downward price movement may be exhausted, potentially presenting a rebound opportunity."
        )
    else:
        suggestions.append(
            f"MOMENTUM: The RSI is currently neutral (RSI = {latest_rsi:.2f}), "
            "not indicating extreme overbought or oversold conditions."
        )
    
    return suggestions

def main():
    print("\n" + "="*70)
    print("  CRYPTOCURRENCY TRADING BOT - DEMO MODE")
    print("="*70)
    print("\nGenerating sample reports to demonstrate bot capabilities...")
    print("Note: This uses simulated data. In production, it connects to real exchanges.")
    print()
    
    # Generate sample data
    logging.info("Generating sample cryptocurrency data...")
    df = generate_sample_data()
    
    # Create sample chart
    symbol = "BTC/USDT"
    chart_path = create_sample_chart(df.tail(365), symbol)
    
    # Generate suggestions
    suggestions = generate_sample_suggestions(df)
    
    # Sample exchange results
    exchange_results = [
        {
            'name': 'binance',
            'total_spot_pairs': 2143,
            'usdt_quoted_pairs': 427,
            'supports_fetchOHLCV': True,
            'rate_limit_ms': 50
        },
        {
            'name': 'kucoin',
            'total_spot_pairs': 1876,
            'usdt_quoted_pairs': 384,
            'supports_fetchOHLCV': True,
            'rate_limit_ms': 10
        },
        {
            'name': 'bybit',
            'total_spot_pairs': 543,
            'usdt_quoted_pairs': 289,
            'supports_fetchOHLCV': True,
            'rate_limit_ms': 20
        },
        {
            'name': 'gate',
            'total_spot_pairs': 2891,
            'usdt_quoted_pairs': 512,
            'supports_fetchOHLCV': True,
            'rate_limit_ms': 50
        }
    ]
    
    # Print sample analysis
    print("\n--- Sample Technical Analysis ---")
    for suggestion in suggestions:
        print(f"• {suggestion}\n")
    
    # Generate reports
    logging.info("\nGenerating professional reports...")
    generate_comprehensive_report(
        symbol=symbol,
        exchange_results=exchange_results,
        suggestions=suggestions,
        df_with_indicators=df,
        chart_path=chart_path
    )
    
    print("\n" + "="*70)
    print("✓ DEMO REPORTS GENERATED SUCCESSFULLY!")
    print("="*70)
    print("\n📁 Your demo reports are ready in the outputs folder:")
    print("  📄 crypto_analysis_report.docx")
    print("  📊 crypto_analysis_presentation.pptx")
    print("  📈 BTC-USDT_chart.png")
    print("\nThese demonstrate what the bot produces with real market data.")
    print("Open them to see professional trading analysis reports!")
    print("="*70 + "\n")

if __name__ == '__main__':
    main()
