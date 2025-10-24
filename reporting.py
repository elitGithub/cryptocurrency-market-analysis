#
# File: reporting.py
# Description: Module for generating visual and textual reports.
#

import mplfinance as mpf
import pandas as pd
import logging

def generate_chart(df: pd.DataFrame, symbol: str, short_ma: int, long_ma: int):
    """
    Generates and displays a candlestick chart with moving averages and volume.

    Args:
        df (pd.DataFrame): The DataFrame with OHLCV and indicator data.
        symbol (str): The trading symbol for the chart title.
        short_ma (int): The short moving average period.
        long_ma (int): The long moving average period.
    """
    if df.empty:
        logging.warning("Cannot generate chart: DataFrame is empty.")
        return

    logging.info(f"Generating chart for {symbol}...")

    # Create additional plot objects for our indicators
    add_plots = [
        mpf.make_addplot(df['SMA_short'], color='blue', width=0.7),
        mpf.make_addplot(df['SMA_long'], color='orange', width=0.7),
        mpf.make_addplot(df['BB_lower'], color='gray', linestyle='dashdot', width=0.5),
        mpf.make_addplot(df['BB_upper'], color='gray', linestyle='dashdot', width=0.5),
        # Plot RSI in a separate panel
        mpf.make_addplot(df['RSI'], panel=2, color='purple', ylabel='RSI')
    ]

    # Define the plot aesthetics
    chart_style = 'yahoo'
    chart_title = f'\n{symbol} - {df.index.name.capitalize()}ly Chart'

    # Generate the plot
    try:
        mpf.plot(
            df,
            type='candle',
            style=chart_style,
            title=chart_title,
            ylabel='Price (USDT)',
            volume=True,
            ylabel_lower='Volume',
            addplot=add_plots,
            panel_ratios=(3, 1, 1), # Main panel, Volume, RSI panel
            figscale=1.5,
            # Save the figure to a file
            savefig=f'{symbol.replace("/", "-")}_chart.png'
        )
        logging.info(f"Chart saved as '{symbol.replace('/', '-')}_chart.png'")
    except Exception as e:
        logging.error(f"Failed to generate chart for {symbol}: {e}")
