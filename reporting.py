#
# File: reporting.py
# Description: Module for generating visual and textual reports.
#

import mplfinance as mpf
import pandas as pd
import logging
import os # Import os

# Definition accepts 5 arguments including save_path
def generate_chart(df: pd.DataFrame, symbol: str, short_ma: int, long_ma: int, save_path: str):
    """
    Generates and saves a candlestick chart with moving averages and volume.

    Args:
        df (pd.DataFrame): The DataFrame with OHLCV and indicator data. Must not be None or empty.
        symbol (str): The trading symbol for the chart title.
        short_ma (int): The short moving average period (used for logging info if needed).
        long_ma (int): The long moving average period (used for logging info if needed).
        save_path (str): The full file path to save the chart to.
    """
    # *** FIXED: Check df validity Robustly ***
    if df is None or df.empty:
        logging.warning("Cannot generate chart: Input DataFrame is empty or None.")
        return

    # Check for essential OHLCV columns used by mplfinance
    ohlcv_cols = ['open', 'high', 'low', 'close', 'volume']
    missing_ohlcv = [col for col in ohlcv_cols if col not in df.columns]
    if missing_ohlcv:
        logging.warning(f"Cannot generate chart: DataFrame missing essential OHLCV columns: {missing_ohlcv}")
        return

    # Check if essential OHLCV columns have *any* valid data for plotting
    if not df[ohlcv_cols].notna().any().all():
        logging.warning("Cannot generate chart: Essential OHLCV data columns contain only NaN values.")
        return

    logging.info(f"Generating chart for {symbol}...")

    # Create additional plot objects, handle potential NaNs by plotting only valid series
    add_plots = []
    indicator_cols = ['SMA_short', 'SMA_long', 'BB_lower', 'BB_upper', 'RSI']

    # Check existence and validity for each indicator before adding
    if 'SMA_short' in df.columns and df['SMA_short'].notna().any():
        add_plots.append(mpf.make_addplot(df['SMA_short'], color='blue', width=0.7))
    else:
        logging.warning("SMA_short data missing or all NaN, skipping plot.")

    if 'SMA_long' in df.columns and df['SMA_long'].notna().any():
        add_plots.append(mpf.make_addplot(df['SMA_long'], color='orange', width=0.7))
    else:
        logging.warning("SMA_long data missing or all NaN, skipping plot.")

    if 'BB_lower' in df.columns and df['BB_lower'].notna().any():
        add_plots.append(mpf.make_addplot(df['BB_lower'], color='gray', linestyle='dashdot', width=0.5))
    else:
        logging.warning("BB_lower data missing or all NaN, skipping plot.")

    if 'BB_upper' in df.columns and df['BB_upper'].notna().any():
        add_plots.append(mpf.make_addplot(df['BB_upper'], color='gray', linestyle='dashdot', width=0.5))
    else:
        logging.warning("BB_upper data missing or all NaN, skipping plot.")

    if 'RSI' in df.columns and df['RSI'].notna().any():
        # Plot RSI in a separate panel
        add_plots.append(mpf.make_addplot(df['RSI'], panel=2, color='purple', ylabel='RSI'))
    else:
        logging.warning("RSI data missing or all NaN, skipping plot.")


    # Define the plot aesthetics
    chart_style = 'yahoo'
    index_name = df.index.name.capitalize() if df.index.name else 'Time'
    chart_title = f'\n{symbol} - {index_name} Chart'

    # Ensure the directory for save_path exists
    save_dir = os.path.dirname(save_path)
    try:
        if save_dir: # Only create if path includes a directory part
            os.makedirs(save_dir, exist_ok=True)
    except OSError as e:
        logging.error(f"Failed to create directory for chart '{save_path}': {e}")
        return

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
            addplot=add_plots if add_plots else None, # Pass None if list is empty
            panel_ratios=(3, 1, 1) if 'RSI' in df.columns and df['RSI'].notna().any() else (4,1), # Adjust ratios if RSI panel isn't used
            figscale=1.5,
            savefig=save_path, # Save to the specified full path
            show_nontrading=False # Avoid issues in non-GUI environments
        )
        logging.info(f"Chart saved successfully as '{save_path}'")
    except Exception as e:
        # Log the full traceback for detailed error info
        logging.exception(f"Failed to generate or save chart for {symbol} to '{save_path}': {e}")