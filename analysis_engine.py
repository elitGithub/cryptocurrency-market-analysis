#
# File: analysis_engine.py
# Description: Module for performing technical analysis and generating trading signals.
#

import pandas as pd
import pandas_ta as ta
import numpy as np
import logging

def calculate_indicators(df: pd.DataFrame, short_ma: int, long_ma: int) -> pd.DataFrame:
    """
    Calculates a set of technical indicators and appends them to the DataFrame.
    This version is more robust and avoids 'append=True' and 'rename'.

    Args:
        df (pd.DataFrame): The OHLCV DataFrame.
        short_ma (int): The period for the short-term moving average.
        long_ma (int): The period for the long-term moving average.

    Returns:
        pd.DataFrame: The DataFrame with indicator columns appended.
    """
    if df.empty:
        return df

    logging.info("Calculating technical indicators...")

    # --- New, Robust Method ---

    # 1. Calculate SMAs and assign directly
    df['SMA_short'] = df.ta.sma(length=short_ma)
    df['SMA_long'] = df.ta.sma(length=long_ma)

    # 2. Calculate RSI and assign directly
    df['RSI'] = df.ta.rsi(length=14) # Using the common 14-period default

    # 3. Calculate Bollinger Bands
    # We will define the parameters here
    bb_length = 20
    bb_std = 2.0

    # pandas-ta bbands returns a DataFrame with multiple columns
    bbands_df = df.ta.bbands(length=bb_length, std=bb_std)

    # --- FINAL ROBUST FIX ---
    # Don't guess column names. Use the column *position*, which is stable.
    # The bbands() function always returns columns in this order:
    # 0: Lower Band
    # 1: Middle Band
    # 2: Upper Band
    # 3: Bandwidth
    # 4: Percent
    # We access the columns by their integer position using .iloc

    df['BB_lower'] = bbands_df.iloc[:, 0]
    df['BB_middle'] = bbands_df.iloc[:, 1]
    df['BB_upper'] = bbands_df.iloc[:, 2]

    # --- End of New Method ---

    # Drop rows with NaN values resulting from indicator calculation
    # (e.g., the first 200 rows for the SMA_long)
    df.dropna(inplace=True)

    logging.info("Indicators calculated successfully.")
    return df

def analyze_trends_and_generate_suggestions(df: pd.DataFrame) -> list:
    """
    Analyzes the indicator data to generate human-readable suggestions.
    (This function remains unchanged)

    Args:
        df (pd.DataFrame): The DataFrame with indicators.

    Returns:
        list: A list of strings containing textual suggestions.
    """
    if df.empty or len(df) < 2:
        return ["Not enough data for analysis."]

    suggestions = []

    # --- Rule 1: Moving Average Crossover ---
    latest = df.iloc[-1]
    previous = df.iloc[-2]

    # Check for Golden Cross
    if latest['SMA_short'] > latest['SMA_long'] and previous['SMA_short'] <= previous['SMA_long']:
        suggestions.append(
            "BULLISH SIGNAL: A 'Golden Cross' may have occurred recently. "
            "The short-term moving average has crossed above the long-term moving average, "
            "suggesting a potential shift to an uptrend."
        )

    # Check for Death Cross
    elif latest['SMA_short'] < latest['SMA_long'] and previous['SMA_short'] >= previous['SMA_long']:
        suggestions.append(
            "BEARISH SIGNAL: A 'Death Cross' may have occurred recently. "
            "The short-term moving average has crossed below the long-term moving average, "
            "suggesting a potential shift to a downtrend."
        )

    # --- Rule 2: Current Trend Status ---
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

    # --- Rule 3: RSI Overbought/Oversold ---
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

def generate_final_report(suggestions: list, symbol: str) -> str:
    """
    Formats the list of suggestions into a final, readable report.
    (This function remains unchanged)

    Args:
        suggestions (list): List of suggestion strings.
        symbol (str): The trading symbol being analyzed.

    Returns:
        str: A formatted, multi-line string report.
    """
    header = f"--- Automated Technical Analysis Report for {symbol} ---\n"
    if not suggestions:
        return header + "No specific signals were generated based on the current ruleset.\n"

    body = "\n".join([f"- {s}" for s in suggestions])
    footer = "\n--- End of Report ---\n"

    return header + body + footer