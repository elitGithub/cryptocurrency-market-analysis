"""Technical Analyzer - Calculates indicators and generates trading signals."""
import pandas as pd
import pandas_ta as ta
import logging
from typing import Dict, List, Optional


class TechnicalAnalyzer:
    """Performs technical analysis on market data."""

    def __init__(self, short_ma: int = 50, long_ma: int = 200):
        self.short_ma = short_ma
        self.long_ma = long_ma

    def calculate_indicators(self, df: pd.DataFrame) -> Optional[pd.DataFrame]:
        """Calculate technical indicators on OHLCV data."""
        if df is None or df.empty:
            logging.warning("Empty DataFrame provided for indicator calculation")
            return None

        logging.info("Calculating technical indicators...")
        df = df.copy()

        df['SMA_short'] = ta.sma(df['close'], length=self.short_ma)
        df['SMA_long'] = ta.sma(df['close'], length=self.long_ma)
        df['RSI'] = ta.rsi(df['close'], length=14)

        bbands = ta.bbands(df['close'], length=20, std=2.0)
        if bbands is not None and not bbands.empty:
            df['BB_lower'] = bbands.iloc[:, 0]
            df['BB_middle'] = bbands.iloc[:, 1]
            df['BB_upper'] = bbands.iloc[:, 2]

        df.dropna(inplace=True)
        logging.info(f"✓ Calculated indicators ({len(df)} valid rows)")
        return df

    def determine_signal(self, df: pd.DataFrame) -> Dict:
        """Analyze indicators and generate a trading signal."""
        if df is None or df.empty or len(df) < 2:
            return {
                'signal': 'HOLD',
                'confidence': 'LOW',
                'reasoning': 'Insufficient data for analysis',
                'score': 0,
                'rsi': 50.0,
                'price': 0.0
            }

        latest = df.iloc[-1]
        previous = df.iloc[-2]
        score = 0
        reasons = []

        has_ma = all(col in df.columns for col in ['SMA_short', 'SMA_long'])
        has_rsi = 'RSI' in df.columns
        has_bb = all(col in df.columns for col in ['BB_lower', 'BB_upper', 'close'])

        # Moving Average Analysis
        if has_ma and pd.notna(latest['SMA_short']) and pd.notna(latest['SMA_long']):
            if latest['SMA_short'] > latest['SMA_long']:
                if previous['SMA_short'] <= previous['SMA_long']:
                    score += 2
                    reasons.append("Golden Cross detected (strong buy signal)")
                else:
                    score += 1
                    reasons.append("Uptrend confirmed by moving averages")
            else:
                if previous['SMA_short'] >= previous['SMA_long']:
                    score -= 2
                    reasons.append("Death Cross detected (strong sell signal)")
                else:
                    score -= 1
                    reasons.append("Downtrend confirmed by moving averages")

        # RSI Analysis
        rsi = 50.0
        if has_rsi and pd.notna(latest['RSI']):
            rsi = latest['RSI']
            if rsi < 30:
                score += 1
                reasons.append(f"Oversold conditions (RSI: {rsi:.1f})")
            elif rsi > 70:
                score -= 1
                reasons.append(f"Overbought conditions (RSI: {rsi:.1f})")
            else:
                reasons.append(f"Neutral momentum (RSI: {rsi:.1f})")

        # Bollinger Bands Analysis
        price = latest.get('close', 0.0)
        if has_bb and price and pd.notna(price):
            if pd.notna(latest['BB_lower']) and price < latest['BB_lower']:
                score += 1
                reasons.append("Price below lower Bollinger Band (potential reversal)")
            elif pd.notna(latest['BB_upper']) and price > latest['BB_upper']:
                score -= 1
                reasons.append("Price above upper Bollinger Band (potential correction)")

        # Determine signal
        if score >= 2:
            signal = 'BUY'
            confidence = 'HIGH' if score >= 3 else 'MEDIUM'
        elif score <= -2:
            signal = 'SELL'
            confidence = 'HIGH' if score <= -3 else 'MEDIUM'
        else:
            signal = 'HOLD'
            confidence = 'MEDIUM' if abs(score) == 1 else 'LOW'

        return {
            'signal': signal,
            'confidence': confidence,
            'reasoning': ' | '.join(reasons) if reasons else "No clear signals",
            'score': score,
            'rsi': float(rsi) if pd.notna(rsi) else 50.0,
            'price': float(price) if pd.notna(price) else 0.0
        }

    def generate_suggestions(self, df: pd.DataFrame) -> List[str]:
        """Generate human-readable analysis suggestions."""
        if df is None or df.empty or len(df) < 2:
            return ["Not enough data for analysis"]
        
        suggestions = []
        latest = df.iloc[-1]
        previous = df.iloc[-2]
        
        # Moving Average Crossover
        if 'SMA_short' in df.columns and 'SMA_long' in df.columns:
            if latest['SMA_short'] > latest['SMA_long']:
                if previous['SMA_short'] <= previous['SMA_long']:
                    suggestions.append(
                        "BULLISH SIGNAL: A 'Golden Cross' occurred recently. "
                        "The short-term moving average crossed above the long-term average, "
                        "suggesting a potential shift to an uptrend."
                    )
                else:
                    suggestions.append(
                        "CURRENT TREND: The asset is in a long-term uptrend, "
                        "as the short-term moving average is above the long-term average."
                    )
            else:
                if previous['SMA_short'] >= previous['SMA_long']:
                    suggestions.append(
                        "BEARISH SIGNAL: A 'Death Cross' occurred recently. "
                        "The short-term moving average crossed below the long-term average, "
                        "suggesting a potential shift to a downtrend."
                    )
                else:
                    suggestions.append(
                        "CURRENT TREND: The asset is in a long-term downtrend, "
                        "as the short-term moving average is below the long-term average."
                    )
        
        # RSI Analysis
        if 'RSI' in df.columns and pd.notna(latest['RSI']):
            rsi = latest['RSI']
            if rsi > 70:
                suggestions.append(
                    f"MOMENTUM WARNING: The asset is overbought (RSI = {rsi:.2f}). "
                    "Recent upward movement may be losing momentum and could pull back."
                )
            elif rsi < 30:
                suggestions.append(
                    f"MOMENTUM OPPORTUNITY: The asset is oversold (RSI = {rsi:.2f}). "
                    "Recent downward movement may be exhausted, presenting a rebound opportunity."
                )
            else:
                suggestions.append(
                    f"MOMENTUM: The RSI is neutral (RSI = {rsi:.2f}), "
                    "not indicating extreme conditions."
                )
        
        return suggestions
