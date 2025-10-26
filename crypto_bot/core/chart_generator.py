"""
Chart Generator - Creates technical analysis charts.
"""
import mplfinance as mpf
import pandas as pd
import logging
import os


class ChartGenerator:
    """Generates technical analysis charts using matplotlib."""
    
    @staticmethod
    def generate(
        df: pd.DataFrame, 
        symbol: str, 
        short_ma: int, 
        long_ma: int, 
        save_path: str
    ) -> bool:
        """
        Generate and save a candlestick chart with indicators.
        
        Args:
            df: DataFrame with OHLCV and indicator data
            symbol: Trading symbol for chart title
            short_ma: Short MA period (for legend)
            long_ma: Long MA period (for legend)
            save_path: Full path to save the chart
            
        Returns:
            True if successful, False otherwise
        """
        if df is None or df.empty:
            logging.warning("Cannot generate_node chart: DataFrame is empty")
            return False
        
        # Check for essential OHLCV columns
        ohlcv_cols = ['open', 'high', 'low', 'close', 'volume']
        missing_cols = [col for col in ohlcv_cols if col not in df.columns]
        if missing_cols:
            logging.warning(f"Cannot generate_node chart: Missing columns {missing_cols}")
            return False
        
        # Check if essential data exists
        if not df[ohlcv_cols].notna().any().all():
            logging.warning("Cannot generate_node chart: Essential data columns are all NaN")
            return False
        
        logging.info(f"Generating chart for {symbol}...")
        
        # Create additional plot objects
        add_plots = []
        
        # Add moving averages if available
        if 'SMA_short' in df.columns and df['SMA_short'].notna().any():
            add_plots.append(mpf.make_addplot(df['SMA_short'], color='blue', width=0.7))
        
        if 'SMA_long' in df.columns and df['SMA_long'].notna().any():
            add_plots.append(mpf.make_addplot(df['SMA_long'], color='orange', width=0.7))
        
        # Add Bollinger Bands if available
        if 'BB_lower' in df.columns and df['BB_lower'].notna().any():
            add_plots.append(mpf.make_addplot(
                df['BB_lower'], color='gray', linestyle='dashdot', width=0.5
            ))
        
        if 'BB_upper' in df.columns and df['BB_upper'].notna().any():
            add_plots.append(mpf.make_addplot(
                df['BB_upper'], color='gray', linestyle='dashdot', width=0.5
            ))
        
        # Add RSI in separate panel if available
        has_rsi = 'RSI' in df.columns and df['RSI'].notna().any()
        if has_rsi:
            add_plots.append(mpf.make_addplot(
                df['RSI'], panel=2, color='purple', ylabel='RSI'
            ))
        
        # Prepare chart configuration
        index_name = df.index.name.capitalize() if df.index.name else 'Time'
        chart_title = f'\n{symbol} - {index_name} Chart'
        panel_ratios = (3, 1, 1) if has_rsi else (4, 1)
        
        # Ensure save directory exists
        save_dir = os.path.dirname(save_path)
        if save_dir:
            try:
                os.makedirs(save_dir, exist_ok=True)
            except OSError as e:
                logging.error(f"Failed to create directory '{save_dir}': {e}")
                return False
        
        # Generate the chart
        try:
            mpf.plot(
                df,
                type='candle',
                style='yahoo',
                title=chart_title,
                ylabel='Price (USDT)',
                volume=True,
                ylabel_lower='Volume',
                addplot=add_plots if add_plots else None,
                panel_ratios=panel_ratios,
                figscale=1.5,
                savefig=save_path,
                show_nontrading=False
            )
            
            logging.info(f"✓ Chart saved to '{save_path}'")
            return True
            
        except Exception as e:
            logging.exception(f"Failed to generate_node chart: {e}")
            return False