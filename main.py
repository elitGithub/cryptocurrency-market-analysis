#
# File: main.py
# Description: The main entry point for the market analysis tool.
#

import configparser
from exchange_manager import ExchangeManager, logging
from market_scanner import analyze_exchange_markets
from data_fetcher import fetch_full_ohlcv, ohlcv_to_dataframe
from analysis_engine import calculate_indicators, analyze_trends_and_generate_suggestions, generate_final_report
from reporting import generate_chart
from advanced_report_generator import generate_comprehensive_reports
import pandas as pd
import os
from datetime import datetime # Import datetime

# Configure basic logging
logging.basicConfig(level=logging.INFO, format='%(asctime)s - %(levelname)s - %(message)s')

def main():
    """Main function to run the entire analysis workflow."""

    # --- 1. Configuration ---
    config = configparser.ConfigParser()
    try:
        if not config.read('config.ini'):
            logging.error("Failed to read config.ini. Make sure the file exists in the script's directory.")
            return
    except configparser.Error as e:
        logging.error(f"Error parsing config.ini: {e}")
        return

    try:
        target_exchanges = [e.strip() for e in config.get('exchanges', 'target_exchanges', fallback='').split(',') if e.strip()]
        symbol_to_analyze = config.get('analysis', 'symbol', fallback='BTC/USDT')
        timeframe = config.get('analysis', 'timeframe', fallback='1d')
        history_days = config.getint('analysis', 'history_days', fallback=730)
        short_ma = config.getint('analysis', 'short_ma', fallback=50)
        long_ma = config.getint('analysis', 'long_ma', fallback=200)
        # Read base output directory from config - Corrected to 'output' (singular)
        base_output_dir = config.get('paths', 'output_dir', fallback='output')
    except (configparser.NoSectionError, configparser.NoOptionError, ValueError) as e:
        logging.error(f"Configuration error in config.ini: {e}")
        return

    if not target_exchanges:
        logging.error("No target_exchanges specified in config.ini under [exchanges].")
        return

    # --- Create Directories and Date Prefix ---
    # Define the subfolder for reports, inside the base_output_dir
    reports_output_dir = os.path.join(base_output_dir, 'reports')
    # Get current date for filename prefix
    date_prefix = datetime.now().strftime("%Y-%m-%d") + "_"

    # Ensure both directories exist
    try:
        # Create base output directory (e.g., 'output')
        os.makedirs(base_output_dir, exist_ok=True)
        # Create reports subfolder (e.g., 'output/reports')
        os.makedirs(reports_output_dir, exist_ok=True)
        logging.info(f"Using base output directory: {os.path.abspath(base_output_dir)}")
        logging.info(f"Using reports directory: {os.path.abspath(reports_output_dir)}")
    except OSError as e:
        logging.error(f"Failed to create output directories: {e}")
        return # Cannot proceed

    # --- 2. Exchange Connection and Market Scanning ---
    manager = ExchangeManager(target_exchanges)
    connected_exchanges = manager.get_all_exchanges()

    if not connected_exchanges:
        logging.error("No exchanges could be initialized. Exiting.")
        return

    logging.info("\n--- Starting Exchange Market Analysis ---")
    analysis_results = []
    for ex_id, ex_instance in connected_exchanges.items():
        analysis_results.append(analyze_exchange_markets(ex_id, ex_instance))

    results_df = pd.DataFrame(analysis_results)
    if results_df.empty:
        logging.warning("Exchange analysis resulted in an empty DataFrame.")
    else:
        print("\nExchange Curation Report:")
        print(results_df.to_string(index=False))
        print("\n-------------------------------------------\n")

    # --- 3. Deep Dive Analysis ---
    target_exchange_id = target_exchanges[0]
    exchange = manager.get_exchange(target_exchange_id)

    if not exchange:
        logging.error(f"Could not get exchange instance for '{target_exchange_id}'. Aborting analysis.")
        return

    # --- 4. Data Acquisition and Initial Processing ---
    # Initialize variables to handle potential failures
    df = None
    df_with_indicators = None
    suggestions = ["Report generation started."] # Default suggestion

    try:
        ohlcv_data = fetch_full_ohlcv(exchange, symbol_to_analyze, timeframe, history_days)
        if not ohlcv_data:
            logging.error(f"Failed to fetch historical data for {symbol_to_analyze} on {target_exchange_id}.")
            suggestions = ["Failed to fetch market data."]
        else:
            df = ohlcv_to_dataframe(ohlcv_data)
            if df is None or df.empty:
                logging.error(f"DataFrame is empty after processing OHLCV data for {symbol_to_analyze}.")
                suggestions = ["Processed market data resulted in empty DataFrame."]
            else:
                # --- 5. Analysis ---
                df_with_indicators = calculate_indicators(df.copy(), short_ma, long_ma)
                if df_with_indicators is None or df_with_indicators.empty:
                    logging.error("Analysis engine returned empty DataFrame after calculating indicators.")
                    suggestions = ["Indicator calculation failed or resulted in empty data."]
                    # Keep df_with_indicators as None/empty
                else:
                    suggestions = analyze_trends_and_generate_suggestions(df_with_indicators)

    except Exception as e:
        logging.exception(f"An error occurred during data acquisition or analysis: {e}")
        suggestions = [f"An error occurred: {e}"]
        # Ensure df_with_indicators is None if analysis failed
        df_with_indicators = None


    # --- 6. Console Reporting ---
    text_report = generate_final_report(suggestions, symbol_to_analyze)
    print(text_report)

    # --- 7. Chart Generation ---
    chart_path = None # Initialize chart_path
    # Generate chart only if indicator calculation was successful
    if df_with_indicators is not None and not df_with_indicators.empty:
        try:
            display_df = df_with_indicators.tail(365)
            # Define chart path in the base output directory (it's temporary)
            chart_filename = f'{date_prefix}{symbol_to_analyze.replace("/", "-")}_chart.png'
            chart_path = os.path.join(base_output_dir, chart_filename)

            # Call generate_chart with the explicit save_path (5 arguments)
            generate_chart(display_df, symbol_to_analyze, short_ma, long_ma, chart_path)
        except Exception as e:
            logging.exception(f"Failed during chart generation: {e}")
            chart_path = None # Ensure chart_path is None if generation failed
    else:
        logging.warning("Skipping chart generation as DataFrame with indicators is empty or calculation failed.")


    # --- 8. Generate Comprehensive Reports ---
    # Always call this function, but pass potentially None/empty data
    # The function itself should handle these cases gracefully
    signal_data = generate_comprehensive_reports(
        results_df,             # Exchange comparison data (DataFrame)
        symbol_to_analyze,      # Symbol like BTC/USDT (str)
        df_with_indicators,     # DataFrame with indicators (can be None/empty)
        suggestions,            # List of text suggestions
        chart_path,             # Path to temporary chart file (can be None)
        reports_output_dir,     # Path to the final reports subfolder (str) - CORRECT
        date_prefix             # Date prefix for filenames (str) - CORRECT
    )

    # --- 9. Final Output ---
    # Check signal_data existence and the 'signal' key for ERROR
    if signal_data and signal_data.get('signal') != 'ERROR':
        print(f"\n{'='*60}")
        print(f"✅ REPORTS GENERATED SUCCESSFULLY (check logs for any warnings).")
        print(f"{'='*60}")
        # Construct expected filenames with prefix
        report_base_name = f"{date_prefix}Crypto_Market_Analysis"
        print(f"Reports saved to: {os.path.abspath(reports_output_dir)}/") # Show absolute path
        print(f"  • PowerPoint: {report_base_name}.pptx")
        print(f"  • Word Doc:   {report_base_name}.docx")
        print(f"  • HTML Report: {report_base_name}.html")
        print(f"\nTrading Signal: {signal_data.get('signal', 'N/A')} (Confidence: {signal_data.get('confidence', 'N/A')})")
        print(f"{'='*60}\n")
    else:
        print(f"\n{'='*60}")
        print(f"❌ REPORT GENERATION FAILED or completed with errors.")
        print(f"{'='*60}")
        reason = "Unknown error during report generation."
        if signal_data and 'reasoning' in signal_data:
            reason = signal_data['reasoning']
        elif not signal_data:
            reason = "Report generation function returned None."
        print(f"Reason: {reason}")
        print(f"{'='*60}\n")


    # Clean up the temporary chart file from the base output directory if it exists
    if chart_path and os.path.exists(chart_path):
        try:
            os.remove(chart_path)
            logging.info(f"Cleaned up temporary chart file: {chart_path}")
        except Exception as e:
            logging.warning(f"Could not clean up chart file '{chart_path}': {e}")


if __name__ == '__main__':
    main()