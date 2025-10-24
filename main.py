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

def main():
    """Main function to run the entire analysis workflow."""

    # --- 1. Configuration ---
    config = configparser.ConfigParser()
    config.read('config.ini')

    target_exchanges = [e.strip() for e in config['exchanges']['target_exchanges'].split(',')]
    symbol_to_analyze = config['analysis']['symbol']
    timeframe = config['analysis']['timeframe']
    history_days = int(config['analysis']['history_days'])
    short_ma = int(config['analysis']['short_ma'])
    long_ma = int(config['analysis']['long_ma'])

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

    # Display comparative table
    results_df = pd.DataFrame(analysis_results)
    print("\nExchange Curation Report:")
    print(results_df.to_string(index=False))
    print("\n-------------------------------------------\n")

    # --- 3. Deep Dive Analysis on a Selected Exchange/Symbol ---
    target_exchange_id = target_exchanges[0]
    exchange = manager.get_exchange(target_exchange_id)

    if not exchange:
        logging.error(f"Could not proceed with analysis for '{target_exchange_id}'.")
        return

    # --- 4. Data Acquisition ---
    ohlcv_data = fetch_full_ohlcv(exchange, symbol_to_analyze, timeframe, history_days)
    if not ohlcv_data:
        logging.error(f"Failed to fetch historical data for {symbol_to_analyze}. Aborting analysis.")
        return

    df = ohlcv_to_dataframe(ohlcv_data)

    # --- 5. Analysis ---
    df_with_indicators = calculate_indicators(df.copy(), short_ma, long_ma)
    suggestions = analyze_trends_and_generate_suggestions(df_with_indicators)

    # --- 6. Console Reporting ---
    text_report = generate_final_report(suggestions, symbol_to_analyze)
    print(text_report)

    # Generate the visual chart
    display_df = df_with_indicators.tail(365)
    generate_chart(display_df, symbol_to_analyze, short_ma, long_ma)
    
    # --- 7. Generate Comprehensive Reports (PowerPoint + Word) ---
    chart_path = f'{symbol_to_analyze.replace("/", "-")}_chart.png'
    signal_data = generate_comprehensive_reports(
        results_df,
        symbol_to_analyze,
        df_with_indicators,
        suggestions,
        chart_path
    )
    
    print(f"\n{'='*60}")
    print(f"✅ REPORTS GENERATED SUCCESSFULLY!")
    print(f"{'='*60}")
    print(f"PowerPoint: /mnt/user-data/outputs/Crypto_Market_Analysis.pptx")
    print(f"Word Doc:   /mnt/user-data/outputs/Crypto_Market_Analysis.docx")
    print(f"\nTrading Signal: {signal_data['signal']} (Confidence: {signal_data['confidence']})")
    print(f"{'='*60}\n")

if __name__ == '__main__':
    main()
