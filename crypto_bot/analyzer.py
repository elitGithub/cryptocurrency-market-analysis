"""
Crypto Analyzer - Main orchestrator class for cryptocurrency market analysis.
"""
import os
import logging
from datetime import datetime
from typing import Dict, Optional
import pandas as pd

from .core.exchange_manager import ExchangeManager
from .core.market_scanner import MarketScanner
from .core.data_fetcher import DataFetcher
from .core.technical_analyzer import TechnicalAnalyzer
from .core.chart_generator import ChartGenerator
from .reports.docx_generator import DOCXReportGenerator
from .reports.pptx_generator import PPTXReportGenerator
from .reports.html_generator import HTMLReportGenerator


class CryptoAnalyzer:
    """Main orchestrator for cryptocurrency market analysis and reporting."""
    
    def __init__(
        self,
        exchange_ids: list,
        symbol: str,
        timeframe: str = '1d',
        history_days: int = 730,
        short_ma: int = 50,
        long_ma: int = 200,
        output_dir: str = 'output'
    ):
        """
        Initialize the crypto analyzer.
        
        Args:
            exchange_ids: List of exchange identifiers
            symbol: Trading symbol (e.g., 'BTC/USDT')
            timeframe: Timeframe for analysis (e.g., '1d', '4h')
            history_days: Days of historical data to fetch
            short_ma: Short-term moving average period
            long_ma: Long-term moving average period
            output_dir: Base output directory
        """
        self.exchange_ids = exchange_ids
        self.symbol = symbol
        self.timeframe = timeframe
        self.history_days = history_days
        self.short_ma = short_ma
        self.long_ma = long_ma
        self.output_dir = output_dir
        self.reports_dir = os.path.join(output_dir, 'reports')
        
        # Initialize components
        self.exchange_manager = ExchangeManager(exchange_ids)
        self.technical_analyzer = TechnicalAnalyzer(short_ma, long_ma)
        
        # Data storage
        self.exchange_results: Optional[pd.DataFrame] = None
        self.df_with_indicators: Optional[pd.DataFrame] = None
        self.suggestions: list = []
        self.signal_data: Dict = {}
        self.chart_path: Optional[str] = None
        
        # Ensure directories exist
        self._ensure_directories()
    
    def _ensure_directories(self) -> None:
        """Ensure output directories exist."""
        os.makedirs(self.output_dir, exist_ok=True)
        os.makedirs(self.reports_dir, exist_ok=True)
    
    def run(self) -> bool:
        """
        Run the complete analysis workflow.
        
        Returns:
            True if successful, False otherwise
        """
        logging.info("="*60)
        logging.info("Starting Cryptocurrency Market Analysis")
        logging.info("="*60)
        
        try:
            # Step 1: Scan exchanges
            if not self._scan_exchanges():
                logging.error("Exchange scanning failed")
                return False
            
            # Step 2: Fetch and analyze data
            if not self._fetch_and_analyze_data():
                logging.error("Data fetching/analysis failed")
                return False
            
            # Step 3: Generate chart
            self._generate_chart()
            
            # Step 4: Generate reports
            if not self._generate_reports():
                logging.error("Report generation failed")
                return False
            
            # Step 5: Cleanup
            self._cleanup()
            
            # Step 6: Display summary
            self._display_summary()
            
            return True
            
        except Exception as e:
            logging.exception(f"Unexpected error during analysis: {e}")
            return False
    
    def _scan_exchanges(self) -> bool:
        """Scan all exchanges for market data."""
        logging.info("\n--- Scanning Exchanges ---")
        
        exchanges = self.exchange_manager.get_all_exchanges()
        if not exchanges:
            logging.error("No exchanges initialized")
            return False
        
        self.exchange_results = MarketScanner.scan_all_exchanges(exchanges)
        
        if self.exchange_results.empty:
            logging.warning("Exchange scan returned no results")
            return False
        
        print("\n" + self.exchange_results.to_string(index=False))
        return True
    
    def _fetch_and_analyze_data(self) -> bool:
        """Fetch market data and perform technical analysis."""
        logging.info("\n--- Fetching Market Data ---")
        
        # Get primary exchange
        primary_exchange_id = self.exchange_ids[0]
        primary_exchange = self.exchange_manager.get_exchange(primary_exchange_id)
        
        if not primary_exchange:
            logging.error(f"Could not get exchange '{primary_exchange_id}'")
            return False
        
        # Fetch data
        data_fetcher = DataFetcher(primary_exchange)
        ohlcv_data = data_fetcher.fetch_ohlcv(
            self.symbol, self.timeframe, self.history_days
        )
        
        if not ohlcv_data:
            logging.error("Failed to fetch OHLCV data")
            self.suggestions = ["Failed to fetch market data"]
            return False
        
        # Convert to DataFrame
        df = data_fetcher.to_dataframe(ohlcv_data)
        
        if df is None or df.empty:
            logging.error("DataFrame is empty after conversion")
            self.suggestions = ["Data conversion failed"]
            return False
        
        # Calculate indicators
        logging.info("\n--- Performing Technical Analysis ---")
        self.df_with_indicators = self.technical_analyzer.calculate_indicators(df)
        
        if self.df_with_indicators is None or self.df_with_indicators.empty:
            logging.error("Indicator calculation failed")
            self.suggestions = ["Indicator calculation failed"]
            return False
        
        # Generate signal and suggestions
        self.signal_data = self.technical_analyzer.determine_signal(self.df_with_indicators)
        self.suggestions = self.technical_analyzer.generate_suggestions(self.df_with_indicators)
        
        # Print analysis
        print("\n--- Technical Analysis Report ---")
        for suggestion in self.suggestions:
            print(f"• {suggestion}\n")
        
        return True
    
    def _generate_chart(self) -> None:
        """Generate technical analysis chart."""
        logging.info("\n--- Generating Chart ---")
        
        if self.df_with_indicators is None or self.df_with_indicators.empty:
            logging.warning("Skipping chart generation - no data available")
            return
        
        # Use last 365 days for chart
        display_df = self.df_with_indicators.tail(365)
        
        # Generate chart in base output dir (temporary)
        date_prefix = datetime.now().strftime("%Y-%m-%d_")
        chart_filename = f'{date_prefix}{self.symbol.replace("/", "-")}_chart.png'
        self.chart_path = os.path.join(self.output_dir, chart_filename)
        
        ChartGenerator.generate(
            display_df, self.symbol, 
            self.short_ma, self.long_ma, 
            self.chart_path
        )
    
    def _generate_reports(self) -> bool:
        """Generate all report formats."""
        logging.info("\n--- Generating Reports ---")
        
        date_prefix = datetime.now().strftime("%Y-%m-%d_")
        
        # Initialize generators
        docx_gen = DOCXReportGenerator(self.reports_dir)
        pptx_gen = PPTXReportGenerator(self.reports_dir)
        html_gen = HTMLReportGenerator(self.reports_dir)
        
        success_count = 0
        
        # Generate DOCX
        if docx_gen.generate_node(
            self.symbol, self.signal_data, self.exchange_results,
            self.suggestions, self.df_with_indicators, 
            self.chart_path, date_prefix
        ):
            success_count += 1
        
        # Generate PPTX
        if pptx_gen.generate_node(
            self.symbol, self.signal_data, self.exchange_results,
            self.suggestions, self.df_with_indicators, 
            self.chart_path, date_prefix
        ):
            success_count += 1
        
        # Generate HTML
        if html_gen.generate_node(
            self.symbol, self.signal_data, self.exchange_results,
            self.suggestions, self.df_with_indicators, 
            self.chart_path, date_prefix
        ):
            success_count += 1
        
        return success_count >= 2  # At least 2 out of 3 should succeed
    
    def _cleanup(self) -> None:
        """Clean up temporary files."""
        if self.chart_path and os.path.exists(self.chart_path):
            try:
                os.remove(self.chart_path)
                logging.info(f"Cleaned up temporary chart: {self.chart_path}")
            except Exception as e:
                logging.warning(f"Could not remove chart: {e}")
    
    def _display_summary(self) -> None:
        """Display final summary."""
        print("\n" + "="*60)
        print("✅ ANALYSIS COMPLETE")
        print("="*60)
        
        print(f"\nReports saved to: {os.path.abspath(self.reports_dir)}/")
        
        date_prefix = datetime.now().strftime("%Y-%m-%d_")
        print(f"  • PowerPoint: {date_prefix}Crypto_Market_Analysis.pptx")
        print(f"  • Word Doc:   {date_prefix}Crypto_Market_Analysis.docx")
        print(f"  • HTML Report: {date_prefix}Crypto_Market_Analysis.html")
        
        print(f"\nTrading Signal: {self.signal_data.get('signal', 'N/A')} "
              f"(Confidence: {self.signal_data.get('confidence', 'N/A')})")
        print("="*60 + "\n")