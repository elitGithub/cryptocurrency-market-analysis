"""Crypto Analyzer - Main orchestrator class for cryptocurrency market analysis."""
import os
import logging
import tempfile
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
        output_base_dir: str = 'output'
    ):
        self.exchange_ids = exchange_ids
        self.symbol = symbol
        self.timeframe = timeframe
        self.history_days = history_days
        self.short_ma = short_ma
        self.long_ma = long_ma
        self.output_base_dir = output_base_dir
        
        # Create date-based output directory
        date_str = datetime.now().strftime("%Y-%m-%d")
        self.output_dir = os.path.join(output_base_dir, date_str)
        
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
    
    def run(self) -> bool:
        """Run the complete analysis workflow."""
        logging.info("="*60)
        logging.info("Starting Cryptocurrency Market Analysis")
        logging.info("="*60)
        
        try:
            if not self._scan_exchanges():
                logging.error("Exchange scanning failed")
                return False
            
            if not self._fetch_and_analyze_data():
                logging.error("Data fetching/analysis failed")
                return False
            
            self._generate_chart()
            
            if not self._generate_reports():
                logging.error("Report generation failed")
                return False
            
            self._cleanup()
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
        
        primary_exchange_id = self.exchange_ids[0]
        primary_exchange = self.exchange_manager.get_exchange(primary_exchange_id)
        
        if not primary_exchange:
            logging.error(f"Could not get exchange '{primary_exchange_id}'")
            return False
        
        data_fetcher = DataFetcher(primary_exchange)
        ohlcv_data = data_fetcher.fetch_ohlcv(
            self.symbol, self.timeframe, self.history_days
        )
        
        if not ohlcv_data:
            logging.error("Failed to fetch OHLCV data")
            self.suggestions = ["Failed to fetch market data"]
            return False
        
        df = data_fetcher.to_dataframe(ohlcv_data)
        
        if df is None or df.empty:
            logging.error("DataFrame is empty after conversion")
            self.suggestions = ["Data conversion failed"]
            return False
        
        logging.info("\n--- Performing Technical Analysis ---")
        self.df_with_indicators = self.technical_analyzer.calculate_indicators(df)
        
        if self.df_with_indicators is None or self.df_with_indicators.empty:
            logging.error("Indicator calculation failed")
            self.suggestions = ["Indicator calculation failed"]
            return False
        
        self.signal_data = self.technical_analyzer.determine_signal(self.df_with_indicators)
        self.suggestions = self.technical_analyzer.generate_suggestions(self.df_with_indicators)
        
        print("\n--- Technical Analysis Report ---")
        for suggestion in self.suggestions:
            print(f"• {suggestion}\n")
        
        return True
    
    def _generate_chart(self) -> None:
        """Generate technical analysis chart in temp location."""
        logging.info("\n--- Generating Chart ---")
        
        if self.df_with_indicators is None or self.df_with_indicators.empty:
            logging.warning("Skipping chart generation - no data available")
            return
        
        display_df = self.df_with_indicators.tail(365)
        
        # Create temp file for chart
        temp_fd, temp_path = tempfile.mkstemp(suffix='.png', prefix='chart_')
        os.close(temp_fd)
        self.chart_path = temp_path
        
        ChartGenerator.generate(
            display_df, self.symbol, 
            self.short_ma, self.long_ma, 
            self.chart_path
        )
    
    def _generate_reports(self) -> bool:
        """Generate all report formats."""
        logging.info("\n--- Generating Reports ---")
        
        docx_gen = DOCXReportGenerator(self.output_dir)
        pptx_gen = PPTXReportGenerator(self.output_dir)
        html_gen = HTMLReportGenerator(self.output_dir)
        
        success_count = 0
        
        if docx_gen.generate(
            self.symbol, self.signal_data, self.exchange_results,
            self.suggestions, self.df_with_indicators, self.chart_path
        ):
            success_count += 1
        
        if pptx_gen.generate(
            self.symbol, self.signal_data, self.exchange_results,
            self.suggestions, self.df_with_indicators, self.chart_path
        ):
            success_count += 1
        
        if html_gen.generate(
            self.symbol, self.signal_data, self.exchange_results,
            self.suggestions, self.df_with_indicators, self.chart_path
        ):
            success_count += 1
        
        return success_count >= 2
    
    def _cleanup(self) -> None:
        """Clean up temporary files."""
        if self.chart_path and os.path.exists(self.chart_path):
            try:
                os.remove(self.chart_path)
                logging.info(f"Cleaned up temporary chart")
            except Exception as e:
                logging.warning(f"Could not remove chart: {e}")
    
    def _display_summary(self) -> None:
        """Display final summary."""
        print("\n" + "="*60)
        print("✅ ANALYSIS COMPLETE")
        print("="*60)
        
        print(f"\nReports saved to: {os.path.abspath(self.output_dir)}/")
        print(f"  • PowerPoint: Crypto_Market_Analysis.pptx")
        print(f"  • Word Doc:   Crypto_Market_Analysis.docx")
        print(f"  • HTML Report: Crypto_Market_Analysis.html")
        
        print(f"\nTrading Signal: {self.signal_data.get('signal', 'N/A')} "
              f"(Confidence: {self.signal_data.get('confidence', 'N/A')})")
        print("="*60 + "\n")
