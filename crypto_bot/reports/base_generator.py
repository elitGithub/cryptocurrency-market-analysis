"""
Base Report Generator - Abstract base class for report generation.
"""
import os
import logging
from abc import ABC, abstractmethod
from typing import Dict, List, Optional
import pandas as pd


class ReportGenerator(ABC):
    """Abstract base class for all report generators."""

    def __init__(self, output_dir: str):
        """
        Initialize report generator.

        Args:
            output_dir: Directory to save reports
        """
        self.output_dir = output_dir
        self._ensure_output_dir()

    def _ensure_output_dir(self) -> None:
        """Ensure output directory exists."""
        try:
            os.makedirs(self.output_dir, exist_ok=True)
        except OSError as e:
            logging.error(f"Failed to create output directory '{self.output_dir}': {e}")
            raise

    @abstractmethod
    def generate_node(
        self,
        symbol: str,
        signal_data: Dict,
        exchange_results: pd.DataFrame,
        suggestions: List[str],
        df: Optional[pd.DataFrame],
        chart_path: Optional[str],
        date_prefix: str
    ) -> bool:
        """
        Generate a report.

        Args:
            symbol: Trading symbol
            signal_data: Trading signal information
            exchange_results: Exchange comparison data
            suggestions: Analysis suggestions
            df: DataFrame with indicators (can be None)
            chart_path: Path to chart image (can be None)
            date_prefix: Date prefix for filename

        Returns:
            True if successful, False otherwise
        """
        pass