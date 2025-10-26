"""Base Report Generator - Abstract base class for report generation."""
import os
import logging
from abc import ABC, abstractmethod
from typing import Dict, List, Optional
import pandas as pd


class ReportGenerator(ABC):
    """Abstract base class for all report generators."""

    def __init__(self, output_dir: str):
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
    def generate(
        self,
        symbol: str,
        signal_data: Dict,
        exchange_results: pd.DataFrame,
        suggestions: List[str],
        df: Optional[pd.DataFrame],
        chart_path: Optional[str]
    ) -> bool:
        """Generate a report."""
        pass
