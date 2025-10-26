"""DOCX Report Generator - Creates Word document reports."""
from __future__ import annotations
import os
import logging
from typing import Dict, List, Optional
import pandas as pd
from .base_generator import ReportGenerator

try:
    from docx import Document
    from docx.shared import Inches, Pt, RGBColor
    from docx.enum.text import WD_ALIGN_PARAGRAPH
    DOCX_AVAILABLE = True
except ImportError:
    DOCX_AVAILABLE = False


class DOCXReportGenerator(ReportGenerator):
    """Generates Word document reports using python-docx library."""

    def generate(
        self,
        symbol: str,
        signal_data: Dict,
        exchange_results: pd.DataFrame,
        suggestions: List[str],
        df: Optional[pd.DataFrame],
        chart_path: Optional[str]
    ) -> bool:
        """Generate a Word document report."""
        if not DOCX_AVAILABLE:
            logging.error("python-docx not installed. Run: pip install python-docx")
            return False

        logging.info("Generating Word document report...")

        docx_filename = "Crypto_Market_Analysis.docx"
        docx_path = os.path.join(self.output_dir, docx_filename)

        try:
            doc = Document()
            
            self._add_title(doc, symbol)
            self._add_executive_summary(doc, signal_data)
            self._add_exchange_comparison(doc, exchange_results)
            self._add_technical_analysis(doc, suggestions)
            
            if chart_path and os.path.exists(chart_path):
                self._add_chart(doc, chart_path)
            
            self._add_educational_content(doc)
            self._add_disclaimer(doc)

            doc.save(docx_path)
            logging.info(f"✓ Word document generated: {docx_filename}")
            return True

        except Exception as e:
            logging.error(f"Error during DOCX generation: {e}")
            return False
    
    def _add_title(self, doc: Document, symbol: str) -> None:
        """Add title and subtitle."""
        title = doc.add_heading('Cryptocurrency Market Analysis Report', 0)
        title.alignment = WD_ALIGN_PARAGRAPH.CENTER

        subtitle = doc.add_paragraph(symbol)
        subtitle.alignment = WD_ALIGN_PARAGRAPH.CENTER
        subtitle.runs[0].font.size = Pt(16)
        subtitle.runs[0].font.color.rgb = RGBColor(127, 140, 141)
    
    def _add_executive_summary(self, doc: Document, signal_data: Dict) -> None:
        """Add executive summary section."""
        doc.add_heading('Executive Summary', 1)

        signal = signal_data.get('signal', 'N/A')
        confidence = signal_data.get('confidence', 'N/A')
        reasoning = signal_data.get('reasoning', 'N/A')
        price = signal_data.get('price', 0.0)
        rsi = signal_data.get('rsi', 50.0)

        signal_colors = {
            'BUY': RGBColor(46, 204, 113), 
            'SELL': RGBColor(231, 76, 60), 
            'HOLD': RGBColor(243, 156, 18)
        }
        signal_color = signal_colors.get(signal, RGBColor(128, 128, 128))

        p = doc.add_paragraph()
        p.add_run('Trading Recommendation: ')
        run = p.add_run(signal)
        run.bold = True
        run.font.size = Pt(16)
        run.font.color.rgb = signal_color

        doc.add_paragraph(f'Confidence Level: {confidence}').runs[0].bold = True
        doc.add_paragraph(reasoning).runs[0].italic = True
        doc.add_paragraph(f'Current Price: ${price:,.2f}')
        doc.add_paragraph(f'RSI: {rsi:.1f}')
    
    def _add_exchange_comparison(self, doc: Document, exchange_results: pd.DataFrame) -> None:
        """Add exchange comparison section."""
        doc.add_heading('Exchange Comparison', 1)
        doc.add_paragraph('Analysis of cryptocurrency exchanges:')

        if not exchange_results.empty:
            table = doc.add_table(rows=1, cols=4)
            table.style = 'Light Grid Accent 1'

            hdr_cells = table.rows[0].cells
            hdr_cells[0].text = 'Exchange'
            hdr_cells[1].text = 'Total Pairs'
            hdr_cells[2].text = 'USDT Pairs'
            hdr_cells[3].text = 'Has OHLCV'

            for _, row in exchange_results.iterrows():
                row_cells = table.add_row().cells
                row_cells[0].text = str(row.get('name', 'N/A'))
                row_cells[1].text = str(row.get('total_spot_pairs', 'N/A'))
                row_cells[2].text = str(row.get('usdt_quoted_pairs', 'N/A'))
                row_cells[3].text = 'Yes' if row.get('supports_fetchOHLCV', False) else 'No'
    
    def _add_technical_analysis(self, doc: Document, suggestions: List[str]) -> None:
        """Add technical analysis section."""
        doc.add_heading('Technical Analysis Details', 1)
        for suggestion in suggestions:
            doc.add_paragraph(suggestion, style='List Bullet')
    
    def _add_chart(self, doc: Document, chart_path: str) -> None:
        """Add chart to document."""
        doc.add_heading('Price Chart', 2)
        doc.add_picture(chart_path, width=Inches(6))
    
    def _add_educational_content(self, doc: Document) -> None:
        """Add educational content section."""
        doc.add_heading('Understanding Technical Indicators', 1)
        
        doc.add_heading('Moving Averages', 2)
        doc.add_paragraph('Smooth price data to identify trends', style='List Bullet')
        doc.add_paragraph('Golden Cross: Short MA > Long MA (bullish)', style='List Bullet')
        doc.add_paragraph('Death Cross: Short MA < Long MA (bearish)', style='List Bullet')

        doc.add_heading('RSI (Relative Strength Index)', 2)
        doc.add_paragraph('Measures momentum (0-100 scale)', style='List Bullet')
        doc.add_paragraph('RSI > 70: Overbought (potential pullback)', style='List Bullet')
        doc.add_paragraph('RSI < 30: Oversold (potential bounce)', style='List Bullet')
    
    def _add_disclaimer(self, doc: Document) -> None:
        """Add disclaimer section."""
        doc.add_heading('Important Disclaimer', 2)
        disclaimer_p = doc.add_paragraph()
        disclaimer_p.add_run('This report is for educational purposes only').bold = True
        doc.add_paragraph('NOT financial advice', style='List Bullet')
        doc.add_paragraph('Cryptocurrency trading has substantial risk', style='List Bullet')
        doc.add_paragraph('Always do your own research (DYOR)', style='List Bullet')
        doc.add_paragraph('Never invest more than you can afford to lose', style='List Bullet')
