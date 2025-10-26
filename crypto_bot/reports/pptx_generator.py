"""PPTX Report Generator - Creates PowerPoint presentations."""
from __future__ import annotations
import os
import logging
from typing import Dict, List, Optional
import pandas as pd
from .base_generator import ReportGenerator

try:
    from pptx import Presentation
    from pptx.util import Inches, Pt
    from pptx.enum.text import PP_ALIGN
    from pptx.dml.color import RGBColor
    PPTX_AVAILABLE = True
except ImportError:
    PPTX_AVAILABLE = False


class PPTXReportGenerator(ReportGenerator):
    """Generates PowerPoint presentations using python-pptx library."""

    def generate(
        self,
        symbol: str,
        signal_data: Dict,
        exchange_results: pd.DataFrame,
        suggestions: List[str],
        df: Optional[pd.DataFrame],
        chart_path: Optional[str]
    ) -> bool:
        """Generate a PowerPoint presentation."""
        if not PPTX_AVAILABLE:
            logging.error("python-pptx not installed. Run: pip install python-pptx")
            return False

        logging.info("Generating PowerPoint presentation...")

        pptx_filename = "Crypto_Market_Analysis.pptx"
        pptx_path = os.path.join(self.output_dir, pptx_filename)

        try:
            prs = Presentation()
            prs.slide_width = Inches(10)
            prs.slide_height = Inches(7.5)

            self._add_title_slide(prs, symbol)
            self._add_summary_slide(prs, signal_data)
            self._add_exchange_slide(prs, exchange_results)
            self._add_analysis_slide(prs, suggestions)
            
            if chart_path and os.path.exists(chart_path):
                self._add_chart_slide(prs, chart_path)
            
            self._add_ma_slide(prs)
            self._add_rsi_slide(prs, signal_data.get('rsi', 50.0))
            self._add_disclaimer_slide(prs)

            prs.save(pptx_path)
            logging.info(f"✓ PowerPoint generated: {pptx_filename}")
            return True

        except Exception as e:
            logging.error(f"Error during PPTX generation: {e}")
            return False
    
    def _add_title_slide(self, prs: Presentation, symbol: str) -> None:
        """Add title slide."""
        slide = prs.slides.add_slide(prs.slide_layouts[6])

        title_box = slide.shapes.add_textbox(Inches(1), Inches(2.5), Inches(8), Inches(2))
        tf = title_box.text_frame
        tf.text = "Cryptocurrency Market Analysis"
        p = tf.paragraphs[0]
        p.alignment = PP_ALIGN.CENTER
        p.font.size = Pt(44)
        p.font.bold = True
        p.font.color.rgb = RGBColor(255, 255, 255)

        subtitle_box = slide.shapes.add_textbox(Inches(1), Inches(4.5), Inches(8), Inches(1))
        stf = subtitle_box.text_frame
        stf.text = symbol
        sp = stf.paragraphs[0]
        sp.alignment = PP_ALIGN.CENTER
        sp.font.size = Pt(32)
        sp.font.color.rgb = RGBColor(200, 200, 200)

        slide.background.fill.solid()
        slide.background.fill.fore_color.rgb = RGBColor(102, 126, 234)
    
    def _add_summary_slide(self, prs: Presentation, signal_data: Dict) -> None:
        """Add executive summary slide."""
        slide = prs.slides.add_slide(prs.slide_layouts[5])

        title_box = slide.shapes.add_textbox(Inches(0.5), Inches(0.5), Inches(9), Inches(0.8))
        tf = title_box.text_frame
        tf.text = "Executive Summary"
        p = tf.paragraphs[0]
        p.font.size = Pt(32)
        p.font.bold = True

        signal = signal_data.get('signal', 'HOLD')
        confidence = signal_data.get('confidence', 'LOW')
        reasoning = signal_data.get('reasoning', 'N/A')
        price = signal_data.get('price', 0.0)
        rsi = signal_data.get('rsi', 50.0)

        signal_colors = {
            'BUY': RGBColor(46, 204, 113), 
            'SELL': RGBColor(231, 76, 60), 
            'HOLD': RGBColor(243, 156, 18)
        }
        signal_color = signal_colors.get(signal, RGBColor(128, 128, 128))

        signal_box = slide.shapes.add_textbox(Inches(1.5), Inches(2), Inches(7), Inches(1.5))
        signal_box.fill.solid()
        signal_box.fill.fore_color.rgb = signal_color
        stf = signal_box.text_frame
        stf.text = signal
        sp = stf.paragraphs[0]
        sp.alignment = PP_ALIGN.CENTER
        sp.font.size = Pt(48)
        sp.font.bold = True
        sp.font.color.rgb = RGBColor(255, 255, 255)

        details_box = slide.shapes.add_textbox(Inches(1), Inches(4), Inches(8), Inches(2))
        dtf = details_box.text_frame
        dtf.text = f"Confidence: {confidence}\nPrice: ${price:,.2f} | RSI: {rsi:.1f}\n{reasoning}"
        for p in dtf.paragraphs:
            p.font.size = Pt(16)
    
    def _add_exchange_slide(self, prs: Presentation, exchange_results: pd.DataFrame) -> None:
        """Add exchange comparison slide."""
        slide = prs.slides.add_slide(prs.slide_layouts[5])

        title_box = slide.shapes.add_textbox(Inches(0.5), Inches(0.5), Inches(9), Inches(0.8))
        tf = title_box.text_frame
        tf.text = "Exchange Comparison"
        p = tf.paragraphs[0]
        p.font.size = Pt(32)
        p.font.bold = True

        if not exchange_results.empty:
            rows = len(exchange_results) + 1
            cols = 4
            
            table = slide.shapes.add_table(
                rows, cols, Inches(1), Inches(2), Inches(8), Inches(4)
            ).table

            table.cell(0, 0).text = 'Exchange'
            table.cell(0, 1).text = 'Total Pairs'
            table.cell(0, 2).text = 'USDT Pairs'
            table.cell(0, 3).text = 'Has OHLCV'

            for i in range(cols):
                cell = table.cell(0, i)
                cell.fill.solid()
                cell.fill.fore_color.rgb = RGBColor(46, 64, 83)
                for paragraph in cell.text_frame.paragraphs:
                    for run in paragraph.runs:
                        run.font.color.rgb = RGBColor(255, 255, 255)
                        run.font.bold = True

            for idx, row in enumerate(exchange_results.itertuples(), 1):
                table.cell(idx, 0).text = str(row.name)
                table.cell(idx, 1).text = str(row.total_spot_pairs)
                table.cell(idx, 2).text = str(row.usdt_quoted_pairs)
                table.cell(idx, 3).text = 'Yes' if row.supports_fetchOHLCV else 'No'
    
    def _add_analysis_slide(self, prs: Presentation, suggestions: List[str]) -> None:
        """Add technical analysis slide."""
        slide = prs.slides.add_slide(prs.slide_layouts[5])

        title_box = slide.shapes.add_textbox(Inches(0.5), Inches(0.5), Inches(9), Inches(0.8))
        tf = title_box.text_frame
        tf.text = "Key Takeaways"
        p = tf.paragraphs[0]
        p.font.size = Pt(32)
        p.font.bold = True

        text_box = slide.shapes.add_textbox(Inches(1), Inches(2), Inches(8), Inches(4.5))
        tf = text_box.text_frame
        for suggestion in suggestions[:3]:
            p = tf.add_paragraph()
            p.text = suggestion
            p.level = 0
            p.font.size = Pt(18)
    
    def _add_chart_slide(self, prs: Presentation, chart_path: str) -> None:
        """Add chart slide."""
        slide = prs.slides.add_slide(prs.slide_layouts[5])

        title_box = slide.shapes.add_textbox(Inches(0.5), Inches(0.5), Inches(9), Inches(0.8))
        tf = title_box.text_frame
        tf.text = "Technical Analysis Chart"
        p = tf.paragraphs[0]
        p.font.size = Pt(32)
        p.font.bold = True

        slide.shapes.add_picture(chart_path, Inches(1), Inches(1.5), width=Inches(8))
    
    def _add_ma_slide(self, prs: Presentation) -> None:
        """Add moving averages explanation slide."""
        slide = prs.slides.add_slide(prs.slide_layouts[5])

        title_box = slide.shapes.add_textbox(Inches(0.5), Inches(0.5), Inches(9), Inches(0.8))
        tf = title_box.text_frame
        tf.text = "Understanding Moving Averages"
        p = tf.paragraphs[0]
        p.font.size = Pt(32)
        p.font.bold = True

        text_box = slide.shapes.add_textbox(Inches(1), Inches(2), Inches(8), Inches(4))
        tf = text_box.text_frame

        p = tf.add_paragraph()
        p.text = "Golden Cross: Short MA > Long MA = Bullish"
        p.font.size = Pt(20)
        p.font.color.rgb = RGBColor(46, 204, 113)

        p = tf.add_paragraph()
        p.text = "Death Cross: Short MA < Long MA = Bearish"
        p.font.size = Pt(20)
        p.font.color.rgb = RGBColor(231, 76, 60)
    
    def _add_rsi_slide(self, prs: Presentation, rsi: float) -> None:
        """Add RSI explanation slide."""
        slide = prs.slides.add_slide(prs.slide_layouts[5])

        title_box = slide.shapes.add_textbox(Inches(0.5), Inches(0.5), Inches(9), Inches(0.8))
        tf = title_box.text_frame
        tf.text = "Understanding RSI"
        p = tf.paragraphs[0]
        p.font.size = Pt(32)
        p.font.bold = True

        text_box = slide.shapes.add_textbox(Inches(1), Inches(2), Inches(8), Inches(4))
        tf = text_box.text_frame

        p = tf.add_paragraph()
        p.text = "RSI > 70: Overbought (potential pullback)"
        p.font.size = Pt(20)
        p.font.color.rgb = RGBColor(231, 76, 60)

        p = tf.add_paragraph()
        p.text = "RSI 30-70: Neutral zone"
        p.font.size = Pt(20)

        p = tf.add_paragraph()
        p.text = "RSI < 30: Oversold (potential bounce)"
        p.font.size = Pt(20)
        p.font.color.rgb = RGBColor(46, 204, 113)

        p = tf.add_paragraph()
        p.text = f"\nCurrent RSI: {rsi:.1f}"
        p.font.size = Pt(28)
        p.font.bold = True
        p.alignment = PP_ALIGN.CENTER
    
    def _add_disclaimer_slide(self, prs: Presentation) -> None:
        """Add disclaimer slide."""
        slide = prs.slides.add_slide(prs.slide_layouts[6])
        slide.background.fill.solid()
        slide.background.fill.fore_color.rgb = RGBColor(102, 126, 234)

        title_box = slide.shapes.add_textbox(Inches(1), Inches(2), Inches(8), Inches(4))
        tf = title_box.text_frame

        p = tf.paragraphs[0]
        p.text = "Important Disclaimer"
        p.alignment = PP_ALIGN.CENTER
        p.font.size = Pt(36)
        p.font.bold = True
        p.font.color.rgb = RGBColor(255, 255, 255)

        p = tf.add_paragraph()
        p.text = "\nThis report is for educational purposes only.\nNOT financial advice. Cryptocurrency trading has substantial risk.\nAlways DYOR and never invest more than you can afford to lose."
        p.alignment = PP_ALIGN.CENTER
        p.font.size = Pt(18)
        p.font.color.rgb = RGBColor(255, 255, 255)
