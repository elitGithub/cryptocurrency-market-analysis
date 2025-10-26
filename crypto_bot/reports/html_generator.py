"""HTML Report Generator - Creates standalone HTML reports."""
import os
import base64
import logging
from typing import Dict, List, Optional
import pandas as pd
from .base_generator import ReportGenerator


class HTMLReportGenerator(ReportGenerator):
    """Generates standalone HTML reports with embedded content."""
    
    def generate(
        self,
        symbol: str,
        signal_data: Dict,
        exchange_results: pd.DataFrame,
        suggestions: List[str],
        df: Optional[pd.DataFrame],
        chart_path: Optional[str]
    ) -> bool:
        """Generate a standalone HTML report."""
        logging.info("Generating HTML report...")
        
        try:
            html_filename = "Crypto_Market_Analysis.html"
            html_path = os.path.join(self.output_dir, html_filename)
            
            html_content = self._build_html(
                symbol, signal_data, exchange_results, 
                suggestions, df, chart_path
            )
            
            with open(html_path, 'w', encoding='utf-8') as f:
                f.write(html_content)
            
            logging.info(f"✓ HTML report generated: {html_filename}")
            return True
            
        except Exception as e:
            logging.error(f"Error during HTML generation: {e}")
            return False
    
    def _build_html(
        self,
        symbol: str,
        signal_data: Dict,
        exchange_results: pd.DataFrame,
        suggestions: List[str],
        df: Optional[pd.DataFrame],
        chart_path: Optional[str]
    ) -> str:
        """Build complete HTML content."""
        signal = signal_data.get('signal', 'HOLD')
        confidence = signal_data.get('confidence', 'LOW')
        reasoning = signal_data.get('reasoning', 'N/A')
        price = signal_data.get('price', 0.0)
        rsi = signal_data.get('rsi', 50.0)
        
        signal_colors = {'BUY': '#2ECC71', 'SELL': '#E74C3C', 'HOLD': '#F39C12'}
        signal_color = signal_colors.get(signal, '#808080')
        
        trend_text = self._get_trend_text(df)
        chart_base64 = self._encode_chart(chart_path)
        exchange_table = self._build_exchange_table(exchange_results)
        suggestions_html = self._build_suggestions_list(suggestions)
        chart_section = self._build_chart_section(chart_base64)
        
        return f"""<!DOCTYPE html>
<html lang="en">
<head>
    <meta charset="UTF-8">
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    <title>Crypto Market Analysis - {symbol}</title>
    {self._get_styles()}
</head>
<body>
    {self._build_title_slide(symbol)}
    {self._build_summary_slide(signal, confidence, signal_color, price, rsi, trend_text, reasoning)}
    {self._build_exchange_slide(exchange_table)}
    {self._build_analysis_slide(suggestions_html)}
    {chart_section}
    {self._build_ma_slide()}
    {self._build_rsi_slide(rsi)}
    {self._build_disclaimer_slide()}
</body>
</html>"""
    
    def _get_styles(self) -> str:
        """Return CSS styles."""
        return """<style>
        :root {
            --color-primary: #1C2833;
            --color-secondary: #2E4053;
            --color-accent: #E74C3C;
            --color-success: #2ECC71;
            --color-warning: #F39C12;
            --color-bg: #F4F6F6;
            --color-text: #2C3E50;
        }
        
        body {
            font-family: Arial, sans-serif;
            color: var(--color-text);
            background-color: var(--color-bg);
            margin: 0;
            padding: 20px 0;
        }
        
        h1 { color: var(--color-primary); font-size: 48px; font-weight: 700; }
        h2 { color: var(--color-secondary); font-size: 32px; font-weight: 700; }
        h3 { color: var(--color-secondary); font-size: 24px; font-weight: 700; }
        p { font-size: 18px; line-height: 1.5; }
        ul { font-size: 18px; line-height: 2; }
        
        .slide-container {
            width: 960px;
            margin: 20px auto;
            padding: 40px;
            background-color: #fff;
            box-shadow: 0 4px 12px rgba(0,0,0,0.1);
            border: 1px solid #ddd;
            page-break-after: always;
        }
        
        table {
            width: 100%;
            border-collapse: collapse;
            margin-top: 20px;
        }
        
        th {
            background-color: #2E4053;
            color: white;
            padding: 12px;
            text-align: left;
            font-size: 18px;
        }
        
        td {
            padding: 12px;
            border-bottom: 1px solid #ddd;
            font-size: 16px;
        }
        
        tr:hover { background-color: #f5f5f5; }
        
        .signal-box {
            color: white;
            padding: 30px;
            border-radius: 12px;
            margin: 20px 0;
            text-align: center;
        }
        
        .signal-text {
            font-size: 56px;
            font-weight: 700;
            margin: 10px 0;
        }
        
        @media print {
            body { background-color: #fff; padding: 0; }
            .slide-container { margin: 0; box-shadow: none; border: none; }
        }
    </style>"""
    
    def _build_title_slide(self, symbol: str) -> str:
        """Build title slide."""
        return f"""
    <div class="slide-container" style="background: linear-gradient(135deg, #667eea 0%, #764ba2 100%); color: white; text-align: center; padding-top: 150px;">
        <h1 style="color: white;">Cryptocurrency Market Analysis Report</h1>
        <h2 style="color: #AAB7B8; margin-top: 40px;">{symbol}</h2>
    </div>"""
    
    def _build_summary_slide(self, signal: str, confidence: str, signal_color: str, 
                            price: float, rsi: float, trend_text: str, reasoning: str) -> str:
        """Build executive summary slide."""
        return f"""
    <div class="slide-container">
        <h2>Executive Summary</h2>
        <div class="signal-box" style="background: {signal_color};">
            <div class="signal-text">{signal}</div>
            <p style="font-size: 24px;">Confidence: {confidence}</p>
        </div>
        <p><strong>Current Price:</strong> ${price:,.2f}</p>
        <p><strong>RSI:</strong> {rsi:.1f}</p>
        <p><strong>Trend:</strong> {trend_text}</p>
        <p style="margin-top: 20px;"><strong>Analysis:</strong> {reasoning}</p>
    </div>"""
    
    def _build_exchange_slide(self, exchange_table: str) -> str:
        """Build exchange comparison slide."""
        return f"""
    <div class="slide-container">
        <h2>Exchange Comparison</h2>
        <p>Analysis of multiple cryptocurrency exchanges:</p>
        {exchange_table}
    </div>"""
    
    def _build_analysis_slide(self, suggestions_html: str) -> str:
        """Build technical analysis slide."""
        return f"""
    <div class="slide-container">
        <h2>Technical Analysis Details</h2>
        <p>Key findings from indicator analysis:</p>
        <ul style="margin-top: 20px;">
            {suggestions_html}
        </ul>
    </div>"""
    
    def _build_chart_section(self, chart_base64: str) -> str:
        """Build chart section if chart exists."""
        if not chart_base64:
            return ""
        
        return f"""
    <div class="slide-container">
        <h2>Technical Analysis Chart</h2>
        <div style="text-align: center; margin-top: 30px;">
            <img src="{chart_base64}" style="max-width: 100%; height: auto;" alt="Technical Analysis Chart">
        </div>
    </div>"""
    
    def _build_ma_slide(self) -> str:
        """Build moving averages explanation slide."""
        return """
    <div class="slide-container">
        <h2>Understanding Moving Averages</h2>
        <h3>What Are Moving Averages?</h3>
        <p>Moving averages smooth out price data to identify trends over time.</p>
        <ul>
            <li><strong style="color: var(--color-success);">Golden Cross:</strong> Short-term MA crosses above long-term MA = Bullish signal</li>
            <li><strong style="color: var(--color-accent);">Death Cross:</strong> Short-term MA crosses below long-term MA = Bearish signal</li>
            <li>Price above MA = Uptrend</li>
            <li>Price below MA = Downtrend</li>
        </ul>
    </div>"""
    
    def _build_rsi_slide(self, rsi: float) -> str:
        """Build RSI explanation slide."""
        return f"""
    <div class="slide-container">
        <h2>Understanding RSI</h2>
        <h3>Relative Strength Index (0-100)</h3>
        <ul>
            <li><strong style="color: var(--color-accent);">RSI &gt; 70:</strong> Overbought - potential pullback</li>
            <li><strong style="color: var(--color-warning);">RSI 30-70:</strong> Neutral zone</li>
            <li><strong style="color: var(--color-success);">RSI &lt; 30:</strong> Oversold - potential bounce</li>
            <li>Divergence: When price and RSI move in opposite directions</li>
        </ul>
    </div>"""
    
    def _build_disclaimer_slide(self) -> str:
        """Build disclaimer slide."""
        return """
    <div class="slide-container" style="background: linear-gradient(135deg, #667eea 0%, #764ba2 100%); color: white;">
        <h2 style="color: white;">Important Disclaimer</h2>
        <p style="font-size: 20px; font-weight: bold;">This report is for educational purposes only. NOT financial advice.</p>
        <ul style="font-size: 18px;">
            <li>Cryptocurrency trading has substantial risk of loss</li>
            <li>Always do your own research (DYOR)</li>
            <li>Never invest more than you can afford to lose</li>
            <li>Past performance does not guarantee future results</li>
            <li>Consult a qualified financial advisor</li>
        </ul>
    </div>"""
    
    def _get_trend_text(self, df: Optional[pd.DataFrame]) -> str:
        """Determine trend text from dataframe."""
        if df is None or df.empty:
            return "N/A"
        
        if 'SMA_short' not in df.columns or 'SMA_long' not in df.columns:
            return "N/A"
        
        last = df.iloc[-1]
        if pd.notna(last['SMA_short']) and pd.notna(last['SMA_long']):
            return "Bullish trend" if last['SMA_short'] > last['SMA_long'] else "Bearish trend"
        
        return "N/A"
    
    def _encode_chart(self, chart_path: Optional[str]) -> str:
        """Encode chart image to base64."""
        if not chart_path or not os.path.exists(chart_path):
            return ""
        
        try:
            with open(chart_path, "rb") as img_file:
                return f"data:image/png;base64,{base64.b64encode(img_file.read()).decode('utf-8')}"
        except Exception as e:
            logging.warning(f"Failed to encode chart: {e}")
            return ""
    
    def _build_exchange_table(self, exchange_results: pd.DataFrame) -> str:
        """Build HTML table for exchange comparison."""
        if exchange_results is None or exchange_results.empty:
            return "<p>No exchange data available.</p>"
        
        rows = []
        for _, row in exchange_results.iterrows():
            name = str(row.get('name', 'N/A'))
            total_pairs = str(row.get('total_spot_pairs', 'N/A'))
            usdt_pairs = str(row.get('usdt_quoted_pairs', 'N/A'))
            has_ohlcv = row.get('supports_fetchOHLCV', False)
            ohlcv_text = '✓' if has_ohlcv else '✗'
            ohlcv_color = 'var(--color-success)' if has_ohlcv else 'var(--color-accent)'
            
            rows.append(f"""
            <tr>
                <td>{name}</td>
                <td style="text-align: center;">{total_pairs}</td>
                <td style="text-align: center;">{usdt_pairs}</td>
                <td style="text-align: center; color: {ohlcv_color}; font-weight: bold;">{ohlcv_text}</td>
            </tr>""")
        
        return f"""
        <table>
            <thead>
                <tr>
                    <th>Exchange</th>
                    <th style="text-align: center;">Total Pairs</th>
                    <th style="text-align: center;">USDT Pairs</th>
                    <th style="text-align: center;">Has OHLCV</th>
                </tr>
            </thead>
            <tbody>
                {''.join(rows)}
            </tbody>
        </table>"""
    
    def _build_suggestions_list(self, suggestions: List[str]) -> str:
        """Build suggestions HTML list."""
        return '\n'.join([
            f'<li style="margin-bottom: 15px;">{s}</li>' 
            for s in suggestions
        ])
