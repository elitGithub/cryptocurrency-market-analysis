"""
HTML Report Generator - Creates standalone HTML reports.
"""
import os
import base64
import logging
from typing import Dict, List, Optional
import pandas as pd
from .base_generator import ReportGenerator


class HTMLReportGenerator(ReportGenerator):
    """Generates standalone HTML reports with embedded content."""
    
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
        """Generate a standalone HTML report."""
        logging.info("Generating HTML report...")
        
        try:
            # Prepare data
            html_filename = f"{date_prefix}Crypto_Market_Analysis.html"
            html_path = os.path.join(self.output_dir, html_filename)
            
            # Generate HTML content
            html_content = self._generate_html(
                symbol, signal_data, exchange_results, 
                suggestions, df, chart_path
            )
            
            # Write to file
            with open(html_path, 'w', encoding='utf-8') as f:
                f.write(html_content)
            
            logging.info(f"✓ HTML report generated: {html_filename}")
            return True
            
        except Exception as e:
            logging.error(f"Error during HTML generation: {e}")
            return False
    
    def _generate_html(
        self,
        symbol: str,
        signal_data: Dict,
        exchange_results: pd.DataFrame,
        suggestions: List[str],
        df: Optional[pd.DataFrame],
        chart_path: Optional[str]
    ) -> str:
        """Generate complete HTML content."""
        
        # Prepare data
        signal = signal_data.get('signal', 'HOLD')
        confidence = signal_data.get('confidence', 'LOW')
        reasoning = signal_data.get('reasoning', 'N/A')
        price = signal_data.get('price', 0.0)
        rsi = signal_data.get('rsi', 50.0)
        
        signal_colors = {'BUY': '#2ECC71', 'SELL': '#E74C3C', 'HOLD': '#F39C12'}
        signal_color = signal_colors.get(signal, '#808080')
        
        # Trend determination
        trend_text = "N/A"
        if df is not None and not df.empty and 'SMA_short' in df.columns and 'SMA_long' in df.columns:
            last = df.iloc[-1]
            if pd.notna(last['SMA_short']) and pd.notna(last['SMA_long']):
                trend_text = "Bullish trend" if last['SMA_short'] > last['SMA_long'] else "Bearish trend"
        
        # Get chart as base64
        chart_base64 = ""
        if chart_path and os.path.exists(chart_path):
            try:
                with open(chart_path, "rb") as img_file:
                    chart_base64 = f"data:image/png;base64,{base64.b64encode(img_file.read()).decode('utf-8')}"
            except Exception as e:
                logging.warning(f"Failed to encode chart: {e}")
        
        # Generate exchange table
        exchange_table = self._generate_exchange_table(exchange_results)
        
        # Generate suggestions list
        suggestions_html = '\n'.join([
            f'<li style="margin-bottom: 15px;">{s}</li>' 
            for s in suggestions
        ])
        
        # Chart section
        chart_section = ""
        if chart_base64:
            chart_section = f"""
            <div class="slide-container">
                <h2>Technical Analysis Chart</h2>
                <div style="text-align: center; margin-top: 30px;">
                    <img src="{chart_base64}" style="max-width: 100%; height: auto;" alt="Technical Analysis Chart">
                </div>
            </div>
            """
        
        # Build complete HTML
        html = f"""<!DOCTYPE html>
<html lang="en">
<head>
    <meta charset="UTF-8">
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    <title>Crypto Market Analysis - {symbol}</title>
    <style>
        :root {{
            --color-primary: #1C2833;
            --color-secondary: #2E4053;
            --color-accent: #E74C3C;
            --color-success: #2ECC71;
            --color-warning: #F39C12;
            --color-bg: #F4F6F6;
            --color-text: #2C3E50;
        }}
        
        body {{
            font-family: Arial, sans-serif;
            color: var(--color-text);
            background-color: var(--color-bg);
            margin: 0;
            padding: 20px 0;
        }}
        
        h1 {{ color: var(--color-primary); font-size: 48px; font-weight: 700; }}
        h2 {{ color: var(--color-secondary); font-size: 32px; font-weight: 700; }}
        h3 {{ color: var(--color-secondary); font-size: 24px; font-weight: 700; }}
        p {{ font-size: 18px; line-height: 1.5; }}
        ul {{ font-size: 18px; line-height: 2; }}
        
        .slide-container {{
            width: 960px;
            margin: 20px auto;
            padding: 40px;
            background-color: #fff;
            box-shadow: 0 4px 12px rgba(0,0,0,0.1);
            border: 1px solid #ddd;
            page-break-after: always;
        }}
        
        table {{
            width: 100%;
            border-collapse: collapse;
            margin-top: 20px;
        }}
        
        th {{
            background-color: #2E4053;
            color: white;
            padding: 12px;
            text-align: left;
            font-size: 18px;
        }}
        
        td {{
            padding: 12px;
            border-bottom: 1px solid #ddd;
            font-size: 16px;
        }}
        
        tr:hover {{ background-color: #f5f5f5; }}
        
        .signal-box {{
            background: {signal_color};
            color: white;
            padding: 30px;
            border-radius: 12px;
            margin: 20px 0;
            text-align: center;
        }}
        
        .signal-text {{
            font-size: 56px;
            font-weight: 700;
            margin: 10px 0;
        }}
        
        @media print {{
            body {{ background-color: #fff; padding: 0; }}
            .slide-container {{ margin: 0; box-shadow: none; border: none; page-break-after: always; }}
        }}
    </style>
</head>
<body>
    <!-- Title Slide -->
    <div class="slide-container" style="background: linear-gradient(135deg, #667eea 0%, #764ba2 100%); color: white; text-align: center; padding-top: 150px;">
        <h1 style="color: white;">Cryptocurrency Market Analysis Report</h1>
        <h2 style="color: #AAB7B8; margin-top: 40px;">{symbol}</h2>
    </div>
    
    <!-- Executive Summary -->
    <div class="slide-container">
        <h2>Executive Summary</h2>
        <div class="signal-box">
            <div class="signal-text">{signal}</div>
            <p style="font-size: 24px;">Confidence: {confidence}</p>
        </div>
        <p><strong>Current Price:</strong> ${price:,.2f}</p>
        <p><strong>RSI:</strong> {rsi:.1f}</p>
        <p><strong>Trend:</strong> {trend_text}</p>
        <p style="margin-top: 20px;"><strong>Analysis:</strong> {reasoning}</p>
    </div>
    
    <!-- Exchange Comparison -->
    <div class="slide-container">
        <h2>Exchange Comparison</h2>
        <p>Analysis of multiple cryptocurrency exchanges:</p>
        {exchange_table}
    </div>
    
    <!-- Technical Analysis -->
    <div class="slide-container">
        <h2>Technical Analysis Details</h2>
        <p>Key findings from indicator analysis:</p>
        <ul style="margin-top: 20px;">
            {suggestions_html}
        </ul>
    </div>
    
    {chart_section}
    
    <!-- Understanding Moving Averages -->
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
    </div>
    
    <!-- Understanding RSI -->
    <div class="slide-container">
        <h2>Understanding RSI</h2>
        <h3>Relative Strength Index (0-100)</h3>
        <ul>
            <li><strong style="color: var(--color-accent);">RSI &gt; 70:</strong> Overbought - potential pullback</li>
            <li><strong style="color: var(--color-warning);">RSI 30-70:</strong> Neutral zone</li>
            <li><strong style="color: var(--color-success);">RSI &lt; 30:</strong> Oversold - potential bounce</li>
            <li>Divergence: When price and RSI move in opposite directions</li>
        </ul>
    </div>
    
    <!-- Disclaimer -->
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
    </div>
</body>
</html>"""
        
        return html
    
    def _generate_exchange_table(self, exchange_results: pd.DataFrame) -> str:
        """Generate HTML table for exchange comparison."""
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
            </tr>
            """)
        
        table = f"""
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
        </table>
        """
        
        return table