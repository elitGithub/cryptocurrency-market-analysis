#
# File: advanced_report_generator.py
# Description: Generates comprehensive PowerPoint and Word reports with trading signals
#

import pandas as pd
import logging
import subprocess
import json
import os

def determine_trading_signal(df: pd.DataFrame) -> dict:
    """
    Analyzes indicators and generates a clear BUY/SELL/HOLD signal.
    
    Args:
        df: DataFrame with technical indicators
        
    Returns:
        dict with signal, confidence, and reasoning
    """
    if df.empty or len(df) < 2:
        return {
            'signal': 'HOLD',
            'confidence': 'LOW',
            'reasoning': 'Insufficient data for analysis',
            'score': 0
        }

    latest = df.iloc[-1]
    previous = df.iloc[-2]

    # Scoring system: -2 to +2 for each indicator
    score = 0
    reasons = []

    # 1. Moving Average Crossover (strongest signal)
    if latest['SMA_short'] > latest['SMA_long']:
        if previous['SMA_short'] <= previous['SMA_long']:
            score += 2
            reasons.append("Golden Cross detected (strong buy signal)")
        else:
            score += 1
            reasons.append("Uptrend confirmed by moving averages")
    else:
        if previous['SMA_short'] >= previous['SMA_long']:
            score -= 2
            reasons.append("Death Cross detected (strong sell signal)")
        else:
            score -= 1
            reasons.append("Downtrend confirmed by moving averages")

    # 2. RSI Analysis
    rsi = latest['RSI']
    if rsi < 30:
        score += 1
        reasons.append(f"Oversold conditions (RSI: {rsi:.1f})")
    elif rsi > 70:
        score -= 1
        reasons.append(f"Overbought conditions (RSI: {rsi:.1f})")
    else:
        reasons.append(f"Neutral momentum (RSI: {rsi:.1f})")

    # 3. Price vs Bollinger Bands
    price = latest['close']
    if price < latest['BB_lower']:
        score += 1
        reasons.append("Price below lower Bollinger Band (potential reversal)")
    elif price > latest['BB_upper']:
        score -= 1
        reasons.append("Price above upper Bollinger Band (potential correction)")

    # Determine signal based on score
    if score >= 2:
        signal = 'BUY'
        confidence = 'HIGH' if score >= 3 else 'MEDIUM'
    elif score <= -2:
        signal = 'SELL'
        confidence = 'HIGH' if score <= -3 else 'MEDIUM'
    else:
        signal = 'HOLD'
        confidence = 'MEDIUM' if abs(score) == 1 else 'LOW'

    return {
        'signal': signal,
        'confidence': confidence,
        'reasoning': ' | '.join(reasons),
        'score': score,
        'rsi': rsi,
        'price': price
    }

def generate_powerpoint_report(
        signal_data: dict,
        exchange_results: pd.DataFrame,
        symbol: str,
        df: pd.DataFrame,
        chart_path: str
):
    """
    Generates a PowerPoint presentation with trading analysis.
    """
    logging.info("Generating PowerPoint report...")

    # Install html2pptx if needed
    subprocess.run(
        "npm list -g @ant/html2pptx || npm install -g /mnt/skills/public/pptx/html2pptx.tgz",
        shell=True,
        capture_output=True
    )

    # Create shared CSS
    css_content = """
:root {
    --color-primary: #1C2833;
    --color-secondary: #2E4053;
    --color-accent: #E74C3C;
    --color-success: #2ECC71;
    --color-warning: #F39C12;
    --color-bg: #F4F6F6;
    --color-text: #2C3E50;
    --font-main: Arial, sans-serif;
}

body {
    font-family: var(--font-main);
    color: var(--color-text);
}

h1 { color: var(--color-primary); font-size: 48px; font-weight: bold; }
h2 { color: var(--color-secondary); font-size: 32px; font-weight: bold; }
h3 { color: var(--color-secondary); font-size: 24px; font-weight: bold; }
p { font-size: 18px; line-height: 1.5; }
.big-signal { font-size: 72px; font-weight: bold; }
.confidence { font-size: 24px; font-style: italic; }
"""

    with open('shared.css', 'w') as f:
        f.write(css_content)

    # Slide 1: Title & Executive Summary
    signal_color = {
        'BUY': '#2ECC71',
        'SELL': '#E74C3C',
        'HOLD': '#F39C12'
    }[signal_data['signal']]

    slide1 = f"""<!DOCTYPE html>
<html>
<head>
<style>{css_content}</style>
</head>
<body style="width: 960px; height: 540px; margin: 0; padding: 0; background: linear-gradient(135deg, #1C2833 0%, #2E4053 100%);">
    <div class="col" style="height: 100%; justify-content: center; align-items: center; padding: 60px;">
        <h1 style="color: white; text-align: center; margin-bottom: 40px;">Crypto Market Analysis Report</h1>
        <h2 style="color: #AAB7B8; text-align: center; margin-bottom: 60px;">{symbol}</h2>
        <div style="text-align: center;">
            <p class="big-signal" style="color: {signal_color};">{signal_data['signal']}</p>
            <p class="confidence" style="color: #BDC3C7;">Confidence: {signal_data['confidence']}</p>
        </div>
    </div>
</body>
</html>"""

    with open('slide1.html', 'w') as f:
        f.write(slide1)

    # Slide 2: Detailed Recommendation
    slide2 = f"""<!DOCTYPE html>
<html>
<head>
<style>{css_content}</style>
</head>
<body style="width: 960px; height: 540px; margin: 0; padding: 0; background: var(--color-bg);">
    <div class="col" style="height: 100%; padding: 40px;">
        <h2 style="margin-bottom: 30px;">Recommendation Details</h2>
        <div style="background: white; padding: 30px; border-radius: 10px; box-shadow: 0 4px 6px rgba(0,0,0,0.1);">
            <div class="row" style="align-items: center; margin-bottom: 20px;">
                <h1 style="color: {signal_color}; margin: 0;">{signal_data['signal']}</h1>
                <p style="margin-left: 20px; font-size: 20px;">({signal_data['confidence']} Confidence)</p>
            </div>
            <h3 style="margin-bottom: 15px;">Analysis:</h3>
            <p style="font-size: 18px; line-height: 1.6;">{signal_data['reasoning']}</p>
            <div class="row" style="margin-top: 30px; gap: 40px;">
                <div>
                    <p style="font-size: 16px; color: #7F8C8D; margin-bottom: 5px;">Current Price</p>
                    <p style="font-size: 24px; font-weight: bold;">${signal_data['price']:,.2f}</p>
                </div>
                <div>
                    <p style="font-size: 16px; color: #7F8C8D; margin-bottom: 5px;">RSI</p>
                    <p style="font-size: 24px; font-weight: bold;">{signal_data['rsi']:.1f}</p>
                </div>
            </div>
        </div>
    </div>
</body>
</html>"""

    with open('slide2.html', 'w') as f:
        f.write(slide2)

    # Slide 3: Exchange Comparison
    slide3 = f"""<!DOCTYPE html>
<html>
<head>
<style>{css_content}</style>
</head>
<body style="width: 960px; height: 540px; margin: 0; padding: 0; background: var(--color-bg);">
    <div class="col" style="height: 100%; padding: 40px;">
        <h2 style="margin-bottom: 30px;">Exchange Comparison</h2>
        <div class="placeholder" style="flex: 1;"></div>
    </div>
</body>
</html>"""

    with open('slide3.html', 'w') as f:
        f.write(slide3)

    # Slide 4: Technical Chart
    slide4 = f"""<!DOCTYPE html>
<html>
<head>
<style>{css_content}</style>
</head>
<body style="width: 960px; height: 540px; margin: 0; padding: 0; background: var(--color-bg);">
    <div class="col" style="height: 100%; padding: 40px;">
        <h2 style="margin-bottom: 20px;">Technical Analysis Chart</h2>
        <div style="flex: 1; display: flex; justify-content: center; align-items: center;">
            <img src="{os.path.basename(chart_path)}" style="max-width: 100%; max-height: 100%; object-fit: contain;" />
        </div>
    </div>
</body>
</html>"""

    with open('slide4.html', 'w') as f:
        f.write(slide4)

    # Slide 5: What This Means
    slide5 = f"""<!DOCTYPE html>
<html>
<head>
<style>{css_content}</style>
</head>
<body style="width: 960px; height: 540px; margin: 0; padding: 0; background: var(--color-bg);">
    <div class="col" style="height: 100%; padding: 40px;">
        <h2 style="margin-bottom: 30px;">What This Means For You</h2>
        <div style="background: white; padding: 30px; border-radius: 10px;">
            <h3 style="margin-bottom: 15px;">Key Takeaways:</h3>
            <ul style="font-size: 18px; line-height: 2;">
                <li><strong>Signal:</strong> {signal_data['signal']} recommendation based on multiple indicators</li>
                <li><strong>Moving Averages:</strong> {"Bullish trend" if df.iloc[-1]['SMA_short'] > df.iloc[-1]['SMA_long'] else "Bearish trend"} detected</li>
                <li><strong>Momentum:</strong> RSI at {signal_data['rsi']:.1f} ({"Oversold" if signal_data['rsi'] < 30 else "Overbought" if signal_data['rsi'] > 70 else "Neutral"})</li>
                <li><strong>Risk Level:</strong> {"Lower risk entry point" if signal_data['signal'] == 'BUY' else "Higher risk - consider taking profits" if signal_data['signal'] == 'SELL' else "Wait for clearer signals"}</li>
            </ul>
        </div>
    </div>
</body>
</html>"""

    with open('slide5.html', 'w') as f:
        f.write(slide5)

    # Slide 6: Educational - Moving Averages
    slide6 = """<!DOCTYPE html>
<html>
<head>
<style>{css_content}</style>
</head>
<body style="width: 960px; height: 540px; margin: 0; padding: 0; background: var(--color-bg);">
    <div class="col" style="height: 100%; padding: 40px;">
        <h2 style="margin-bottom: 30px;">Learn: Moving Averages</h2>
        <div style="background: white; padding: 30px; border-radius: 10px;">
            <h3 style="margin-bottom: 15px;">What Are Moving Averages?</h3>
            <p style="margin-bottom: 20px;">Moving averages smooth out price data to identify trends by averaging prices over a specific period.</p>
            <h3 style="margin-bottom: 15px;">Key Signals:</h3>
            <ul style="font-size: 18px; line-height: 2;">
                <li><strong>Golden Cross:</strong> Short-term MA crosses above long-term MA (bullish)</li>
                <li><strong>Death Cross:</strong> Short-term MA crosses below long-term MA (bearish)</li>
                <li><strong>Price Above MA:</strong> Generally indicates uptrend</li>
                <li><strong>Price Below MA:</strong> Generally indicates downtrend</li>
            </ul>
        </div>
    </div>
</body>
</html>""".replace('{css_content}', css_content)

    with open('slide6.html', 'w') as f:
        f.write(slide6)

    # Slide 7: Educational - RSI
    slide7 = """<!DOCTYPE html>
<html>
<head>
<style>{css_content}</style>
</head>
<body style="width: 960px; height: 540px; margin: 0; padding: 0; background: var(--color-bg);">
    <div class="col" style="height: 100%; padding: 40px;">
        <h2 style="margin-bottom: 30px;">Learn: RSI (Relative Strength Index)</h2>
        <div style="background: white; padding: 30px; border-radius: 10px;">
            <h3 style="margin-bottom: 15px;">What Is RSI?</h3>
            <p style="margin-bottom: 20px;">RSI measures the speed and magnitude of price changes, ranging from 0 to 100.</p>
            <h3 style="margin-bottom: 15px;">How to Read RSI:</h3>
            <ul style="font-size: 18px; line-height: 2;">
                <li><strong>Above 70:</strong> Overbought - price may be due for a pullback</li>
                <li><strong>Below 30:</strong> Oversold - price may be due for a bounce</li>
                <li><strong>30-70:</strong> Neutral zone - no extreme conditions</li>
                <li><strong>Divergence:</strong> When price and RSI move in opposite directions</li>
            </ul>
        </div>
    </div>
</body>
</html>""".replace('{css_content}', css_content)

    with open('slide7.html', 'w') as f:
        f.write(slide7)

    # Create JavaScript to generate PPTX
    js_code = f"""
const pptxgen = require("pptxgenjs");
const {{ html2pptx }} = require("@ant/html2pptx");
const fs = require('fs');

(async () => {{
    const pptx = new pptxgen();
    pptx.layout = "LAYOUT_16x9";
    pptx.author = "Crypto Trading Bot";
    pptx.title = "Market Analysis Report - {symbol}";
    
    // Slide 1: Title
    await html2pptx("slide1.html", pptx);
    
    // Slide 2: Recommendation
    await html2pptx("slide2.html", pptx);
    
    // Slide 3: Exchange Comparison
    const {{ slide: slide3, placeholders: p3 }} = await html2pptx("slide3.html", pptx);
    
    // Add exchange comparison table
    const tableData = [
        [{{'text': 'Exchange', 'options': {{'fill': {{'color': '2E4053'}}, 'color': 'FFFFFF', 'bold': true}}}},
         {{'text': 'Total Pairs', 'options': {{'fill': {{'color': '2E4053'}}, 'color': 'FFFFFF', 'bold': true}}}},
         {{'text': 'USDT Pairs', 'options': {{'fill': {{'color': '2E4053'}}, 'color': 'FFFFFF', 'bold': true}}}},
         {{'text': 'Has OHLCV', 'options': {{'fill': {{'color': '2E4053'}}, 'color': 'FFFFFF', 'bold': true}}}}]
    ];
    
    const exchanges = {json.dumps(exchange_results.to_dict('records'))};
    exchanges.forEach(ex => {{
        tableData.push([
            ex.name,
            ex.total_spot_pairs.toString(),
            ex.usdt_quoted_pairs.toString(),
            ex.supports_fetchOHLCV ? 'Yes' : 'No'
        ]);
    }});
    
    slide3.addTable(tableData, {{
        ...p3[0],
        border: {{ pt: 1, color: 'CCCCCC' }},
        fontSize: 16,
        align: 'center',
        valign: 'middle'
    }});
    
    // Slide 4: Chart
    await html2pptx("slide4.html", pptx);
    
    // Slide 5: What This Means
    await html2pptx("slide5.html", pptx);
    
    // Slide 6: Educational - MA
    await html2pptx("slide6.html", pptx);
    
    // Slide 7: Educational - RSI
    await html2pptx("slide7.html", pptx);
    
    await pptx.writeFile("/mnt/user-data/outputs/Crypto_Market_Analysis.pptx");
    console.log("PowerPoint generated successfully!");
}})();
"""

    with open('generate_pptx.js', 'w') as f:
        f.write(js_code)

    # Copy chart to working directory if not already here
    if os.path.exists(chart_path):
        import shutil
        chart_filename = os.path.basename(chart_path)
        # Only copy if source and destination are different
        if os.path.abspath(chart_path) != os.path.abspath(chart_filename):
            shutil.copy(chart_path, chart_filename)

    # Run the script
    result = subprocess.run(
        'NODE_PATH="$(npm root -g)" node generate_pptx.js 2>&1',
        shell=True,
        capture_output=True,
        text=True
    )

    if result.returncode != 0:
        logging.error(f"PowerPoint generation failed: {result.stderr}")
    else:
        logging.info("PowerPoint report generated successfully!")

def generate_word_report(
        signal_data: dict,
        exchange_results: pd.DataFrame,
        symbol: str,
        suggestions: list,
        df: pd.DataFrame
):
    """
    Generates a comprehensive Word document report.
    """
    logging.info("Generating Word document report...")

    # Create JavaScript for Word document
    js_code = f"""
const {{ Document, Packer, Paragraph, TextRun, Table, TableRow, TableCell, AlignmentType, 
        HeadingLevel, BorderStyle, WidthType, ShadingType, VerticalAlign, LevelFormat, PageBreak }} = require('docx');
const fs = require('fs');

const signal_color = {{'BUY': 'GREEN', 'SELL': 'RED', 'HOLD': 'ORANGE'}};
const signal = '{signal_data['signal']}';
const confidence = '{signal_data['confidence']}';

const tableBorder = {{ style: BorderStyle.SINGLE, size: 1, color: "CCCCCC" }};

const doc = new Document({{
    styles: {{
        default: {{ 
            document: {{ run: {{ font: "Arial", size: 24 }} }} 
        }},
        paragraphStyles: [
            {{ id: "Title", name: "Title", basedOn: "Normal",
               run: {{ size: 56, bold: true, color: "1C2833", font: "Arial" }},
               paragraph: {{ spacing: {{ before: 240, after: 120 }}, alignment: AlignmentType.CENTER }} }},
            {{ id: "Heading1", name: "Heading 1", basedOn: "Normal", quickFormat: true,
               run: {{ size: 32, bold: true, color: "2E4053", font: "Arial" }},
               paragraph: {{ spacing: {{ before: 240, after: 180 }}, outlineLevel: 0 }} }},
            {{ id: "Heading2", name: "Heading 2", basedOn: "Normal", quickFormat: true,
               run: {{ size: 28, bold: true, color: "2E4053", font: "Arial" }},
               paragraph: {{ spacing: {{ before: 180, after: 120 }}, outlineLevel: 1 }} }},
            {{ id: "Heading3", name: "Heading 3", basedOn: "Normal", quickFormat: true,
               run: {{ size: 24, bold: true, color: "34495E", font: "Arial" }},
               paragraph: {{ spacing: {{ before: 120, after: 80 }}, outlineLevel: 2 }} }}
        ]
    }},
    numbering: {{
        config: [
            {{ reference: "bullet-list",
               levels: [{{ level: 0, format: LevelFormat.BULLET, text: "•", alignment: AlignmentType.LEFT,
                          style: {{ paragraph: {{ indent: {{ left: 720, hanging: 360 }} }} }} }}] }},
            {{ reference: "numbered-list",
               levels: [{{ level: 0, format: LevelFormat.DECIMAL, text: "%1.", alignment: AlignmentType.LEFT,
                          style: {{ paragraph: {{ indent: {{ left: 720, hanging: 360 }} }} }} }}] }}
        ]
    }},
    sections: [{{
        properties: {{
            page: {{ margin: {{ top: 1440, right: 1440, bottom: 1440, left: 1440 }} }}
        }},
        children: [
            // Title
            new Paragraph({{
                heading: HeadingLevel.TITLE,
                children: [new TextRun("Cryptocurrency Market Analysis Report")]
            }}),
            new Paragraph({{
                alignment: AlignmentType.CENTER,
                spacing: {{ after: 400 }},
                children: [new TextRun({{
                    text: "{symbol}",
                    size: 32,
                    color: "7F8C8D"
                }})]
            }}),
            
            // Executive Summary
            new Paragraph({{
                heading: HeadingLevel.HEADING_1,
                children: [new TextRun("Executive Summary")]
            }}),
            new Paragraph({{
                spacing: {{ after: 200 }},
                children: [
                    new TextRun("Trading Recommendation: "),
                    new TextRun({{
                        text: signal,
                        bold: true,
                        size: 32,
                        color: signal === 'BUY' ? '2ECC71' : signal === 'SELL' ? 'E74C3C' : 'F39C12'
                    }})
                ]
            }}),
            new Paragraph({{
                spacing: {{ after: 200 }},
                children: [
                    new TextRun("Confidence Level: "),
                    new TextRun({{ text: confidence, bold: true, size: 24 }})
                ]
            }}),
            new Paragraph({{
                spacing: {{ after: 300 }},
                children: [new TextRun({{
                    text: "{signal_data['reasoning']}",
                    italics: true
                }})]
            }}),
            
            new Paragraph({{
                spacing: {{ after: 200 }},
                children: [
                    new TextRun("Current Price: "),
                    new TextRun({{
                        text: "${signal_data['price']:,.2f}",
                        bold: true,
                        size: 28
                    }})
                ]
            }}),
            new Paragraph({{
                spacing: {{ after: 400 }},
                children: [
                    new TextRun("RSI: "),
                    new TextRun({{
                        text: "{signal_data['rsi']:.1f}",
                        bold: true,
                        size: 28
                    }})
                ]
            }}),
            
            // Page Break
            new Paragraph({{ children: [new PageBreak()] }}),
            
            // Exchange Analysis
            new Paragraph({{
                heading: HeadingLevel.HEADING_1,
                children: [new TextRun("Exchange Comparison")]
            }}),
            new Paragraph({{
                spacing: {{ after: 200 }},
                children: [new TextRun("Below is a comparison of exchanges available for trading. Choose exchanges with good USDT pair availability and OHLCV support for technical analysis.")]
            }}),
"""

    # Add exchange table
    js_code += """
            new Table({
                columnWidths: [2340, 2340, 2340, 2340],
                margins: { top: 100, bottom: 100, left: 180, right: 180 },
                rows: [
                    new TableRow({
                        tableHeader: true,
                        children: [
                            new TableCell({
                                borders: { top: tableBorder, bottom: tableBorder, left: tableBorder, right: tableBorder },
                                width: { size: 2340, type: WidthType.DXA },
                                shading: { fill: "2E4053", type: ShadingType.CLEAR },
                                verticalAlign: VerticalAlign.CENTER,
                                children: [new Paragraph({
                                    alignment: AlignmentType.CENTER,
                                    children: [new TextRun({ text: "Exchange", bold: true, color: "FFFFFF", size: 22 })]
                                })]
                            }),
                            new TableCell({
                                borders: { top: tableBorder, bottom: tableBorder, left: tableBorder, right: tableBorder },
                                width: { size: 2340, type: WidthType.DXA },
                                shading: { fill: "2E4053", type: ShadingType.CLEAR },
                                children: [new Paragraph({
                                    alignment: AlignmentType.CENTER,
                                    children: [new TextRun({ text: "Total Pairs", bold: true, color: "FFFFFF", size: 22 })]
                                })]
                            }),
                            new TableCell({
                                borders: { top: tableBorder, bottom: tableBorder, left: tableBorder, right: tableBorder },
                                width: { size: 2340, type: WidthType.DXA },
                                shading: { fill: "2E4053", type: ShadingType.CLEAR },
                                children: [new Paragraph({
                                    alignment: AlignmentType.CENTER,
                                    children: [new TextRun({ text: "USDT Pairs", bold: true, color: "FFFFFF", size: 22 })]
                                })]
                            }),
                            new TableCell({
                                borders: { top: tableBorder, bottom: tableBorder, left: tableBorder, right: tableBorder },
                                width: { size: 2340, type: WidthType.DXA },
                                shading: { fill: "2E4053", type: ShadingType.CLEAR },
                                children: [new Paragraph({
                                    alignment: AlignmentType.CENTER,
                                    children: [new TextRun({ text: "Has OHLCV", bold: true, color: "FFFFFF", size: 22 })]
                                })]
                            })
                        ]
                    }),
"""

    # Add exchange data rows
    for _, row in exchange_results.iterrows():
        js_code += f"""
                    new TableRow({{
                        children: [
                            new TableCell({{
                                borders: {{ top: tableBorder, bottom: tableBorder, left: tableBorder, right: tableBorder }},
                                width: {{ size: 2340, type: WidthType.DXA }},
                                children: [new Paragraph({{ children: [new TextRun("{row['name']}")] }})]
                            }}),
                            new TableCell({{
                                borders: {{ top: tableBorder, bottom: tableBorder, left: tableBorder, right: tableBorder }},
                                width: {{ size: 2340, type: WidthType.DXA }},
                                children: [new Paragraph({{ alignment: AlignmentType.CENTER, children: [new TextRun("{row['total_spot_pairs']}")] }})]
                            }}),
                            new TableCell({{
                                borders: {{ top: tableBorder, bottom: tableBorder, left: tableBorder, right: tableBorder }},
                                width: {{ size: 2340, type: WidthType.DXA }},
                                children: [new Paragraph({{ alignment: AlignmentType.CENTER, children: [new TextRun("{row['usdt_quoted_pairs']}")] }})]
                            }}),
                            new TableCell({{
                                borders: {{ top: tableBorder, bottom: tableBorder, left: tableBorder, right: tableBorder }},
                                width: {{ size: 2340, type: WidthType.DXA }},
                                children: [new Paragraph({{ alignment: AlignmentType.CENTER, children: [new TextRun("{"Yes" if row['supports_fetchOHLCV'] else "No"}")] }})]
                            }})
                        ]
                    }}),
"""

    js_code += """
                ]
            }),
            
            // Page Break
            new Paragraph({ children: [new PageBreak()] }),
            
            // Detailed Analysis
            new Paragraph({
                heading: HeadingLevel.HEADING_1,
                children: [new TextRun("Technical Analysis Details")]
            }),
"""

    # Add suggestions
    for i, suggestion in enumerate(suggestions):
        js_code += f"""
            new Paragraph({{
                spacing: {{ after: 150 }},
                numbering: {{ reference: "bullet-list", level: 0 }},
                children: [new TextRun("{suggestion}")]
            }}),
"""

    js_code += """
            
            // Page Break
            new Paragraph({ children: [new PageBreak()] }),
            
            // Educational Section
            new Paragraph({
                heading: HeadingLevel.HEADING_1,
                children: [new TextRun("Educational Guide for Beginners")]
            }}),
            
            new Paragraph({
                heading: HeadingLevel.HEADING_2,
                children: [new TextRun("Understanding Moving Averages")]
            }}),
            new Paragraph({
                spacing: { after: 200 },
                children: [new TextRun("Moving averages are one of the most fundamental tools in technical analysis. They smooth out price data by averaging prices over a specific period, making it easier to identify trends.")]
            }}),
            new Paragraph({
                heading: HeadingLevel.HEADING_3,
                children: [new TextRun("Key Concepts:")]
            }}),
            new Paragraph({
                numbering: { reference: "bullet-list", level: 0 },
                children: [new TextRun("Golden Cross: When a short-term moving average crosses above a long-term moving average, it signals a potential uptrend (bullish signal).")]
            }),
            new Paragraph({
                numbering: { reference: "bullet-list", level: 0 },
                children: [new TextRun("Death Cross: When a short-term moving average crosses below a long-term moving average, it signals a potential downtrend (bearish signal).")]
            }),
            new Paragraph({
                spacing: { after: 300 },
                numbering: { reference: "bullet-list", level: 0 },
                children: [new TextRun("Trend Confirmation: When price stays above the moving average, it suggests an uptrend. When price stays below, it suggests a downtrend.")]
            }),
            
            new Paragraph({
                heading: HeadingLevel.HEADING_2,
                children: [new TextRun("Understanding RSI (Relative Strength Index)")]
            }}),
            new Paragraph({
                spacing: { after: 200 },
                children: [new TextRun("RSI measures the speed and magnitude of price changes on a scale from 0 to 100. It helps identify overbought or oversold conditions.")]
            }}),
            new Paragraph({
                heading: HeadingLevel.HEADING_3,
                children: [new TextRun("How to Read RSI:")]
            }}),
            new Paragraph({
                numbering: { reference: "bullet-list", level: 0 },
                children: [new TextRun("Above 70: The asset is considered overbought, meaning it may be due for a price correction or pullback.")]
            }),
            new Paragraph({
                numbering: { reference: "bullet-list", level: 0 },
                children: [new TextRun("Below 30: The asset is considered oversold, meaning it may be due for a price bounce or recovery.")]
            }),
            new Paragraph({
                numbering: { reference: "bullet-list", level: 0 },
                children: [new TextRun("Between 30-70: This is the neutral zone where no extreme conditions are present.")]
            }),
            new Paragraph({
                spacing: { after: 300 },
                numbering: { reference: "bullet-list", level: 0 },
                children: [new TextRun("Divergence: When price makes a new high but RSI doesn't, it can signal weakness and a potential reversal.")]
            }),
            
            new Paragraph({
                heading: HeadingLevel.HEADING_2,
                children: [new TextRun("Understanding Bollinger Bands")]
            }),
            new Paragraph({
                spacing: { after: 200 },
                children: [new TextRun("Bollinger Bands consist of three lines: a middle band (moving average) and two outer bands that represent standard deviations from the middle band.")]
            }),
            new Paragraph({
                heading: HeadingLevel.HEADING_3,
                children: [new TextRun("How to Use Bollinger Bands:")]
            }),
            new Paragraph({
                numbering: { reference: "bullet-list", level: 0 },
                children: [new TextRun("Price touching upper band: May indicate overbought conditions.")]
            }),
            new Paragraph({
                numbering: { reference: "bullet-list", level: 0 },
                children: [new TextRun("Price touching lower band: May indicate oversold conditions.")]
            }),
            new Paragraph({
                numbering: { reference: "bullet-list", level: 0 },
                children: [new TextRun("Band squeeze: When bands narrow, it suggests low volatility and a potential breakout coming.")]
            }),
            new Paragraph({
                numbering: { reference: "bullet-list", level: 0 },
                children: [new TextRun("Band expansion: Wide bands indicate high volatility.")]
            }),
            
            // Page Break
            new Paragraph({ children: [new PageBreak()] }),
            
            // Getting Started
            new Paragraph({
                heading: HeadingLevel.HEADING_1,
                children: [new TextRun("Getting Started with Crypto Trading")]
            }}),
            
            new Paragraph({
                heading: HeadingLevel.HEADING_2,
                children: [new TextRun("Step 1: Choose Your Exchange")]
            }),
            new Paragraph({
                spacing: { after: 200 },
                children: [new TextRun("Based on the exchange comparison in this report, select an exchange with good USDT pair availability and technical data support. Popular choices include Binance, KuCoin, and Bybit.")]
            }),
            
            new Paragraph({
                heading: HeadingLevel.HEADING_2,
                children: [new TextRun("Step 2: Start Small")]
            }),
            new Paragraph({
                spacing: { after: 200 },
                children: [new TextRun("Begin with small amounts that you can afford to lose. Never invest money you need for essential expenses. As a beginner, consider starting with 1-5% of your total investment capital on any single trade.")]
            }),
            
            new Paragraph({
                heading: HeadingLevel.HEADING_2,
                children: [new TextRun("Step 3: Set Stop-Losses")]
            }),
            new Paragraph({
                spacing: { after: 200 },
                children: [new TextRun("Always use stop-loss orders to limit potential losses. A common approach is to set a stop-loss at 5-10% below your entry price for long positions.")]
            }),
            
            new Paragraph({
                heading: HeadingLevel.HEADING_2,
                children: [new TextRun("Step 4: Keep Learning")]
            }),
            new Paragraph({
                spacing: { after: 200 },
                children: [new TextRun("Continue educating yourself about technical analysis, market trends, and risk management. The crypto market is volatile and requires ongoing learning and adaptation.")]
            }),
            
            new Paragraph({
                heading: HeadingLevel.HEADING_2,
                children: [new TextRun("Important Disclaimers")]
            }),
            new Paragraph({
                numbering: { reference: "bullet-list", level: 0 },
                children: [new TextRun("This report is for educational purposes only and is not financial advice.")]
            }),
            new Paragraph({
                numbering: { reference: "bullet-list", level: 0 },
                children: [new TextRun("Past performance does not guarantee future results.")]
            }),
            new Paragraph({
                numbering: { reference: "bullet-list", level: 0 },
                children: [new TextRun("Cryptocurrency trading involves substantial risk of loss.")]
            }),
            new Paragraph({
                numbering: { reference: "bullet-list", level: 0 },
                children: [new TextRun("Always do your own research (DYOR) before making any investment decisions.")]
            }),
            new Paragraph({
                numbering: { reference: "bullet-list", level: 0 },
                children: [new TextRun("Consider consulting with a qualified financial advisor before trading.")]
            })
        ]
    }]
}});

Packer.toBuffer(doc).then(buffer => {
    fs.writeFileSync("/mnt/user-data/outputs/Crypto_Market_Analysis.docx", buffer);
    console.log("Word document generated successfully!");
});
"""

    with open('generate_docx.js', 'w') as f:
        f.write(js_code)

    # Run the script
    result = subprocess.run(
        'node generate_docx.js 2>&1',
        shell=True,
        capture_output=True,
        text=True
    )

    if result.returncode != 0:
        logging.error(f"Word document generation failed: {result.stderr}")
    else:
        logging.info("Word document report generated successfully!")

def generate_comprehensive_reports(
        exchange_results: pd.DataFrame,
        symbol: str,
        df: pd.DataFrame,
        suggestions: list,
        chart_path: str
):
    """
    Main function to generate both PowerPoint and Word reports.
    """
    # Determine trading signal
    signal_data = determine_trading_signal(df)

    logging.info(f"\n{'='*60}")
    logging.info(f"TRADING SIGNAL: {signal_data['signal']} (Confidence: {signal_data['confidence']})")
    logging.info(f"{'='*60}")
    logging.info(f"Reasoning: {signal_data['reasoning']}")
    logging.info(f"{'='*60}\n")

    # Generate both reports
    generate_powerpoint_report(signal_data, exchange_results, symbol, df, chart_path)
    generate_word_report(signal_data, exchange_results, symbol, suggestions, df)

    return signal_data