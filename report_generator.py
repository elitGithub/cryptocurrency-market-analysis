#
# File: report_generator.py
# Description: Module for generating comprehensive PowerPoint and Word reports
#

import logging
from datetime import datetime
import subprocess
import os

def determine_trading_signal(suggestions: list, df, latest_price: float) -> dict:
    """
    Analyzes suggestions and indicators to provide a clear trading signal.
    
    Args:
        suggestions: List of analysis suggestions
        df: DataFrame with indicators
        latest_price: Current price
    
    Returns:
        dict: Trading signal with action, confidence, and reasoning
    """
    if df.empty or len(df) < 2:
        return {
            'action': 'HOLD',
            'confidence': 'LOW',
            'reasoning': 'Insufficient data for analysis',
            'price': latest_price
        }
    
    latest = df.iloc[-1]
    previous = df.iloc[-2]
    
    # Scoring system
    bullish_score = 0
    bearish_score = 0
    signals = []
    
    # Check moving average crossover
    if latest['SMA_short'] > latest['SMA_long']:
        if previous['SMA_short'] <= previous['SMA_long']:
            # Golden cross just happened
            bullish_score += 3
            signals.append("Golden Cross detected (strong bullish)")
        else:
            # In uptrend
            bullish_score += 1
            signals.append("Price above long-term average (bullish)")
    else:
        if previous['SMA_short'] >= previous['SMA_long']:
            # Death cross just happened
            bearish_score += 3
            signals.append("Death Cross detected (strong bearish)")
        else:
            # In downtrend
            bearish_score += 1
            signals.append("Price below long-term average (bearish)")
    
    # Check RSI
    rsi = latest['RSI']
    if rsi > 70:
        bearish_score += 2
        signals.append(f"Overbought RSI ({rsi:.1f}) suggests caution")
    elif rsi < 30:
        bullish_score += 2
        signals.append(f"Oversold RSI ({rsi:.1f}) suggests buying opportunity")
    elif 40 <= rsi <= 60:
        signals.append(f"Neutral RSI ({rsi:.1f})")
    
    # Check Bollinger Bands
    if latest['close'] > latest['BB_upper']:
        bearish_score += 1
        signals.append("Price above upper Bollinger Band (overbought)")
    elif latest['close'] < latest['BB_lower']:
        bullish_score += 1
        signals.append("Price below lower Bollinger Band (oversold)")
    
    # Determine action
    net_score = bullish_score - bearish_score
    
    if net_score >= 3:
        action = 'STRONG BUY'
        confidence = 'HIGH'
    elif net_score >= 1:
        action = 'BUY'
        confidence = 'MEDIUM'
    elif net_score <= -3:
        action = 'STRONG SELL'
        confidence = 'HIGH'
    elif net_score <= -1:
        action = 'SELL'
        confidence = 'MEDIUM'
    else:
        action = 'HOLD'
        confidence = 'LOW'
    
    return {
        'action': action,
        'confidence': confidence,
        'reasoning': ' | '.join(signals),
        'price': latest_price,
        'rsi': rsi,
        'trend': 'UPTREND' if latest['SMA_short'] > latest['SMA_long'] else 'DOWNTREND'
    }


def generate_word_report(symbol: str, exchange_results: list, signal: dict, 
                        suggestions: list, df, short_ma: int, long_ma: int):
    """
    Generates a comprehensive Word document report.
    
    Args:
        symbol: Trading symbol
        exchange_results: List of exchange analysis results
        signal: Trading signal dictionary
        suggestions: List of analysis suggestions
        df: DataFrame with indicators
        short_ma: Short MA period
        long_ma: Long MA period
    """
    logging.info("Generating Word document report...")
    
    # Create JavaScript file for Word document generation
    js_code = f"""
const {{ Document, Packer, Paragraph, TextRun, HeadingLevel, AlignmentType, BorderStyle }} = require('docx');
const fs = require('fs');

const doc = new Document({{
    sections: [{{
        properties: {{}},
        children: [
            // Title
            new Paragraph({{
                text: "Cryptocurrency Market Analysis Report",
                heading: HeadingLevel.TITLE,
                alignment: AlignmentType.CENTER,
                spacing: {{ after: 400 }}
            }}),
            
            new Paragraph({{
                children: [
                    new TextRun({{
                        text: "Generated: {datetime.now().strftime('%B %d, %Y at %I:%M %p')}",
                        italics: true
                    }})
                ],
                alignment: AlignmentType.CENTER,
                spacing: {{ after: 400 }}
            }}),
            
            // Executive Summary
            new Paragraph({{
                text: "EXECUTIVE SUMMARY",
                heading: HeadingLevel.HEADING_1,
                spacing: {{ before: 400, after: 200 }},
                border: {{
                    bottom: {{ color: "2E75B6", space: 1, value: BorderStyle.SINGLE, size: 12 }}
                }}
            }}),
            
            new Paragraph({{
                children: [
                    new TextRun({{
                        text: "Asset: ",
                        bold: true
                    }}),
                    new TextRun({{
                        text: "{symbol}"
                    }})
                ],
                spacing: {{ after: 100 }}
            }}),
            
            new Paragraph({{
                children: [
                    new TextRun({{
                        text: "Current Price: ",
                        bold: true
                    }}),
                    new TextRun({{
                        text: "${signal['price']:.2f}"
                    }})
                ],
                spacing: {{ after: 100 }}
            }}),
            
            new Paragraph({{
                children: [
                    new TextRun({{
                        text: "Trend: ",
                        bold: true
                    }}),
                    new TextRun({{
                        text: "{signal['trend']}",
                        color: "{('70AD47' if signal['trend'] == 'UPTREND' else 'C00000')}"
                    }})
                ],
                spacing: {{ after: 100 }}
            }}),
            
            // Trading Recommendation Box
            new Paragraph({{
                text: "",
                spacing: {{ before: 200, after: 200 }}
            }}),
            
            new Paragraph({{
                children: [
                    new TextRun({{
                        text: "📊 TRADING RECOMMENDATION",
                        bold: true,
                        size: 28
                    }})
                ],
                spacing: {{ after: 100 }}
            }}),
            
            new Paragraph({{
                children: [
                    new TextRun({{
                        text: "Action: ",
                        bold: true,
                        size: 24
                    }}),
                    new TextRun({{
                        text: "{signal['action']}",
                        bold: true,
                        size: 24,
                        color: "{('70AD47' if 'BUY' in signal['action'] else 'C00000' if 'SELL' in signal['action'] else '808080')}"
                    }})
                ],
                spacing: {{ after: 100 }}
            }}),
            
            new Paragraph({{
                children: [
                    new TextRun({{
                        text: "Confidence: ",
                        bold: true
                    }}),
                    new TextRun({{
                        text: "{signal['confidence']}"
                    }})
                ],
                spacing: {{ after: 100 }}
            }}),
            
            new Paragraph({{
                children: [
                    new TextRun({{
                        text: "Reasoning: ",
                        bold: true
                    }}),
                    new TextRun({{
                        text: "{signal['reasoning']}"
                    }})
                ],
                spacing: {{ after: 300 }}
            }}),
            
            // Exchange Comparison
            new Paragraph({{
                text: "EXCHANGE COMPARISON",
                heading: HeadingLevel.HEADING_1,
                spacing: {{ before: 400, after: 200 }},
                border: {{
                    bottom: {{ color: "2E75B6", space: 1, value: BorderStyle.SINGLE, size: 12 }}
                }}
            }}),
            
            new Paragraph({{
                text: "Based on the analysis of multiple exchanges, here are the key findings:",
                spacing: {{ after: 200 }}
            }}),
"""
    
    # Add exchange data
    for result in exchange_results:
        js_code += f"""
            new Paragraph({{
                children: [
                    new TextRun({{
                        text: "{result['name'].upper()}",
                        bold: true,
                        size: 24
                    }})
                ],
                spacing: {{ before: 200, after: 100 }}
            }}),
            
            new Paragraph({{
                children: [
                    new TextRun({{
                        text: "  • Total Spot Pairs: ",
                        bold: true
                    }}),
                    new TextRun({{
                        text: "{result['total_spot_pairs']}"
                    }})
                ],
                spacing: {{ after: 50 }}
            }}),
            
            new Paragraph({{
                children: [
                    new TextRun({{
                        text: "  • USDT Quoted Pairs: ",
                        bold: true
                    }}),
                    new TextRun({{
                        text: "{result['usdt_quoted_pairs']}"
                    }})
                ],
                spacing: {{ after: 50 }}
            }}),
            
            new Paragraph({{
                children: [
                    new TextRun({{
                        text: "  • Historical Data Available: ",
                        bold: true
                    }}),
                    new TextRun({{
                        text: "{'Yes' if result['supports_fetchOHLCV'] else 'No'}",
                        color: "{('70AD47' if result['supports_fetchOHLCV'] else 'C00000')}"
                    }})
                ],
                spacing: {{ after: 50 }}
            }}),
            
            new Paragraph({{
                children: [
                    new TextRun({{
                        text: "  • Rate Limit: ",
                        bold: true
                    }}),
                    new TextRun({{
                        text: "{result['rate_limit_ms']}ms between requests"
                    }})
                ],
                spacing: {{ after: 100 }}
            }}),
"""
    
    js_code += f"""
            // Technical Analysis Details
            new Paragraph({{
                text: "TECHNICAL ANALYSIS DETAILS",
                heading: HeadingLevel.HEADING_1,
                spacing: {{ before: 400, after: 200 }},
                border: {{
                    bottom: {{ color: "2E75B6", space: 1, value: BorderStyle.SINGLE, size: 12 }}
                }}
            }}),
            
            new Paragraph({{
                text: "The analysis uses several technical indicators to generate trading signals. Here's what the current data shows:",
                spacing: {{ after: 200 }}
            }}),
"""
    
    # Add suggestions
    for i, suggestion in enumerate(suggestions):
        js_code += f"""
            new Paragraph({{
                children: [
                    new TextRun({{
                        text: "{suggestion}",
                    }})
                ],
                spacing: {{ after: 150 }}
            }}),
"""
    
    # Add educational section
    js_code += f"""
            // Educational Section
            new Paragraph({{
                text: "UNDERSTANDING TECHNICAL INDICATORS",
                heading: HeadingLevel.HEADING_1,
                spacing: {{ before: 400, after: 200 }},
                border: {{
                    bottom: {{ color: "2E75B6", space: 1, value: BorderStyle.SINGLE, size: 12 }}
                }}
            }}),
            
            new Paragraph({{
                text: "For beginners, here's what these indicators mean and how they help in trading decisions:",
                spacing: {{ after: 200 }}
            }}),
            
            new Paragraph({{
                text: "Moving Averages (MA)",
                heading: HeadingLevel.HEADING_2,
                spacing: {{ before: 200, after: 100 }}
            }}),
            
            new Paragraph({{
                text: "Moving averages smooth out price data to identify trends. We use two types:",
                spacing: {{ after: 100 }}
            }}),
            
            new Paragraph({{
                children: [
                    new TextRun({{
                        text: "  • Short-term MA ({short_ma}-day): ",
                        bold: true
                    }}),
                    new TextRun({{
                        text: "Reacts quickly to price changes, good for short-term trends."
                    }})
                ],
                spacing: {{ after: 50 }}
            }}),
            
            new Paragraph({{
                children: [
                    new TextRun({{
                        text: "  • Long-term MA ({long_ma}-day): ",
                        bold: true
                    }}),
                    new TextRun({{
                        text: "Shows the overall long-term trend direction."
                    }})
                ],
                spacing: {{ after: 100 }}
            }}),
            
            new Paragraph({{
                children: [
                    new TextRun({{
                        text: "Golden Cross: ",
                        bold: true,
                        color: "70AD47"
                    }}),
                    new TextRun({{
                        text: "When short-term MA crosses above long-term MA = bullish signal (potential upward trend)."
                    }})
                ],
                spacing: {{ after: 50 }}
            }}),
            
            new Paragraph({{
                children: [
                    new TextRun({{
                        text: "Death Cross: ",
                        bold: true,
                        color: "C00000"
                    }}),
                    new TextRun({{
                        text: "When short-term MA crosses below long-term MA = bearish signal (potential downward trend)."
                    }})
                ],
                spacing: {{ after: 200 }}
            }}),
            
            new Paragraph({{
                text: "Relative Strength Index (RSI)",
                heading: HeadingLevel.HEADING_2,
                spacing: {{ before: 200, after: 100 }}
            }}),
            
            new Paragraph({{
                text: "RSI measures momentum on a scale of 0-100. It tells you if an asset might be overbought or oversold:",
                spacing: {{ after: 100 }}
            }}),
            
            new Paragraph({{
                children: [
                    new TextRun({{
                        text: "  • RSI > 70: ",
                        bold: true,
                        color: "C00000"
                    }}),
                    new TextRun({{
                        text: "Overbought - the price may have risen too quickly and could drop soon."
                    }})
                ],
                spacing: {{ after: 50 }}
            }}),
            
            new Paragraph({{
                children: [
                    new TextRun({{
                        text: "  • RSI < 30: ",
                        bold: true,
                        color: "70AD47"
                    }}),
                    new TextRun({{
                        text: "Oversold - the price may have fallen too much and could bounce back."
                    }})
                ],
                spacing: {{ after: 50 }}
            }}),
            
            new Paragraph({{
                children: [
                    new TextRun({{
                        text: "  • RSI 30-70: ",
                        bold: true
                    }}),
                    new TextRun({{
                        text: "Neutral zone - no extreme conditions detected."
                    }})
                ],
                spacing: {{ after: 200 }}
            }}),
            
            new Paragraph({{
                text: "Bollinger Bands",
                heading: HeadingLevel.HEADING_2,
                spacing: {{ before: 200, after: 100 }}
            }}),
            
            new Paragraph({{
                text: "Bollinger Bands show price volatility. They consist of three lines:",
                spacing: {{ after: 100 }}
            }}),
            
            new Paragraph({{
                children: [
                    new TextRun({{
                        text: "  • Upper Band: ",
                        bold: true
                    }}),
                    new TextRun({{
                        text: "When price touches/exceeds this, the asset might be overbought."
                    }})
                ],
                spacing: {{ after: 50 }}
            }}),
            
            new Paragraph({{
                children: [
                    new TextRun({{
                        text: "  • Middle Band: ",
                        bold: true
                    }}),
                    new TextRun({{
                        text: "The simple moving average (average price)."
                    }})
                ],
                spacing: {{ after: 50 }}
            }}),
            
            new Paragraph({{
                children: [
                    new TextRun({{
                        text: "  • Lower Band: ",
                        bold: true
                    }}),
                    new TextRun({{
                        text: "When price touches/falls below this, the asset might be oversold."
                    }})
                ],
                spacing: {{ after: 200 }}
            }}),
            
            // Disclaimer
            new Paragraph({{
                text: "IMPORTANT DISCLAIMER",
                heading: HeadingLevel.HEADING_1,
                spacing: {{ before: 400, after: 200 }},
                border: {{
                    bottom: {{ color: "C00000", space: 1, value: BorderStyle.SINGLE, size: 12 }}
                }}
            }}),
            
            new Paragraph({{
                children: [
                    new TextRun({{
                        text: "This report is for educational and informational purposes only. It is NOT financial advice. Cryptocurrency trading carries substantial risk of loss. Always do your own research and never invest more than you can afford to lose. Past performance does not guarantee future results.",
                        italics: true
                    }})
                ],
                spacing: {{ after: 200 }}
            }})
        ]
    }}]
}});

Packer.toBuffer(doc).then(buffer => {{
    fs.writeFileSync('crypto_analysis_report.docx', buffer);
    console.log('Word document created successfully!');
}});
"""
    
    # Write and execute JavaScript
    with open('generate_docx.js', 'w') as f:
        f.write(js_code)
    
    try:
        result = subprocess.run(
            ['node', 'generate_docx.js'],
            capture_output=True,
            text=True,
            timeout=30
        )
        if result.returncode == 0:
            logging.info("Word document generated successfully!")
        else:
            logging.error(f"Error generating Word document: {result.stderr}")
    except Exception as e:
        logging.error(f"Failed to generate Word document: {e}")


def generate_powerpoint_report(symbol: str, exchange_results: list, signal: dict, 
                               suggestions: list, df, short_ma: int, long_ma: int):
    """
    Generates a comprehensive PowerPoint presentation.
    
    Args:
        symbol: Trading symbol
        exchange_results: List of exchange analysis results
        signal: Trading signal dictionary
        suggestions: List of analysis suggestions
        df: DataFrame with indicators
        short_ma: Short MA period
        long_ma: Long MA period
    """
    logging.info("Generating PowerPoint presentation...")
    
    # Color scheme
    primary_color = "#2E75B6"
    success_color = "#70AD47"
    danger_color = "#C00000"
    neutral_color = "#808080"
    
    # Determine signal color
    if 'BUY' in signal['action']:
        signal_color = success_color
    elif 'SELL' in signal['action']:
        signal_color = danger_color
    else:
        signal_color = neutral_color
    
    # Create HTML slides
    css_content = f"""
:root {{
    --primary: {primary_color};
    --success: {success_color};
    --danger: {danger_color};
    --neutral: {neutral_color};
    --signal-color: {signal_color};
    --bg-dark: #1E1E1E;
    --bg-light: #F5F5F5;
    --text-dark: #2C2C2C;
    --text-light: #FFFFFF;
}}

body {{
    margin: 0;
    padding: 40px;
    font-family: 'Segoe UI', Tahoma, Geneva, Verdana, sans-serif;
    color: var(--text-dark);
    background: white;
    width: 920px;
    height: 500px;
}}

h1 {{
    font-size: 48px;
    font-weight: 700;
    margin: 0 0 20px 0;
    color: var(--primary);
}}

h2 {{
    font-size: 36px;
    font-weight: 600;
    margin: 0 0 15px 0;
    color: var(--primary);
}}

h3 {{
    font-size: 24px;
    font-weight: 600;
    margin: 0 0 10px 0;
}}

p {{
    font-size: 18px;
    line-height: 1.6;
    margin: 0 0 15px 0;
}}

.signal-box {{
    background: linear-gradient(135deg, var(--signal-color) 0%, var(--signal-color)dd 100%);
    padding: 30px;
    border-radius: 15px;
    color: var(--text-light);
    margin: 20px 0;
}}

.signal-action {{
    font-size: 56px;
    font-weight: 700;
    margin: 10px 0;
}}

.exchange-card {{
    background: var(--bg-light);
    padding: 20px;
    border-radius: 10px;
    margin: 15px 0;
    border-left: 5px solid var(--primary);
}}

.metric {{
    display: inline-block;
    margin-right: 30px;
}}

.metric-label {{
    font-weight: 600;
    color: var(--primary);
}}

.bullet {{
    margin: 10px 0;
    padding-left: 20px;
}}

.highlight {{
    background: #FFF3CD;
    padding: 15px;
    border-radius: 8px;
    border-left: 4px solid #FFC107;
    margin: 15px 0;
}}

.center {{
    text-align: center;
}}
"""
    
    # Slide 1: Title Slide
    slide1_html = f"""<!DOCTYPE html>
<html>
<head>
    <style>{css_content}</style>
</head>
<body class="center">
    <div style="padding-top: 100px;">
        <h1>Cryptocurrency Market Analysis</h1>
        <h2 style="color: var(--text-dark); margin-top: 30px;">{symbol}</h2>
        <p style="font-size: 24px; margin-top: 50px; color: var(--neutral);">
            Generated: {datetime.now().strftime('%B %d, %Y')}
        </p>
    </div>
</body>
</html>"""
    
    # Slide 2: Executive Summary
    slide2_html = f"""<!DOCTYPE html>
<html>
<head>
    <style>{css_content}</style>
</head>
<body>
    <h1>Executive Summary</h1>
    
    <div style="margin-top: 40px;">
        <div class="metric">
            <span class="metric-label">Asset:</span> {symbol}
        </div>
        <div class="metric">
            <span class="metric-label">Price:</span> ${signal['price']:.2f}
        </div>
        <div class="metric">
            <span class="metric-label">Trend:</span> 
            <span style="color: {'var(--success)' if signal['trend'] == 'UPTREND' else 'var(--danger)'}; font-weight: 600;">
                {signal['trend']}
            </span>
        </div>
    </div>
    
    <div class="signal-box">
        <p style="font-size: 24px; margin: 0;">Trading Recommendation</p>
        <div class="signal-action">{signal['action']}</div>
        <p style="font-size: 20px;"><strong>Confidence:</strong> {signal['confidence']}</p>
        <p style="font-size: 18px; margin-top: 15px;">{signal['reasoning']}</p>
    </div>
</body>
</html>"""
    
    # Slide 3: Exchange Comparison
    exchange_cards = ""
    for result in exchange_results:
        supports_color = success_color if result['supports_fetchOHLCV'] else danger_color
        exchange_cards += f"""
    <div class="exchange-card">
        <h3>{result['name'].upper()}</h3>
        <p><span class="metric-label">Total Spot Pairs:</span> {result['total_spot_pairs']}</p>
        <p><span class="metric-label">USDT Quoted Pairs:</span> {result['usdt_quoted_pairs']}</p>
        <p><span class="metric-label">Historical Data:</span> 
            <span style="color: {supports_color}; font-weight: 600;">
                {'✓ Available' if result['supports_fetchOHLCV'] else '✗ Not Available'}
            </span>
        </p>
        <p><span class="metric-label">Rate Limit:</span> {result['rate_limit_ms']}ms</p>
    </div>"""
    
    slide3_html = f"""<!DOCTYPE html>
<html>
<head>
    <style>{css_content}</style>
</head>
<body>
    <h1>Exchange Comparison</h1>
    <p>Analysis of multiple cryptocurrency exchanges:</p>
    {exchange_cards}
</body>
</html>"""
    
    # Slide 4: Technical Analysis
    bullet_points = ""
    for suggestion in suggestions[:3]:  # Limit to first 3 for slide readability
        bullet_points += f'    <div class="bullet">• {suggestion}</div>\n'
    
    slide4_html = f"""<!DOCTYPE html>
<html>
<head>
    <style>{css_content}</style>
</head>
<body>
    <h1>Technical Analysis</h1>
    <p>Key findings from indicator analysis:</p>
    <div style="margin-top: 30px;">
{bullet_points}
    </div>
    
    <div class="highlight" style="margin-top: 40px;">
        <p style="margin: 0;"><strong>Current RSI:</strong> {signal['rsi']:.1f}</p>
        <p style="margin: 5px 0 0 0; font-size: 16px;">
            {'Overbought territory - exercise caution' if signal['rsi'] > 70 else 
             'Oversold territory - potential opportunity' if signal['rsi'] < 30 else
             'Neutral zone - no extreme conditions'}
        </p>
    </div>
</body>
</html>"""
    
    # Slide 5: Understanding Moving Averages
    slide5_html = f"""<!DOCTYPE html>
<html>
<head>
    <style>{css_content}</style>
</head>
<body>
    <h1>Understanding Moving Averages</h1>
    
    <div style="margin-top: 30px;">
        <h3>What are Moving Averages?</h3>
        <p>Moving averages smooth out price data to identify trends over time.</p>
        
        <div class="bullet" style="margin-top: 20px;">
            <strong>Short-term MA ({short_ma}-day):</strong> Reacts quickly to price changes, good for identifying short-term trends and entry/exit points.
        </div>
        
        <div class="bullet">
            <strong>Long-term MA ({long_ma}-day):</strong> Shows the overall long-term trend direction and market sentiment.
        </div>
        
        <div class="highlight" style="margin-top: 30px;">
            <p style="margin: 0; font-size: 20px;"><strong style="color: var(--success);">Golden Cross ✓</strong></p>
            <p style="margin: 5px 0 0 0;">Short-term MA crosses above long-term MA = Bullish signal (potential upward trend)</p>
        </div>
        
        <div class="highlight">
            <p style="margin: 0; font-size: 20px;"><strong style="color: var(--danger);">Death Cross ✗</strong></p>
            <p style="margin: 5px 0 0 0;">Short-term MA crosses below long-term MA = Bearish signal (potential downward trend)</p>
        </div>
    </div>
</body>
</html>"""
    
    # Slide 6: Understanding RSI
    slide6_html = """<!DOCTYPE html>
<html>
<head>
    <style>{css_content}</style>
</head>
<body>
    <h1>Understanding RSI</h1>
    
    <div style="margin-top: 30px;">
        <h3>Relative Strength Index (RSI)</h3>
        <p>RSI measures momentum on a scale of 0-100, showing if an asset is overbought or oversold.</p>
        
        <div style="margin-top: 40px;">
            <div class="highlight" style="background: #FFE6E6; border-color: var(--danger);">
                <p style="margin: 0; font-size: 22px;"><strong style="color: var(--danger);">RSI > 70: Overbought</strong></p>
                <p style="margin: 5px 0 0 0;">Price may have risen too quickly and could drop soon. Consider selling or waiting.</p>
            </div>
            
            <div class="highlight" style="background: #E6F3E6; border-color: var(--success); margin-top: 20px;">
                <p style="margin: 0; font-size: 22px;"><strong style="color: var(--success);">RSI < 30: Oversold</strong></p>
                <p style="margin: 5px 0 0 0;">Price may have fallen too much and could bounce back. Potential buying opportunity.</p>
            </div>
            
            <div class="highlight" style="background: #F0F0F0; border-color: var(--neutral); margin-top: 20px;">
                <p style="margin: 0; font-size: 22px;"><strong style="color: var(--neutral);">RSI 30-70: Neutral</strong></p>
                <p style="margin: 5px 0 0 0;">No extreme conditions detected. Normal market behavior.</p>
            </div>
        </div>
    </div>
</body>
</html>"""
    
    # Slide 7: Understanding Bollinger Bands
    slide7_html = """<!DOCTYPE html>
<html>
<head>
    <style>{css_content}</style>
</head>
<body>
    <h1>Understanding Bollinger Bands</h1>
    
    <div style="margin-top: 30px;">
        <h3>What are Bollinger Bands?</h3>
        <p>Bollinger Bands show price volatility and consist of three lines around the price:</p>
        
        <div style="margin-top: 40px;">
            <div class="bullet">
                <p style="margin: 0;"><strong style="color: var(--danger); font-size: 20px;">Upper Band</strong></p>
                <p style="margin: 5px 0 0 0;">When price touches or exceeds this line, the asset might be overbought. Price may pull back toward the middle.</p>
            </div>
            
            <div class="bullet" style="margin-top: 20px;">
                <p style="margin: 0;"><strong style="font-size: 20px;">Middle Band</strong></p>
                <p style="margin: 5px 0 0 0;">The simple moving average - represents the "fair" price based on recent history.</p>
            </div>
            
            <div class="bullet" style="margin-top: 20px;">
                <p style="margin: 0;"><strong style="color: var(--success); font-size: 20px;">Lower Band</strong></p>
                <p style="margin: 5px 0 0 0;">When price touches or falls below this line, the asset might be oversold. Price may bounce back toward the middle.</p>
            </div>
        </div>
        
        <div class="highlight" style="margin-top: 30px;">
            <p style="margin: 0;"><strong>Pro Tip:</strong> Wider bands = higher volatility. Narrower bands = lower volatility. Bands squeezing together often precede a big price move.</p>
        </div>
    </div>
</body>
</html>"""
    
    # Slide 8: Disclaimer
    slide8_html = """<!DOCTYPE html>
<html>
<head>
    <style>{css_content}</style>
</head>
<body class="center">
    <h1 style="color: var(--danger);">Important Disclaimer</h1>
    
    <div style="padding: 50px; text-align: left; max-width: 800px; margin: 40px auto;">
        <div class="highlight" style="background: #FFE6E6; border-color: var(--danger);">
            <p style="font-size: 20px; line-height: 1.8;">
                This report is for <strong>educational and informational purposes only</strong>. 
                It is <strong>NOT financial advice</strong>.
            </p>
            <p style="font-size: 18px; line-height: 1.8; margin-top: 20px;">
                Cryptocurrency trading carries <strong>substantial risk of loss</strong>. 
                Always do your own research and never invest more than you can afford to lose. 
                Past performance does not guarantee future results.
            </p>
            <p style="font-size: 18px; line-height: 1.8; margin-top: 20px;">
                Consult with a qualified financial advisor before making any investment decisions.
            </p>
        </div>
    </div>
</body>
</html>"""
    
    # Write HTML files
    slides = [
        ('slide1.html', slide1_html),
        ('slide2.html', slide2_html),
        ('slide3.html', slide3_html),
        ('slide4.html', slide4_html),
        ('slide5.html', slide5_html),
        ('slide6.html', slide6_html),
        ('slide7.html', slide7_html),
        ('slide8.html', slide8_html),
    ]
    
    for filename, content in slides:
        with open(f'{filename}', 'w') as f:
            f.write(content)
    
    # Generate PowerPoint using html2pptx
    js_pptx_code = """
const pptxgen = require("pptxgenjs");
const { html2pptx } = require("@ant/html2pptx");

async function generatePresentation() {
    const pptx = new pptxgen();
    pptx.layout = "LAYOUT_16x9";
    pptx.author = "Crypto Trading Bot";
    pptx.title = "Cryptocurrency Market Analysis";
    pptx.subject = "Technical Analysis Report";
    
    // Add all slides
    await html2pptx("slide1.html", pptx);
    await html2pptx("slide2.html", pptx);
    await html2pptx("slide3.html", pptx);
    await html2pptx("slide4.html", pptx);
    await html2pptx("slide5.html", pptx);
    await html2pptx("slide6.html", pptx);
    await html2pptx("slide7.html", pptx);
    await html2pptx("slide8.html", pptx);
    
    await pptx.writeFile("crypto_analysis_presentation.pptx");
    console.log("PowerPoint presentation created successfully!");
}

generatePresentation().catch(console.error);
"""
    
    with open('generate_pptx.js', 'w') as f:
        f.write(js_pptx_code)
    
    try:
        # Check if html2pptx is installed
        check_result = subprocess.run(
            ['npm', 'list', '-g', '@ant/html2pptx'],
            capture_output=True,
            text=True
        )
        
        if '@ant/html2pptx' not in check_result.stdout:
            logging.info("Installing html2pptx...")
            subprocess.run(
                ['npm', 'install', '-g', '/mnt/skills/public/pptx/html2pptx.tgz'],
                capture_output=True,
                text=True,
                timeout=60
            )
        
        result = subprocess.run(
            ['bash', '-c', 'cd /home/claude && NODE_PATH="$(npm root -g)" node generate_pptx.js 2>&1'],
            capture_output=True,
            text=True,
            timeout=60
        )
        
        if result.returncode == 0:
            logging.info("PowerPoint presentation generated successfully!")
        else:
            logging.error(f"Error generating PowerPoint: {result.stderr}")
    except Exception as e:
        logging.error(f"Failed to generate PowerPoint: {e}")
