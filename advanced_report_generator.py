#
# File: advanced_report_generator.py
# Description: Generates comprehensive PowerPoint, Word, and HTML reports with trading signals
#

import pandas as pd
import logging
import subprocess
import json
import os
import base64

def determine_trading_signal(df: pd.DataFrame | None) -> dict:
    """
    Analyzes indicators and generates a clear BUY/SELL/HOLD signal.
    Handles potential missing data or NaNs gracefully.

    Args:
        df: A DataFrame with technical indicators (can be None or empty)

    Returns:
        A dict with signal, confidence, and reasoning
    """
    # Check if DataFrame is valid
    if df is None:
        logging.warning("Insufficient data (DataFrame is None) for signal determination.")
        return {
            'signal': 'HOLD', 'confidence': 'LOW',
            'reasoning': 'Insufficient data for analysis (DataFrame is None)',
            'score': 0, 'rsi': 50.0, 'price': 0.0
        }

    # Check for enough data
    if df.empty or len(df) < 2:
        logging.warning("Insufficient data (Empty or < 2 rows) for signal determination.")
        return {
            'signal': 'HOLD', 'confidence': 'LOW',
            'reasoning': 'Insufficient data for analysis (Empty or < 2 rows)',
            'score': 0, 'rsi': 50.0, 'price': 0.0
        }

    latest = df.iloc[-1]
    previous = df.iloc[-2]

    # Scoring system
    score = 0
    reasons = []
    required_ma_cols = ['SMA_short', 'SMA_long']
    required_rsi_cols = ['RSI']
    required_bb_cols = ['BB_lower', 'BB_upper', 'close']

    # Check for required columns before accessing
    has_ma = all(col in df.columns for col in required_ma_cols)
    has_rsi = all(col in df.columns for col in required_rsi_cols)
    has_bb = all(col in df.columns for col in required_bb_cols)

    # 1. Moving Average Crossover
    if has_ma:
        if pd.notna(latest['SMA_short']) and pd.notna(latest['SMA_long']) and \
                pd.notna(previous['SMA_short']) and pd.notna(previous['SMA_long']):
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
        else:
            reasons.append("MA data contains NaN")
    else:
        reasons.append("MA data columns missing")

    # 2. RSI Analysis
    rsi = 50.0 # Default neutral
    if has_rsi and pd.notna(latest['RSI']):
        rsi = latest['RSI']
        if rsi < 30:
            score += 1
            reasons.append(f"Oversold conditions (RSI: {rsi:.1f})")
        elif rsi > 70:
            score -= 1
            reasons.append(f"Overbought conditions (RSI: {rsi:.1f})")
        else:
            reasons.append(f"Neutral momentum (RSI: {rsi:.1f})")
    else:
        reasons.append("RSI data missing or NaN")

    # 3. Price vs Bollinger Bands
    price = latest.get('close', None)
    if has_bb and price is not None and pd.notna(price) and \
            pd.notna(latest['BB_lower']) and pd.notna(latest['BB_upper']):
        if price < latest['BB_lower']:
            score += 1
            reasons.append("Price below lower Bollinger Band (potential reversal)")
        elif price > latest['BB_upper']:
            score -= 1
            reasons.append("Price above upper Bollinger Band (potential correction)")
    else:
        reasons.append("BB data or price missing/NaN")

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

    final_price = float(price) if price is not None and pd.notna(price) else 0.0
    final_rsi = float(rsi) if pd.notna(rsi) else 50.0

    return {
        'signal': signal, 'confidence': confidence,
        'reasoning': ' | '.join(reasons) if reasons else "No clear signals",
        'score': score, 'rsi': final_rsi, 'price': final_price
    }

def _get_chart_base64(chart_path: str) -> str:
    """Reads a chart image file and returns it as a Base64 encoded string."""
    if not chart_path or not os.path.exists(chart_path):
        logging.warning(f"Chart file not found or path is invalid for Base64 encoding: {chart_path}")
        return ""
    try:
        with open(chart_path, "rb") as image_file:
            encoded_string = base64.b64encode(image_file.read()).decode('utf-8')
        return f"data:image/png;base64,{encoded_string}"
    except Exception as e:
        logging.error(f"Could not encode chart '{chart_path}' to Base64: {e}")
        return ""

def _get_report_html_slides(signal_data, symbol, df: pd.DataFrame | None, chart_path, chart_base64_src):
    """
    Helper function to generate HTML content bodies for reports.
    Requires a df only for trend calculation.
    """
    css_content="""
:root{--color-primary:#1C2833;--color-secondary:#2E4053;--color-accent:#E74C3C;--color-success:#2ECC71;--color-warning:#F39C12;--color-bg:#F4F6F6;--color-text:#2C3E50;--font-main:Arial,sans-serif}
body{font-family:var(--font-main);color:var(--color-text);background-color:var(--color-bg);margin:0;padding:20px 0}
h1{color:var(--color-primary);font-size:48px;font-weight:700}h2{color:var(--color-secondary);font-size:32px;font-weight:700}h3{color:var(--color-secondary);font-size:24px;font-weight:700}
p{font-size:18px;line-height:1.5}ul{font-size:18px;line-height:2}.big-signal{font-size:72px;font-weight:700}.confidence{font-size:24px;font-style:italic}
.slide-container{width:960px;height:540px;margin:20px auto;padding:40px;box-sizing:border-box;background-color:#fff;box-shadow:0 4px 12px rgba(0,0,0,.1);border:1px solid #ddd;page-break-after:always;display:flex;flex-direction:column;overflow:hidden}
.row{display:flex;flex-direction:row}.col{display:flex;flex-direction:column}
@media print{body{background-color:#fff;padding:0}.slide-container{margin:0;box-shadow:none;border:none;height:100vh;padding-top:60px}}
"""
    signal_color = {'BUY':'#2ECC71','SELL':'#E74C3C','HOLD':'#F39C12'}.get(signal_data.get('signal','HOLD'),'#808080')
    signal=signal_data.get('signal','N/A')
    confidence=signal_data.get('confidence','N/A')
    reasoning=signal_data.get('reasoning','N/A')
    price=signal_data.get('price',0.0)
    rsi=signal_data.get('rsi',50.0)

    slide1_body=f"""<body style="width:960px;height:540px;margin:0;padding:0;background:linear-gradient(135deg,#1C2833 0%,#2E4053 100%);"><div class="col" style="height:100%;justify-content:center;align-items:center;padding:60px;"><h1 style="color:white;text-align:center;margin-bottom:40px;">Crypto Market Analysis Report</h1><h2 style="color:#AAB7B8;text-align:center;margin-bottom:60px;">{symbol}</h2><div style="text-align:center;"><p class="big-signal" style="color:{signal_color};">{signal}</p><p class="confidence" style="color:#BDC3C7;">Confidence: {confidence}</p></div></div></body>"""
    slide2_body=f"""<body style="width:960px;height:540px;margin:0;padding:0;background:var(--color-bg);"><div class="col" style="height:100%;padding:40px;"><h2 style="margin-bottom:30px;">Recommendation Details</h2><div style="background:white;padding:30px;border-radius:10px;box-shadow:0 4px 6px rgba(0,0,0,0.1);"><div class="row" style="align-items:center;margin-bottom:20px;"><h1 style="color:{signal_color};margin:0;">{signal}</h1><p style="margin-left:20px;font-size:20px;">({confidence} Confidence)</p></div><h3 style="margin-bottom:15px;">Analysis:</h3><p style="font-size:18px;line-height:1.6;">{reasoning}</p><div class="row" style="margin-top:30px;gap:40px;"><div><p style="font-size:16px;color:#7F8C8D;margin-bottom:5px;">Current Price</p><p style="font-size:24px;font-weight:bold;">${price:,.2f}</p></div><div><p style="font-size:16px;color:#7F8C8D;margin-bottom:5px;">RSI</p><p style="font-size:24px;font-weight:bold;">{rsi:.1f}</p></div></div></div></div></body>"""
    slide3_body_html="""<body style="width:960px;height:540px;margin:0;padding:0;background:var(--color-bg);"><div class="col" style="height:100%;padding:40px;"><h2 style="margin-bottom:30px;">Exchange Comparison</h2><div class="placeholder" style="flex:1;"></div></div></body>"""
    chart_img_tag_html=f'<img src="{chart_base64_src}" style="max-width:100%;max-height:100%;object-fit:contain;" alt="Technical Analysis Chart"/>' if chart_base64_src else '<p>Chart not available.</p>'
    slide4_body_html=f"""<body style="width:960px;height:540px;margin:0;padding:0;background:var(--color-bg);"><div class="col" style="height:100%;padding:40px;"><h2 style="margin-bottom:20px;">Technical Analysis Chart</h2><div style="flex:1;display:flex;justify-content:center;align-items:center;overflow:hidden;">{chart_img_tag_html}</div></div></body>"""
    trend_text="N/A"
    if df is not None and not df.empty and 'SMA_short' in df.columns and 'SMA_long' in df.columns and len(df)>=1:
        last=df.iloc[-1]
        if pd.notna(last['SMA_short']) and pd.notna(last['SMA_long']): trend_text="Bullish trend" if last['SMA_short']>last['SMA_long'] else "Bearish trend"
    rsi_text="Oversold" if rsi<30 else "Overbought" if rsi>70 else "Neutral"
    risk_text=("Lower risk" if signal=='BUY' else "Higher risk" if signal=='SELL' else "Neutral risk")
    slide5_body=f"""<body style="width:960px;height:540px;margin:0;padding:0;background:var(--color-bg);"><div class="col" style="height:100%;padding:40px;"><h2 style="margin-bottom:30px;">What This Means For You</h2><div style="background:white;padding:30px;border-radius:10px;"><h3>Key Takeaways:</h3><ul><li><strong>Signal:</strong> {signal} recommendation</li><li><strong>Moving Averages:</strong> {trend_text} detected</li><li><strong>Momentum:</strong> RSI at {rsi:.1f} ({rsi_text})</li><li><strong>Risk Level:</strong> {risk_text}</li></ul></div></div></body>"""
    slide6_body="""<body style="width:960px;height:540px;margin:0;padding:0;background:var(--color-bg);"><div class="col" style="height:100%;padding:40px;"><h2 style="margin-bottom:30px;">Learn: Moving Averages</h2><div style="background:white;padding:30px;border-radius:10px;"><h3>What Are MAs?</h3><p style="margin-bottom:20px;">Smooth price data to identify trends.</p><h3>Key Signals:</h3><ul><li><strong>Golden Cross:</strong> Short MA > Long MA (bullish)</li><li><strong>Death Cross:</strong> Short MA < Long MA (bearish)</li><li>Price > MA: Uptrend</li><li>Price < MA: Downtrend</li></ul></div></div></body>"""
    slide7_body="""<body style="width:960px;height:540px;margin:0;padding:0;background:var(--color-bg);"><div class="col" style="height:100%;padding:40px;"><h2 style="margin-bottom:30px;">Learn: RSI</h2><div style="background:white;padding:30px;border-radius:10px;"><h3>What Is RSI?</h3><p style="margin-bottom:20px;">Measures momentum (0-100).</p><h3>How to Read:</h3><ul><li><strong>> 70:</strong> Overbought (pullback likely)</li><li><strong>< 30:</strong> Oversold (bounce likely)</li><li>30-70: Neutral</li><li>Divergence: Price/RSI opposite direction</li></ul></div></div></body>"""
    slide8_body_html="""<body style="width:960px;height:540px;margin:0;padding:0;background:var(--color-bg);"><div class="col" style="height:100%;padding:40px;"><h2 style="margin-bottom:30px;color:var(--color-accent);">Important Disclaimer</h2><div style="background:white;padding:30px;border-radius:10px;"><p style="margin-bottom:20px;font-weight:bold;">Educational/Informational ONLY. NOT financial advice.</p><ul style="line-height:1.6;"><li>Crypto trading has substantial risk.</li><li>Always DYOR.</li><li>Never invest more than you can lose.</li><li>Past performance != future results.</li><li>Consult a qualified advisor.</li></ul></div></div></body>"""

    # --- Content for PPTX temporary files ---
    slide1_html_file=f"<!DOCTYPE html><html><head><style>{css_content}</style></head>{slide1_body}</html>"
    slide2_html_file=f"<!DOCTYPE html><html><head><style>{css_content}</style></head>{slide2_body}</html>"
    slide3_body_pptx="""<body style="width:960px;height:540px;margin:0;padding:0;background:var(--color-bg);"><div class="col" style="height:100%;padding:40px;"><h2 style="margin-bottom:30px;">Exchange Comparison</h2><div class="placeholder" style="flex:1;"></div></div></body>"""
    slide3_html_file=f"<!DOCTYPE html><html><head><style>{css_content}</style></head>{slide3_body_pptx}</html>"
    chart_filename_for_pptx=os.path.basename(chart_path) if chart_path and os.path.exists(chart_path) else ""
    chart_img_tag_pptx=f'<img src="{chart_filename_for_pptx}" style="max-width:100%;max-height:100%;object-fit:contain;" alt="Technical Analysis Chart"/>' if chart_filename_for_pptx else '<p>Chart not available.</p>'
    slide4_body_pptx=f"""<body style="width:960px;height:540px;margin:0;padding:0;background:var(--color-bg);"><div class="col" style="height:100%;padding:40px;"><h2 style="margin-bottom:20px;">Technical Analysis Chart</h2><div style="flex:1;display:flex;justify-content:center;align-items:center;overflow:hidden;">{chart_img_tag_pptx}</div></div></body>"""
    slide4_html_file=f"<!DOCTYPE html><html><head><style>{css_content}</style></head>{slide4_body_pptx}</html>"
    slide5_html_file=f"<!DOCTYPE html><html><head><style>{css_content}</style></head>{slide5_body}</html>"
    slide6_html_file=f"<!DOCTYPE html><html><head><style>{css_content}</style></head>{slide6_body}</html>"
    slide7_html_file=f"<!DOCTYPE html><html><head><style>{css_content}</style></head>{slide7_body}</html>"

    return {
        "css":css_content,
        "slides_html_files":[("slide1.html",slide1_html_file),("slide2.html",slide2_html_file),("slide3.html",slide3_html_file),("slide4.html",slide4_html_file),("slide5.html",slide5_html_file),("slide6.html",slide6_html_file),("slide7.html",slide7_html_file)],
        "slides_for_html_report":[slide1_body,slide2_body,slide3_body_html,slide4_body_html,slide5_body,slide6_body,slide7_body,slide8_body_html]
    }


def generate_powerpoint_report(
        exchange_results: pd.DataFrame,
        symbol: str,
        chart_path: str,
        output_dir: str, # This IS the reports subfolder path
        html_slides: dict,
        date_prefix: str
):
    """Generates a PPTX report in output_dir (reports subfolder) with date_prefix."""
    logging.info(f"Generating a PowerPoint report in {output_dir}...")
    temp_files_created = []
    js_script_path = os.path.join(os.getcwd(), 'generate_pptx.js') # Temp script in CWD

    try:
        # Verify/Install html2pptx
        logging.info("Assuming html2pptx dependency is met.")

        # Write temporary HTML slide files needed by the JS script to CWD
        for filename, content in html_slides["slides_html_files"]:
            filepath = os.path.join(os.getcwd(), filename)
            with open(filepath, 'w', encoding='utf-8') as f:
                f.write(content)
            temp_files_created.append(filepath)

        # Create JS code
        pptx_filename = f"{date_prefix}Crypto_Market_Analysis.pptx"
        pptx_path_abs = os.path.abspath(os.path.join(output_dir, pptx_filename)) # Final path in reports subfolder
        js_pptx_path = json.dumps(pptx_path_abs)
        js_symbol = json.dumps(symbol)
        js_exchanges = json.dumps(exchange_results.to_dict('records') if isinstance(exchange_results, pd.DataFrame) else [])

        #
        # --- FIXED: Replaced f-string with string formatting ---
        #
        js_code_template = """
const pptxgen = require("pptxgenjs");
const {{ html2pptx }} = require("@ant/html2pptx");
const fs = require('fs');
(async () => {{
    try {{
        const pptx = new pptxgen(); pptx.layout="LAYOUT_16x9"; pptx.author="Crypto Trading Bot"; pptx.title="Market Analysis Report - " + {js_symbol};
        await html2pptx("slide1.html", pptx); await html2pptx("slide2.html", pptx);
        const {{ slide: slide3, placeholders: p3 }} = await html2pptx("slide3.html", pptx);
        const tableData = [[{{'text':'Exchange','options':{{'fill':{{'color':'2E4053'}},'color':'FFFFFF','bold':true}}}},{{'text':'Total Pairs','options':{{'fill':{{'color':'2E4053'}},'color':'FFFFFF','bold':true}}}},{{'text':'USDT Pairs','options':{{'fill':{{'color':'2E4053'}},'color':'FFFFFF','bold':true}}}},{{'text':'Has OHLCV','options':{{'fill':{{'color':'2E4053'}},'color':'FFFFFF','bold':true}}}}]];
        const exchanges = {js_exchanges}; exchanges.forEach(ex=>{{tableData.push([ex.name||'N/A',ex.total_spot_pairs!=null?ex.total_spot_pairs.toString():'N/A',ex.usdt_quoted_pairs!=null?ex.usdt_quoted_pairs.toString():'N/A',ex.supports_fetchOHLCV?'Yes':'No']);}});
        if(p3&&p3.length>0){{slide3.addTable(tableData,{{...p3[0],border:{{pt:1,color:'CCCCCC'}},fontSize:16,align:'center',valign:'middle'}});}}else{{console.log("Warn: Placeholder missing slide 3.");slide3.addTable(tableData,{{x:0.5,y:1.5,w:9.0,border:{{pt:1,color:'CCCCCC'}},fontSize:16,align:'center',valign:'middle'}});}}
        await html2pptx("slide4.html", pptx); await html2pptx("slide5.html", pptx); await html2pptx("slide6.html", pptx); await html2pptx("slide7.html", pptx);
        await pptx.writeFile({js_pptx_path}); console.log("PowerPoint generated: " + {js_pptx_path});
    }} catch(err){{console.error("Error generating PowerPoint:",err.message,err.stack);process.exit(1);}}
}})();
"""
        js_code = js_code_template.format(
            js_symbol=js_symbol,
            js_exchanges=js_exchanges,
            js_pptx_path=js_pptx_path
        )
        #
        # --- END FIX ---
        #

        with open(js_script_path, 'w', encoding='utf-8') as f:
            f.write(js_code)
        temp_files_created.append(js_script_path)

        # Copy chart to CWD if it exists
        chart_filename = os.path.basename(chart_path) if chart_path else None
        if chart_filename and os.path.exists(chart_path):
            import shutil
            dest_path = os.path.join(os.getcwd(), chart_filename)
            if os.path.abspath(chart_path) != dest_path:
                shutil.copy(chart_path, dest_path)
                logging.info(f"Copied chart to CWD for PPTX: {dest_path}")
                temp_files_created.append(dest_path)
        elif chart_path:
            logging.warning(f"Chart file {chart_path} not found for PPTX.")
        else:
            logging.warning("No chart path provided for PPTX.")

        # Execute Node.js script
        logging.info(f"Executing Node.js script: {js_script_path}")
        node_path_cmd = 'NODE_PATH="$(npm root -g)"'
        node_cmd = f'node {os.path.basename(js_script_path)} 2>&1'
        full_command = f'{node_path_cmd} {node_cmd}'
        result = subprocess.run(full_command, shell=True, capture_output=True, text=True, check=True, cwd=os.getcwd(), timeout=60)
        logging.info(f"PPTX generation script executed: {result.stdout}")

    except subprocess.TimeoutExpired:
        logging.error("Timeout during PPTX Node script.")
    except subprocess.CalledProcessError as e:
        logging.error(f"PPTX script failed (code {e.returncode}): {e.stderr}")
    except Exception as e:
        logging.error(f"Error during PPTX generation: {e}")
    finally: # Unified cleanup
        logging.debug(f"Cleaning up PPTX temp files: {temp_files_created}")
        for fpath in temp_files_created:
            try:
                if os.path.exists(fpath):
                    os.remove(fpath)
            except Exception as e:
                logging.warning(f"Failed cleanup: {fpath} ({e})")


def generate_word_report(
        signal_data: dict,
        exchange_results: pd.DataFrame,
        symbol: str,
        suggestions: list,
        output_dir: str, # This IS the reports subfolder path
        date_prefix: str
):
    """Generates a DOCX report in output_dir (reports subfolder) with date_prefix."""
    logging.info(f"Generating a Word document report in {output_dir}...")
    js_script_path = os.path.join(os.getcwd(), 'generate_docx.js') # Temp script in CWD
    temp_files_created = [js_script_path] # Keep track for cleanup

    try:
        # Safely embed data using json.dumps()
        js_signal=json.dumps(signal_data.get('signal','HOLD'))
        js_confidence=json.dumps(signal_data.get('confidence','LOW'))
        js_symbol=json.dumps(symbol)
        js_reasoning=json.dumps(signal_data.get('reasoning','N/A'))
        js_price=json.dumps(f"${signal_data.get('price',0.0):,.2f}")
        js_rsi=json.dumps(f"{signal_data.get('rsi',50.0):.1f}")
        js_exchanges_data=json.dumps(exchange_results.to_dict('records') if isinstance(exchange_results,pd.DataFrame) else [])
        js_suggestions=json.dumps(suggestions if suggestions else [])

        docx_filename=f"{date_prefix}Crypto_Market_Analysis.docx"
        docx_path_abs=os.path.abspath(os.path.join(output_dir, docx_filename)) # Final path in reports subfolder
        js_docx_path=json.dumps(docx_path_abs)

        #
        # --- FIXED: Replaced f-string with string formatting ---
        #
        js_code_template = """
const {{ Document, Packer, Paragraph, TextRun, Table, TableRow, TableCell, AlignmentType, HeadingLevel, BorderStyle, WidthType, ShadingType, VerticalAlign, LevelFormat, PageBreak }} = require('docx');
const fs = require('fs');
const signal={js_signal}; const confidence={js_confidence}; const symbol={js_symbol}; const reasoning={js_reasoning}; const currentPrice={js_price}; const currentRsi={js_rsi}; const exchangeData={js_exchanges_data}||[]; const analysisSuggestions={js_suggestions}||[];
const tableBorder={{style:BorderStyle.SINGLE,size:1,color:"CCCCCC"}}; const docxFilePath={js_docx_path};
try {{ // Wrap doc creation
    const docChildren=[ /* Static + Placeholders */
        new Paragraph({{heading:HeadingLevel.TITLE,alignment:AlignmentType.CENTER,children:[new TextRun("Cryptocurrency Market Analysis Report")]}}),new Paragraph({{alignment:AlignmentType.CENTER,spacing:{{after:400}},children:[new TextRun({{text:symbol||'N/A',size:32,color:"7F8C8D"}})]}}),
        new Paragraph({{heading:HeadingLevel.HEADING_1,children:[new TextRun("Executive Summary")]}}),new Paragraph({{spacing:{{after:200}},children:[new TextRun("Trading Recommendation: "),new TextRun({{text:signal||'HOLD',bold:true,size:32,color:signal==='BUY'?'2ECC71':signal==='SELL'?'E74C3C':'F39C12'}})]}}),new Paragraph({{spacing:{{after:200}},children:[new TextRun("Confidence Level: "),new TextRun({{text:confidence||'LOW',bold:true,size:24}})]}}),new Paragraph({{spacing:{{after:300}},children:[new TextRun({{text:reasoning||'N/A',italics:true}})]}}),new Paragraph({{spacing:{{after:200}},children:[new TextRun("Current Price: "),new TextRun({{text:currentPrice||'$0.00',bold:true,size:28}})]}}),new Paragraph({{spacing:{{after:400}},children:[new TextRun("RSI: "),new TextRun({{text:currentRsi||'50.0',bold:true,size:28}})]}}),
        new Paragraph({{children:[new PageBreak()]}}),
        new Paragraph({{heading:HeadingLevel.HEADING_1,children:[new TextRun("Exchange Comparison")]}}),new Paragraph({{spacing:{{after:200}},children:[new TextRun("Below is a comparison...")]}}),
        new Table({{columnWidths:[2340,2340,2340,2340],margins:{{top:100,bottom:100,left:180,right:180}},rows:[new TableRow({{tableHeader:true,children:[new TableCell({{borders:tableBorder,width:{{size:2340,type:WidthType.DXA}},shading:{{fill:"2E4053"}},verticalAlign:VerticalAlign.CENTER,children:[new Paragraph({{alignment:AlignmentType.CENTER,children:[new TextRun({{text:"Exchange",bold:true,color:"FFFFFF",size:22}})]}})]}}),new TableCell({{borders:tableBorder,width:{{size:2340,type:WidthType.DXA}},shading:{{fill:"2E4053"}},children:[new Paragraph({{alignment:AlignmentType.CENTER,children:[new TextRun({{text:"Total Pairs",bold:true,color:"FFFFFF",size:22}})]}})]}}),new TableCell({{borders:tableBorder,width:{{size:2340,type:WidthType.DXA}},shading:{{fill:"2E4053"}},children:[new Paragraph({{alignment:AlignmentType.CENTER,children:[new TextRun({{text:"USDT Pairs",bold:true,color:"FFFFFF",size:22}})]}})]}}),new TableCell({{borders:tableBorder,width:{{size:2340,type:WidthType.DXA}},shading:{{fill:"2E4053"}},children:[new Paragraph({{alignment:AlignmentType.CENTER,children:[new TextRun({{text:"Has OHLCV",bold:true,color:"FFFFFF",size:22}})]}})]}})]}})]}}), // Rows added later
        new Paragraph({{children:[new PageBreak()]}}),
        new Paragraph({{heading:HeadingLevel.HEADING_1,children:[new TextRun("Technical Analysis Details")]}}),
        // Suggestions inserted here
        new Paragraph({{children:[new PageBreak()]}}),
        // Static sections (Educational, etc.)
        new Paragraph({{heading:HeadingLevel.HEADING_1,children:[new TextRun("Educational Guide...")]}}),
        // ... rest of static content ...
        new Paragraph({{heading:HeadingLevel.HEADING_2,children:[new TextRun("Important Disclaimers")]}}), new Paragraph({{numbering:{{reference:"bullet-list",level:0}},children:[new TextRun("This report is for educational purposes only...")]}})
    ];
    // --- Dynamic Insertion ---
    const exchangeTable=docChildren.find(c=>c instanceof Table); if(exchangeTable&&exchangeTable.root&&Array.isArray(exchangeTable.root)){{exchangeData.forEach(ex=>{{exchangeTable.root.push(new TableRow({{children:[new TableCell({{borders:tableBorder,width:{{size:2340,type:WidthType.DXA}},children:[new Paragraph({{children:[new TextRun(String(ex.name||'N/A'))]})]}}),new TableCell({{borders:tableBorder,width:{{size:2340,type:WidthType.DXA}},children:[new Paragraph({{alignment:AlignmentType.CENTER,children:[new TextRun(String(ex.total_spot_pairs==null?'N/A':ex.total_spot_pairs))]})]}}),new TableCell({{borders:tableBorder,width:{{size:2340,type:WidthType.DXA}},children:[new Paragraph({{alignment:AlignmentType.CENTER,children:[new TextRun(String(ex.usdt_quoted_pairs==null?'N/A':ex.usdt_quoted_pairs))]})]}}),new TableCell({{borders:tableBorder,width:{{size:2340,type:WidthType.DXA}},children:[new Paragraph({{alignment:AlignmentType.CENTER,children:[new TextRun(ex.supports_fetchOHLCV?"Yes":"No")]})]})]})]}}));}});}}else{{console.error("Could not find exchange table.");}}
    let analysisIdx=-1;for(let i=0;i<docChildren.length;i++){{if(docChildren[i]instanceof Paragraph&&docChildren[i].options.heading===HeadingLevel.HEADING_1){{let txt='';if(docChildren[i].root&&Array.isArray(docChildren[i].root)){{docChildren[i].root.forEach(r=>{{if(r instanceof TextRun&&r.root&&Array.isArray(r.root)){{r.root.forEach(t=>{{if(typeof t==='object'&&t!==null&&typeof t.text==='string'){{txt+=t.text;}}}});}}}});}}if(txt==='Technical Analysis Details'){{analysisIdx=i;break;}}}}}}
    if(analysisIdx!==-1){{const suggestionParas=analysisSuggestions.map(s=>new Paragraph({{spacing:{{after:150}},numbering:{{reference:"bullet-list",level:0}},children:[new TextRun(s||"")]}}));docChildren.splice(analysisIdx+1,0,...suggestionParas);}}else{{console.error("Could not find 'Technical Analysis Details' heading.");}}
    const finalDoc=new Document({{
        styles:{{default:{{document:{{run:{{font:"Arial",size:24}}}}}},paragraphStyles:[{{id:"Title",name:"Title",basedOn:"Normal",run:{{size:56,bold:true,color:"1C2833",font:"Arial"}},paragraph:{{spacing:{{before:240,after:120}},alignment:AlignmentType.CENTER}}}},{{id:"Heading1",name:"Heading 1",basedOn:"Normal",quickFormat:true,run:{{size:32,bold:true,color:"2E4053",font:"Arial"}},paragraph:{{spacing:{{before:240,after:180}},outlineLevel:0}}}},{{id:"Heading2",name:"Heading 2",basedOn:"Normal",quickFormat:true,run:{{size:28,bold:true,color:"2E4053",font:"Arial"}},paragraph:{{spacing:{{before:180,after:120}},outlineLevel:1}}}},{{id:"Heading3",name:"Heading 3",basedOn:"Normal",quickFormat:true,run:{{size:24,bold:true,color:"34495E",font:"Arial"}},paragraph:{{spacing:{{before:120,after:80}},outlineLevel:2}}}}]}},
        numbering:{{config:[{{reference:"bullet-list",levels:[{{level:0,format:LevelFormat.BULLET,text:"•",alignment:AlignmentType.LEFT,style:{{paragraph:{{indent:{{left:720,hanging:360}}}}}}}}]}},{{reference:"numbered-list",levels:[{{level:0,format:LevelFormat.DECIMAL,text:"%1.",alignment:AlignmentType.LEFT,style:{{paragraph:{{indent:{{left:720,hanging:360}}}}}}}}]}}]}},
        sections:[{{properties:{{page:{{margin:{{top:1440,right:1440,bottom:1440,left:1440}}}}}},children:docChildren}}]
    }});
    Packer.toBuffer(finalDoc).then(buffer=>{{fs.writeFileSync(docxFilePath,buffer);console.log("Word document generated: "+docxFilePath);}}).catch(err=>{{console.error("Error generating/writing Word buffer:",err.message,err.stack);process.exit(1);}});
}} catch(err){{console.error("Error creating Word document object:",err.message,err.stack);process.exit(1);}}
"""
        js_code = js_code_template.format(
            js_signal=js_signal,
            js_confidence=js_confidence,
            js_symbol=js_symbol,
            js_reasoning=js_reasoning,
            js_price=js_price,
            js_rsi=js_rsi,
            js_exchanges_data=js_exchanges_data,
            js_suggestions=js_suggestions,
            js_docx_path=js_docx_path
        )
        #
        # --- END FIX ---
        #

        with open(js_script_path, 'w', encoding='utf-8') as f:
            f.write(js_code)

        # Run Node.js script
        logging.info(f"Executing Node.js script: {js_script_path}")
        node_cmd = f'node {os.path.basename(js_script_path)} 2>&1'
        full_command = node_cmd
        result = subprocess.run(full_command, shell=True, capture_output=True, text=True, check=True, cwd=os.getcwd(), timeout=60)
        logging.info(f"Word document report script executed: {result.stdout}")

    except subprocess.TimeoutExpired:
        logging.error("Timeout during Word doc Node script.")
    except subprocess.CalledProcessError as e:
        logging.error(f"Word doc script failed (code {e.returncode}): {e.stderr}")
    except Exception as e:
        logging.error(f"Unexpected error during DOCX generation: {e}")
    finally: # Cleanup JS file
        for fpath in temp_files_created:
            try:
                if os.path.exists(fpath):
                    os.remove(fpath)
            except OSError as e:
                logging.warning(f"Could not clean up temp file {fpath}: {e}")


def generate_single_page_html_report(
        exchange_results: pd.DataFrame,
        symbol: str,
        output_dir: str,        # Configurable directory for the final HTML (reports subfolder)
        html_slides: dict,      # Dict containing pre-rendered HTML slide bodies
        date_prefix: str        # Date prefix for filename
):
    """
    Generates a single-page HTML report with an embedded chart.
    Saves the final HTML to the specified output_dir with date prefix.
    """
    logging.info(f"Generating a single-page HTML report in {output_dir}...")

    # Construct filename with date prefix
    html_filename = f"{date_prefix}Crypto_Market_Analysis.html"
    # Ensure final path uses the passed output_dir (reports subfolder)
    html_path = os.path.join(output_dir, html_filename)

    # Generate table HTML for exchange comparison
    exchange_table_html = """<table style="width:100%;border-collapse:collapse;margin-top:20px;font-size:16px;"><thead><tr style="background-color:var(--color-secondary);color:white;"><th style="padding:10px;border:1px solid #ccc;text-align:left;">Exchange</th><th style="padding:10px;border:1px solid #ccc;">Total Pairs</th><th style="padding:10px;border:1px solid #ccc;">USDT Pairs</th><th style="padding:10px;border:1px solid #ccc;">Has OHLCV</th></tr></thead><tbody>"""
    if isinstance(exchange_results, pd.DataFrame) and not exchange_results.empty:
        for _, row in exchange_results.iterrows():
            name=str(row.get('name','N/A'))
            total_pairs=str(row.get('total_spot_pairs','N/A'))
            usdt_pairs=str(row.get('usdt_quoted_pairs','N/A'))
            has_ohlcv=row.get('supports_fetchOHLCV',False)
            ohlcv_text='✓' if has_ohlcv else '✗'
            ohlcv_color='var(--color-success)' if has_ohlcv else 'var(--color-accent)'
            exchange_table_html += f"""<tr style="background-color:white;border-bottom:1px solid #eee;"><td style="padding:10px;border:1px solid #ccc;">{name}</td><td style="padding:10px;border:1px solid #ccc;text-align:center;">{total_pairs}</td><td style="padding:10px;border:1px solid #ccc;text-align:center;">{usdt_pairs}</td><td style="padding:10px;border:1px solid #ccc;text-align:center;color:{ohlcv_color};font-weight:bold;">{ohlcv_text}</td></tr>"""
    else:
        logging.warning("Exchange results data missing/empty for HTML table.")
        exchange_table_html += '<tr><td colspan="4" style="padding:10px;border:1px solid #ccc;text-align:center;">No exchange data.</td></tr>'
    exchange_table_html += "</tbody></table>"

    # Replace placeholder in slide 3 body content for the HTML report
    placeholder_tag = '<div class="placeholder" style="flex: 1;"></div>'
    slides_list = html_slides.get("slides_for_html_report", [])
    if len(slides_list) > 2:
        if placeholder_tag in slides_list[2]:
            slides_list[2] = slides_list[2].replace(placeholder_tag, f'<div style="flex:1;overflow-y:auto;">{exchange_table_html}</div>')
        else:
            logging.warning("Placeholder missing in HTML slide 3.")
    else:
        logging.error("HTML slides array missing slide 3.")

    # Combine all parts into a single HTML string
    full_html = f"""<!DOCTYPE html><html lang="en"><head><meta charset="UTF-8"><meta name="viewport" content="width=device-width, initial-scale=1.0"><title>Crypto Market Analysis - {symbol}</title><style>{html_slides.get('css','')}</style></head><body>"""

    #
    # --- THIS IS THE FIX for the typo ---
    #
    if slides_list:
        for idx, slide_body in enumerate(slides_list): # Changed 'slide_.body' back to 'slide_body'
            try:
                start=slide_body.index('>')+1
                end=slide_body.rindex('</body')
                content=slide_body[start:end].strip()
                full_html+=f'<div class="slide-container">\n{content}\n</div>\n'
            except (ValueError, IndexError, AttributeError, TypeError) as e:
                logging.warning(f"Error extracting HTML slide {idx}: {e}")
                full_html+=f'<div class="slide-container"><p>Error slide {idx}.</p></div>\n'
    else:
        logging.error("Cannot generate HTML body.")
        full_html+='<p>Error: Report content missing.</p>'
    full_html += "</body>\n</html>"
    #
    # --- END FIX ---
    #

    # Write the combined HTML to file
    try:
        os.makedirs(output_dir, exist_ok=True) # Ensure dir exists
        with open(html_path, 'w', encoding='utf-8') as f:
            f.write(full_html)
        logging.info(f"Single-page HTML report generated: {html_path}")
    except Exception as e:
        logging.error(f"Failed to write HTML report to {html_path}: {e}")


def generate_comprehensive_reports(
        exchange_results: pd.DataFrame,
        symbol: str,
        df: pd.DataFrame | None, # Needed for signal determination & trend text
        suggestions: list,
        chart_path: str, # A path to a temporary chart image (can be None)
        output_dir: str, # Configurable path to the REPORTS SUBFOLDER
        date_prefix: str # Date prefix for filenames
):
    """
    Main function to generate all reports: PowerPoint, Word, and HTML.
    Ensures the output directory exists and cleans up temporary files.
    Accepts the specific reports subfolder path as output_dir.
    Returns the signal_data dict.
    """
    # Ensure the output directory (reports subfolder) exists
    try:
        os.makedirs(output_dir, exist_ok=True)
    except Exception as e:
        logging.error(f"Could not create reports output directory '{output_dir}': {e}")
        return {'signal':'ERROR','confidence':'NONE','reasoning':f'Failed directory creation: {e}','score':0,'rsi':0,'price':0}

    # Determine trading signal (handles None/empty df)
    signal_data = determine_trading_signal(df)

    logging.info(f"\n--- Report Signal ---\nSignal: {signal_data.get('signal','N/A')} (Confidence: {signal_data.get('confidence','N/A')})\nReasoning: {signal_data.get('reasoning','N/A')}\n---------------------\n")

    # Get Base64 chart (handles None chart_path)
    chart_base64_src = _get_chart_base64(chart_path)
    if not chart_base64_src and chart_path and os.path.exists(chart_path):
        logging.warning("Base64 chart encoding failed.")

    # Get HTML content (handles None df)
    html_slides = _get_report_html_slides(signal_data, symbol, df, chart_path, chart_base64_src)

    # Generate all three reports
    generate_powerpoint_report(exchange_results, symbol, chart_path, output_dir, html_slides, date_prefix)
    generate_word_report(signal_data, exchange_results, symbol, suggestions, output_dir, date_prefix)
    generate_single_page_html_report(exchange_results, symbol, output_dir, html_slides, date_prefix)

    # --- Unified Cleanup ---
    logging.info("Cleaning up temporary files...")
    temp_files = [f[0] for f in html_slides.get("slides_html_files",[])]
    temp_files.extend(["generate_pptx.js", "generate_docx.js"])
    chart_filename=os.path.basename(chart_path) if chart_path else None
    if chart_filename:
        chart_copy_path_cwd=os.path.join(os.getcwd(),chart_filename)
        temp_files.append(chart_copy_path_cwd) # Add potential chart copy

    removed_count=0
    failed_count=0
    for filename in temp_files:
        filepath=os.path.abspath(os.path.join(os.getcwd(),os.path.basename(filename))) # Absolute path in CWD
        if os.path.dirname(filepath)==os.getcwd(): # Safety check: Only remove if in CWD
            try:
                if os.path.exists(filepath):
                    os.remove(filepath)
                    logging.debug(f"Removed: {filepath}")
                    removed_count+=1
            except OSError as e:
                logging.warning(f"Cleanup failed: {filepath} ({e})")
                failed_count+=1
        # else: logging.warning(f"Skipping cleanup outside CWD: {filepath}") # Optional
    logging.info(f"Cleanup complete. Removed: {removed_count}, Failed: {failed_count}.")

    return signal_data # Return determined signal data