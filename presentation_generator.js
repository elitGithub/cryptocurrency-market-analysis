const fs = require('fs');
const pptxgen = require('pptxgenjs');
const { html2pptx } = require('@ant/html2pptx');

function generateBuySellHold(suggestions, latest_data) {
    let bullish_signals = 0;
    let bearish_signals = 0;
    
    suggestions.forEach(s => {
        if (s.includes('BULLISH') || s.includes('Golden Cross') || s.includes('oversold')) bullish_signals++;
        if (s.includes('BEARISH') || s.includes('Death Cross') || s.includes('overbought')) bearish_signals++;
    });
    
    const rsi = latest_data.RSI;
    
    if (rsi < 30 || bullish_signals > bearish_signals) {
        return {
            recommendation: 'BUY',
            color: '#27AE60',
            reasoning: `Strong bullish signals detected. RSI at ${rsi.toFixed(2)} suggests potential upside.`
        };
    } else if (rsi > 70 || bearish_signals > bullish_signals) {
        return {
            recommendation: 'SELL',
            color: '#E74C3C',
            reasoning: `Bearish signals present. RSI at ${rsi.toFixed(2)} indicates overbought conditions.`
        };
    } else {
        return {
            recommendation: 'HOLD',
            color: '#F39C12',
            reasoning: `Mixed signals. RSI at ${rsi.toFixed(2)} is neutral. Monitor and wait.`
        };
    }
}

async function generatePresentation(data, outputPath) {
    const { symbol, exchange_results, suggestions, latest_data, chart_path } = data;
    const signal = generateBuySellHold(suggestions, latest_data);
    
    // Create shared CSS
    const sharedCSS = `
        :root {
            --color-primary: #2C3E50;
            --color-secondary: #3498DB;
            --color-accent: #27AE60;
            --color-warning: #F39C12;
            --color-danger: #E74C3C;
            --color-light: #ECF0F1;
            --color-dark: #1C2833;
        }
        
        h1 {
            font-size: 48px;
            font-weight: 600;
            color: var(--color-primary);
            margin: 0;
        }
        
        h2 {
            font-size: 36px;
            font-weight: 600;
            color: var(--color-secondary);
            margin: 0;
        }
        
        h3 {
            font-size: 24px;
            font-weight: 600;
            color: var(--color-dark);
            margin: 0;
        }
        
        p {
            font-size: 18px;
            line-height: 1.4;
            color: var(--color-dark);
            margin: 0;
        }
        
        ul {
            font-size: 18px;
            color: var(--color-dark);
        }
        
        li {
            margin-bottom: 0.5rem;
        }
        
        .highlight {
            background: var(--color-light);
            padding: 1.5rem;
            border-radius: 8px;
        }
        
        .metric-box {
            background: white;
            border: 2px solid var(--color-secondary);
            padding: 1rem;
            border-radius: 8px;
            text-align: center;
        }
        
        .metric-value {
            font-size: 32px;
            font-weight: 600;
            color: var(--color-secondary);
        }
        
        .metric-label {
            font-size: 14px;
            color: #666;
            margin-top: 0.5rem;
        }
    `;
    
    // Slide 1: Title
    const slide1HTML = `
<!DOCTYPE html>
<html><head><style>${sharedCSS}</style></head>
<body class="col center" style="width:960px;height:540px;background:linear-gradient(135deg, #667eea 0%, #764ba2 100%);">
    <h1 style="color:white;font-size:60px;margin-bottom:1rem;">Cryptocurrency Market Analysis</h1>
    <h2 style="color:white;font-size:48px;">${symbol}</h2>
    <p style="color:white;font-size:20px;margin-top:2rem;opacity:0.9;">Comprehensive Technical Analysis & Trading Report</p>
</body></html>
    `;
    fs.writeFileSync('/home/claude/slide1.html', slide1HTML);
    
    // Slide 2: Executive Summary
    const slide2HTML = `
<!DOCTYPE html>
<html><head><style>${sharedCSS}</style></head>
<body class="col" style="width:960px;height:540px;padding:60px;background:white;">
    <h2 style="margin-bottom:2rem;">Executive Summary</h2>
    <div class="highlight" style="background:${signal.color};padding:2rem;border-radius:12px;">
        <h1 style="color:white;font-size:56px;text-align:center;">RECOMMENDATION: ${signal.recommendation}</h1>
    </div>
    <div class="row gap-lg" style="margin-top:2rem;">
        <div class="metric-box fill-width">
            <p class="metric-value">$${latest_data.close.toFixed(2)}</p>
            <p class="metric-label">Current Price</p>
        </div>
        <div class="metric-box fill-width">
            <p class="metric-value" style="color:${latest_data.RSI > 70 ? '#E74C3C' : latest_data.RSI < 30 ? '#27AE60' : '#3498DB'}">${latest_data.RSI.toFixed(0)}</p>
            <p class="metric-label">RSI (14)</p>
        </div>
        <div class="metric-box fill-width">
            <p class="metric-value" style="color:${latest_data.SMA_short > latest_data.SMA_long ? '#27AE60' : '#E74C3C'}">${latest_data.SMA_short > latest_data.SMA_long ? '↑' : '↓'}</p>
            <p class="metric-label">Trend</p>
        </div>
    </div>
    <p style="margin-top:2rem;font-size:18px;">${signal.reasoning}</p>
</body></html>
    `;
    fs.writeFileSync('/home/claude/slide2.html', slide2HTML);
    
    // Slide 3: Exchange Comparison
    const exchangeTable = exchange_results.map(ex => 
        `<tr><td>${ex.name}</td><td>${ex.total_spot_pairs}</td><td>${ex.usdt_quoted_pairs}</td><td>${ex.supports_fetchOHLCV ? '✓' : '✗'}</td></tr>`
    ).join('');
    
    const slide3HTML = `
<!DOCTYPE html>
<html><head><style>
${sharedCSS}
table { width:100%; border-collapse:collapse; margin-top:1.5rem; }
th { background:#3498DB; color:white; padding:0.75rem; text-align:left; font-size:18px; }
td { padding:0.75rem; border-bottom:1px solid #ddd; font-size:16px; }
tr:hover { background:#f5f5f5; }
</style></head>
<body class="col" style="width:960px;height:540px;padding:60px;background:white;">
    <h2 style="margin-bottom:1rem;">Exchange Comparison</h2>
    <table>
        <tr><th>Exchange</th><th>Spot Pairs</th><th>USDT Pairs</th><th>OHLCV Data</th></tr>
        ${exchangeTable}
    </table>
    <p style="margin-top:2rem;"><strong>Recommendation:</strong> ${exchange_results[0].name} offers the most comprehensive trading options with ${exchange_results[0].usdt_quoted_pairs} USDT pairs.</p>
</body></html>
    `;
    fs.writeFileSync('/home/claude/slide3.html', slide3HTML);
    
    // Slide 4: Technical Analysis
    const bullets = suggestions.map(s => {
        const type = s.includes('BULLISH') ? '🟢' : s.includes('BEARISH') ? '🔴' : s.includes('WARNING') ? '⚠️' : '📊';
        return `<li>${type} ${s.replace(/^[A-Z]+: /, '')}</li>`;
    }).join('');
    
    const slide4HTML = `
<!DOCTYPE html>
<html><head><style>${sharedCSS}</style></head>
<body class="col" style="width:960px;height:540px;padding:60px;background:white;">
    <h2 style="margin-bottom:1.5rem;">Technical Analysis</h2>
    <ul style="line-height:1.6;">
        ${bullets}
    </ul>
</body></html>
    `;
    fs.writeFileSync('/home/claude/slide4.html', slide4HTML);
    
    // Slide 5: Chart
    const slide5HTML = `
<!DOCTYPE html>
<html><head><style>${sharedCSS}</style></head>
<body class="col center" style="width:960px;height:540px;padding:40px;background:white;">
    <h2 style="margin-bottom:1rem;">Price Chart with Indicators</h2>
    <div class="placeholder" style="width:100%;height:400px;"></div>
</body></html>
    `;
    fs.writeFileSync('/home/claude/slide5.html', slide5HTML);
    
    // Slide 6: Moving Averages Education
    const slide6HTML = `
<!DOCTYPE html>
<html><head><style>${sharedCSS}</style></head>
<body class="row" style="width:960px;height:540px;background:white;">
    <div class="col fill-width" style="padding:60px;background:#ECF0F1;">
        <h2 style="margin-bottom:1.5rem;">Understanding Moving Averages</h2>
        <h3 style="color:#27AE60;margin-bottom:0.5rem;">Golden Cross 🟢</h3>
        <p style="margin-bottom:1.5rem;">50-day MA crosses above 200-day MA → Bullish signal</p>
        <h3 style="color:#E74C3C;margin-bottom:0.5rem;">Death Cross 🔴</h3>
        <p>50-day MA crosses below 200-day MA → Bearish signal</p>
    </div>
    <div class="col fill-width" style="padding:60px;">
        <h3 style="margin-bottom:1rem;">Current Values</h3>
        <div class="metric-box" style="margin-bottom:1rem;">
            <p class="metric-value">$${latest_data.SMA_short.toFixed(2)}</p>
            <p class="metric-label">50-Day MA</p>
        </div>
        <div class="metric-box">
            <p class="metric-value">$${latest_data.SMA_long.toFixed(2)}</p>
            <p class="metric-label">200-Day MA</p>
        </div>
    </div>
</body></html>
    `;
    fs.writeFileSync('/home/claude/slide6.html', slide6HTML);
    
    // Slide 7: RSI Education
    const slide7HTML = `
<!DOCTYPE html>
<html><head><style>${sharedCSS}</style></head>
<body class="col" style="width:960px;height:540px;padding:60px;background:white;">
    <h2 style="margin-bottom:1.5rem;">Understanding RSI (Relative Strength Index)</h2>
    <div class="row gap-lg" style="margin-bottom:2rem;">
        <div class="fill-width" style="background:#E74C3C20;padding:1.5rem;border-left:4px solid #E74C3C;border-radius:4px;">
            <h3 style="color:#E74C3C;margin-bottom:0.5rem;">Overbought (RSI > 70)</h3>
            <p>Price may be due for a pullback. Consider selling or taking profits.</p>
        </div>
        <div class="fill-width" style="background:#F39C1220;padding:1.5rem;border-left:4px solid #F39C12;border-radius:4px;">
            <h3 style="color:#F39C12;margin-bottom:0.5rem;">Neutral (30-70)</h3>
            <p>Balanced momentum. Monitor for trend continuation.</p>
        </div>
        <div class="fill-width" style="background:#27AE6020;padding:1.5rem;border-left:4px solid #27AE60;border-radius:4px;">
            <h3 style="color:#27AE60;margin-bottom:0.5rem;">Oversold (RSI < 30)</h3>
            <p>Price may bounce back. Potential buying opportunity.</p>
        </div>
    </div>
    <div class="highlight" style="text-align:center;">
        <p class="metric-value" style="font-size:48px;color:${latest_data.RSI > 70 ? '#E74C3C' : latest_data.RSI < 30 ? '#27AE60' : '#F39C12'}">${latest_data.RSI.toFixed(2)}</p>
        <p class="metric-label" style="font-size:18px;">Current RSI Value</p>
    </div>
</body></html>
    `;
    fs.writeFileSync('/home/claude/slide7.html', slide7HTML);
    
    // Slide 8: Risk Management
    const slide8HTML = `
<!DOCTYPE html>
<html><head><style>${sharedCSS}</style></head>
<body class="col" style="width:960px;height:540px;padding:60px;background:linear-gradient(135deg, #667eea 0%, #764ba2 100%);">
    <h2 style="color:white;margin-bottom:2rem;">Risk Management Tips</h2>
    <ul style="color:white;font-size:20px;line-height:1.8;">
        <li>Never invest more than you can afford to lose</li>
        <li>Diversify across multiple assets</li>
        <li>Use stop-loss orders to limit losses</li>
        <li>Technical analysis is a tool, not a guarantee</li>
        <li>Consider fundamentals alongside technicals</li>
    </ul>
    <p style="color:white;margin-top:2rem;font-size:16px;opacity:0.9;">Always conduct your own research before making investment decisions.</p>
</body></html>
    `;
    fs.writeFileSync('/home/claude/slide8.html', slide8HTML);
    
    // Create the presentation
    const pptx = new pptxgen();
    pptx.layout = 'LAYOUT_16x9';
    
    // Add all slides
    await html2pptx('/home/claude/slide1.html', pptx);
    await html2pptx('/home/claude/slide2.html', pptx);
    await html2pptx('/home/claude/slide3.html', pptx);
    await html2pptx('/home/claude/slide4.html', pptx);
    
    // Add chart slide with image if available
    const slide5Result = await html2pptx('/home/claude/slide5.html', pptx);
    if (fs.existsSync(chart_path)) {
        slide5Result.slide.addImage({
            path: chart_path,
            x: slide5Result.placeholders[0].x,
            y: slide5Result.placeholders[0].y,
            w: slide5Result.placeholders[0].w,
            h: slide5Result.placeholders[0].h
        });
    }
    
    await html2pptx('/home/claude/slide6.html', pptx);
    await html2pptx('/home/claude/slide7.html', pptx);
    await html2pptx('/home/claude/slide8.html', pptx);
    
    // Save the presentation
    await pptx.writeFile(outputPath);
    console.log(`PowerPoint presentation generated: ${outputPath}`);
    
    // Clean up HTML files
    for (let i = 1; i <= 8; i++) {
        try {
            fs.unlinkSync(`/home/claude/slide${i}.html`);
        } catch (e) {}
    }
}

module.exports = { generatePresentation };
