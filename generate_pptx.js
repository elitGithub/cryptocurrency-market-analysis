
const pptxgen = require("pptxgenjs");
const { html2pptx } = require("@ant/html2pptx");
const fs = require('fs');

(async () => {
    const pptx = new pptxgen();
    pptx.layout = "LAYOUT_16x9";
    pptx.author = "Crypto Trading Bot";
    pptx.title = "Market Analysis Report - BTC/USDT";
    
    // Slide 1: Title
    await html2pptx("slide1.html", pptx);
    
    // Slide 2: Recommendation
    await html2pptx("slide2.html", pptx);
    
    // Slide 3: Exchange Comparison
    const { slide: slide3, placeholders: p3 } = await html2pptx("slide3.html", pptx);
    
    // Add exchange comparison table
    const tableData = [
        [{'text': 'Exchange', 'options': {'fill': {'color': '2E4053'}, 'color': 'FFFFFF', 'bold': true}},
         {'text': 'Total Pairs', 'options': {'fill': {'color': '2E4053'}, 'color': 'FFFFFF', 'bold': true}},
         {'text': 'USDT Pairs', 'options': {'fill': {'color': '2E4053'}, 'color': 'FFFFFF', 'bold': true}},
         {'text': 'Has OHLCV', 'options': {'fill': {'color': '2E4053'}, 'color': 'FFFFFF', 'bold': true}}]
    ];
    
    const exchanges = [{"name": "binance", "total_spot_pairs": 1593, "usdt_quoted_pairs": 434, "supports_fetchOHLCV": true, "rate_limit_ms": 50}, {"name": "kucoin", "total_spot_pairs": 1311, "usdt_quoted_pairs": 1081, "supports_fetchOHLCV": true, "rate_limit_ms": 10}, {"name": "bybit", "total_spot_pairs": 685, "usdt_quoted_pairs": 510, "supports_fetchOHLCV": true, "rate_limit_ms": 20}, {"name": "gate", "total_spot_pairs": 2602, "usdt_quoted_pairs": 2450, "supports_fetchOHLCV": true, "rate_limit_ms": 50}];
    exchanges.forEach(ex => {
        tableData.push([
            ex.name,
            ex.total_spot_pairs.toString(),
            ex.usdt_quoted_pairs.toString(),
            ex.supports_fetchOHLCV ? 'Yes' : 'No'
        ]);
    });
    
    slide3.addTable(tableData, {
        ...p3[0],
        border: { pt: 1, color: 'CCCCCC' },
        fontSize: 16,
        align: 'center',
        valign: 'middle'
    });
    
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
})();
