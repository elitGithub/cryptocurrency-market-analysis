const fs = require('fs');
const { Document, Packer, Paragraph, TextRun, Table, TableRow, TableCell, ImageRun, 
        AlignmentType, HeadingLevel, BorderStyle, WidthType, LevelFormat, PageBreak,
        ShadingType } = require('docx');

function generateBuySellHold(suggestions, latest_data) {
    // Analyze suggestions to determine recommendation
    let bullish_signals = 0;
    let bearish_signals = 0;
    
    suggestions.forEach(s => {
        if (s.includes('BULLISH') || s.includes('Golden Cross') || s.includes('oversold')) bullish_signals++;
        if (s.includes('BEARISH') || s.includes('Death Cross') || s.includes('overbought')) bearish_signals++;
    });
    
    // Check RSI and MA alignment
    const rsi = latest_data.RSI;
    const ma_bullish = latest_data.SMA_short > latest_data.SMA_long;
    
    if (rsi < 30 || bullish_signals > bearish_signals) {
        return {
            recommendation: 'BUY',
            color: '27AE60',
            reasoning: `Strong bullish signals detected. RSI at ${rsi.toFixed(2)} suggests potential upside. Consider accumulating positions.`
        };
    } else if (rsi > 70 || bearish_signals > bullish_signals) {
        return {
            recommendation: 'SELL',
            color: 'E74C3C',
            reasoning: `Bearish signals present. RSI at ${rsi.toFixed(2)} indicates overbought conditions. Consider taking profits.`
        };
    } else {
        return {
            recommendation: 'HOLD',
            color: 'F39C12',
            reasoning: `Mixed signals. RSI at ${rsi.toFixed(2)} is neutral. Maintain current positions and monitor.`
        };
    }
}

function createDocument(data) {
    const { symbol, exchange_results, suggestions, latest_data, chart_path } = data;
    
    const signal = generateBuySellHold(suggestions, latest_data);
    
    const tableBorder = { style: BorderStyle.SINGLE, size: 1, color: 'CCCCCC' };
    const cellBorders = { top: tableBorder, bottom: tableBorder, left: tableBorder, right: tableBorder };
    
    const doc = new Document({
        styles: {
            default: {
                document: { run: { font: 'Arial', size: 24 } }
            },
            paragraphStyles: [
                {
                    id: 'Title',
                    name: 'Title',
                    basedOn: 'Normal',
                    run: { size: 56, bold: true, color: '1C2833', font: 'Arial' },
                    paragraph: { spacing: { before: 240, after: 120 }, alignment: AlignmentType.CENTER }
                },
                {
                    id: 'Heading1',
                    name: 'Heading 1',
                    basedOn: 'Normal',
                    next: 'Normal',
                    quickFormat: true,
                    run: { size: 32, bold: true, color: '2C3E50', font: 'Arial' },
                    paragraph: { spacing: { before: 240, after: 120 }, outlineLevel: 0 }
                },
                {
                    id: 'Heading2',
                    name: 'Heading 2',
                    basedOn: 'Normal',
                    next: 'Normal',
                    quickFormat: true,
                    run: { size: 28, bold: true, color: '34495E', font: 'Arial' },
                    paragraph: { spacing: { before: 180, after: 100 }, outlineLevel: 1 }
                },
                {
                    id: 'Heading3',
                    name: 'Heading 3',
                    basedOn: 'Normal',
                    next: 'Normal',
                    quickFormat: true,
                    run: { size: 26, bold: true, color: '566573', font: 'Arial' },
                    paragraph: { spacing: { before: 120, after: 80 }, outlineLevel: 2 }
                }
            ]
        },
        numbering: {
            config: [
                {
                    reference: 'bullet-list',
                    levels: [{
                        level: 0,
                        format: LevelFormat.BULLET,
                        text: '•',
                        alignment: AlignmentType.LEFT,
                        style: { paragraph: { indent: { left: 720, hanging: 360 } } }
                    }]
                }
            ]
        },
        sections: [{
            properties: {
                page: {
                    margin: { top: 1440, right: 1440, bottom: 1440, left: 1440 }
                }
            },
            children: [
                // Title
                new Paragraph({
                    heading: HeadingLevel.TITLE,
                    children: [new TextRun(`Cryptocurrency Market Analysis Report`)]
                }),
                new Paragraph({
                    alignment: AlignmentType.CENTER,
                    spacing: { after: 400 },
                    children: [new TextRun({ text: symbol, size: 40, bold: true, color: '3498DB' })]
                }),
                
                // Executive Summary
                new Paragraph({
                    heading: HeadingLevel.HEADING_1,
                    children: [new TextRun('Executive Summary')]
                }),
                
                // Trading Recommendation Box
                new Paragraph({
                    spacing: { before: 120, after: 120 },
                    alignment: AlignmentType.CENTER,
                    children: [new TextRun({
                        text: `TRADING RECOMMENDATION: ${signal.recommendation}`,
                        bold: true,
                        size: 36,
                        color: signal.color
                    })]
                }),
                
                new Paragraph({
                    spacing: { after: 240 },
                    children: [new TextRun(signal.reasoning)]
                }),
                
                // Key Metrics Table
                new Paragraph({
                    heading: HeadingLevel.HEADING_2,
                    children: [new TextRun('Current Market Metrics')]
                }),
                
                new Table({
                    columnWidths: [4680, 4680],
                    margins: { top: 100, bottom: 100, left: 180, right: 180 },
                    rows: [
                        new TableRow({
                            tableHeader: true,
                            children: [
                                new TableCell({
                                    borders: cellBorders,
                                    width: { size: 4680, type: WidthType.DXA },
                                    shading: { fill: '3498DB', type: ShadingType.CLEAR },
                                    children: [new Paragraph({
                                        alignment: AlignmentType.CENTER,
                                        children: [new TextRun({ text: 'Metric', bold: true, color: 'FFFFFF', size: 24 })]
                                    })]
                                }),
                                new TableCell({
                                    borders: cellBorders,
                                    width: { size: 4680, type: WidthType.DXA },
                                    shading: { fill: '3498DB', type: ShadingType.CLEAR },
                                    children: [new Paragraph({
                                        alignment: AlignmentType.CENTER,
                                        children: [new TextRun({ text: 'Value', bold: true, color: 'FFFFFF', size: 24 })]
                                    })]
                                })
                            ]
                        }),
                        new TableRow({
                            children: [
                                new TableCell({
                                    borders: cellBorders,
                                    width: { size: 4680, type: WidthType.DXA },
                                    children: [new Paragraph({ children: [new TextRun({ text: 'Current Price', bold: true })] })]
                                }),
                                new TableCell({
                                    borders: cellBorders,
                                    width: { size: 4680, type: WidthType.DXA },
                                    children: [new Paragraph({ children: [new TextRun(`$${latest_data.close.toFixed(2)} USDT`)] })]
                                })
                            ]
                        }),
                        new TableRow({
                            children: [
                                new TableCell({
                                    borders: cellBorders,
                                    width: { size: 4680, type: WidthType.DXA },
                                    children: [new Paragraph({ children: [new TextRun({ text: 'RSI (14)', bold: true })] })]
                                }),
                                new TableCell({
                                    borders: cellBorders,
                                    width: { size: 4680, type: WidthType.DXA },
                                    children: [new Paragraph({ children: [new TextRun(latest_data.RSI.toFixed(2))] })]
                                })
                            ]
                        }),
                        new TableRow({
                            children: [
                                new TableCell({
                                    borders: cellBorders,
                                    width: { size: 4680, type: WidthType.DXA },
                                    children: [new Paragraph({ children: [new TextRun({ text: '50-Day MA', bold: true })] })]
                                }),
                                new TableCell({
                                    borders: cellBorders,
                                    width: { size: 4680, type: WidthType.DXA },
                                    children: [new Paragraph({ children: [new TextRun(`$${latest_data.SMA_short.toFixed(2)}`)] })]
                                })
                            ]
                        }),
                        new TableRow({
                            children: [
                                new TableCell({
                                    borders: cellBorders,
                                    width: { size: 4680, type: WidthType.DXA },
                                    children: [new Paragraph({ children: [new TextRun({ text: '200-Day MA', bold: true })] })]
                                }),
                                new TableCell({
                                    borders: cellBorders,
                                    width: { size: 4680, type: WidthType.DXA },
                                    children: [new Paragraph({ children: [new TextRun(`$${latest_data.SMA_long.toFixed(2)}`)] })]
                                })
                            ]
                        }),
                        new TableRow({
                            children: [
                                new TableCell({
                                    borders: cellBorders,
                                    width: { size: 4680, type: WidthType.DXA },
                                    children: [new Paragraph({ children: [new TextRun({ text: 'Trend', bold: true })] })]
                                }),
                                new TableCell({
                                    borders: cellBorders,
                                    width: { size: 4680, type: WidthType.DXA },
                                    children: [new Paragraph({ 
                                        children: [new TextRun({ 
                                            text: latest_data.SMA_short > latest_data.SMA_long ? 'Bullish ↑' : 'Bearish ↓',
                                            color: latest_data.SMA_short > latest_data.SMA_long ? '27AE60' : 'E74C3C',
                                            bold: true
                                        })] 
                                    })]
                                })
                            ]
                        })
                    ]
                }),
                
                new Paragraph({ children: [new PageBreak()] }),
                
                // Exchange Comparison
                new Paragraph({
                    heading: HeadingLevel.HEADING_1,
                    children: [new TextRun('Exchange Comparison')]
                }),
                
                new Paragraph({
                    spacing: { after: 120 },
                    children: [new TextRun('Based on our analysis of multiple cryptocurrency exchanges, here are the key findings:')]
                }),
                
                ...createExchangeTable(exchange_results, cellBorders),
                
                new Paragraph({
                    heading: HeadingLevel.HEADING_2,
                    spacing: { before: 240 },
                    children: [new TextRun('Recommended Exchanges')]
                }),
                
                ...createExchangeRecommendations(exchange_results),
                
                new Paragraph({ children: [new PageBreak()] }),
                
                // Technical Analysis
                new Paragraph({
                    heading: HeadingLevel.HEADING_1,
                    children: [new TextRun('Detailed Technical Analysis')]
                }),
                
                ...createTechnicalAnalysis(suggestions),
                
                // Chart
                new Paragraph({
                    heading: HeadingLevel.HEADING_2,
                    spacing: { before: 240 },
                    children: [new TextRun('Price Chart')]
                }),
                
                ...(fs.existsSync(chart_path) ? [
                    new Paragraph({
                        alignment: AlignmentType.CENTER,
                        spacing: { before: 120, after: 240 },
                        children: [new ImageRun({
                            type: 'png',
                            data: fs.readFileSync(chart_path),
                            transformation: { width: 600, height: 400 },
                            altText: { title: 'Price Chart', description: 'Technical analysis chart', name: 'Chart' }
                        })]
                    })
                ] : [
                    new Paragraph({
                        spacing: { after: 240 },
                        children: [new TextRun({ text: 'Chart generation pending...', italics: true })]
                    })
                ]),
                
                new Paragraph({ children: [new PageBreak()] }),
                
                // Educational Section
                new Paragraph({
                    heading: HeadingLevel.HEADING_1,
                    children: [new TextRun('Understanding Technical Indicators')]
                }),
                
                new Paragraph({
                    spacing: { after: 180 },
                    children: [new TextRun('This section explains the key technical indicators used in this analysis.')]
                }),
                
                ...createEducationalContent()
            ]
        }]
    });
    
    return doc;
}

function createExchangeTable(results, cellBorders) {
    const rows = [
        new TableRow({
            tableHeader: true,
            children: [
                new TableCell({
                    borders: cellBorders,
                    width: { size: 2340, type: WidthType.DXA },
                    shading: { fill: '2C3E50', type: ShadingType.CLEAR },
                    children: [new Paragraph({
                        alignment: AlignmentType.CENTER,
                        children: [new TextRun({ text: 'Exchange', bold: true, color: 'FFFFFF', size: 22 })]
                    })]
                }),
                new TableCell({
                    borders: cellBorders,
                    width: { size: 2340, type: WidthType.DXA },
                    shading: { fill: '2C3E50', type: ShadingType.CLEAR },
                    children: [new Paragraph({
                        alignment: AlignmentType.CENTER,
                        children: [new TextRun({ text: 'Spot Pairs', bold: true, color: 'FFFFFF', size: 22 })]
                    })]
                }),
                new TableCell({
                    borders: cellBorders,
                    width: { size: 2340, type: WidthType.DXA },
                    shading: { fill: '2C3E50', type: ShadingType.CLEAR },
                    children: [new Paragraph({
                        alignment: AlignmentType.CENTER,
                        children: [new TextRun({ text: 'USDT Pairs', bold: true, color: 'FFFFFF', size: 22 })]
                    })]
                }),
                new TableCell({
                    borders: cellBorders,
                    width: { size: 2340, type: WidthType.DXA },
                    shading: { fill: '2C3E50', type: ShadingType.CLEAR },
                    children: [new Paragraph({
                        alignment: AlignmentType.CENTER,
                        children: [new TextRun({ text: 'OHLCV Data', bold: true, color: 'FFFFFF', size: 22 })]
                    })]
                })
            ]
        })
    ];
    
    results.forEach(ex => {
        rows.push(new TableRow({
            children: [
                new TableCell({
                    borders: cellBorders,
                    width: { size: 2340, type: WidthType.DXA },
                    children: [new Paragraph({ children: [new TextRun({ text: ex.name, bold: true })] })]
                }),
                new TableCell({
                    borders: cellBorders,
                    width: { size: 2340, type: WidthType.DXA },
                    children: [new Paragraph({ 
                        alignment: AlignmentType.CENTER,
                        children: [new TextRun(ex.total_spot_pairs.toString())] 
                    })]
                }),
                new TableCell({
                    borders: cellBorders,
                    width: { size: 2340, type: WidthType.DXA },
                    children: [new Paragraph({ 
                        alignment: AlignmentType.CENTER,
                        children: [new TextRun(ex.usdt_quoted_pairs.toString())] 
                    })]
                }),
                new TableCell({
                    borders: cellBorders,
                    width: { size: 2340, type: WidthType.DXA },
                    children: [new Paragraph({ 
                        alignment: AlignmentType.CENTER,
                        children: [new TextRun({ 
                            text: ex.supports_fetchOHLCV ? '✓' : '✗',
                            color: ex.supports_fetchOHLCV ? '27AE60' : 'E74C3C',
                            bold: true
                        })] 
                    })]
                })
            ]
        }));
    });
    
    return [new Table({
        columnWidths: [2340, 2340, 2340, 2340],
        margins: { top: 100, bottom: 100, left: 180, right: 180 },
        rows: rows
    })];
}

function createExchangeRecommendations(results) {
    const recommendations = [];
    const best = results.sort((a, b) => b.usdt_quoted_pairs - a.usdt_quoted_pairs)[0];
    
    recommendations.push(
        new Paragraph({
            numbering: { reference: 'bullet-list', level: 0 },
            spacing: { after: 100 },
            children: [new TextRun({ text: `${best.name}`, bold: true }), 
                      new TextRun(` appears to be the best option with ${best.usdt_quoted_pairs} USDT trading pairs and full OHLCV data support.`)]
        })
    );
    
    results.filter(r => r.supports_fetchOHLCV).forEach(ex => {
        if (ex.name !== best.name) {
            recommendations.push(
                new Paragraph({
                    numbering: { reference: 'bullet-list', level: 0 },
                    spacing: { after: 100 },
                    children: [new TextRun({ text: ex.name, bold: true }), 
                              new TextRun(` is also a viable option with ${ex.usdt_quoted_pairs} USDT pairs.`)]
                })
            );
        }
    });
    
    return recommendations;
}

function createTechnicalAnalysis(suggestions) {
    const paragraphs = [];
    
    suggestions.forEach(suggestion => {
        const isSignal = suggestion.includes('SIGNAL');
        const isTrend = suggestion.includes('TREND');
        const isWarning = suggestion.includes('WARNING');
        const isOpportunity = suggestion.includes('OPPORTUNITY');
        
        let heading = 'Technical Signal';
        if (isWarning) heading = 'Warning';
        if (isOpportunity) heading = 'Opportunity';
        if (isTrend) heading = 'Current Trend';
        
        paragraphs.push(
            new Paragraph({
                heading: HeadingLevel.HEADING_3,
                spacing: { before: 180 },
                children: [new TextRun(heading)]
            })
        );
        
        paragraphs.push(
            new Paragraph({
                spacing: { after: 180 },
                children: [new TextRun(suggestion)]
            })
        );
    });
    
    return paragraphs;
}

function createEducationalContent() {
    return [
        new Paragraph({
            heading: HeadingLevel.HEADING_2,
            children: [new TextRun('Moving Averages (MA)')]
        }),
        new Paragraph({
            spacing: { after: 120 },
            children: [new TextRun('Moving averages smooth out price data to identify trends. This analysis uses:')]
        }),
        new Paragraph({
            numbering: { reference: 'bullet-list', level: 0 },
            children: [new TextRun({ text: '50-Day MA (Short-term): ', bold: true }), 
                      new TextRun('Tracks recent price momentum and short-term trends.')]
        }),
        new Paragraph({
            numbering: { reference: 'bullet-list', level: 0 },
            spacing: { after: 180 },
            children: [new TextRun({ text: '200-Day MA (Long-term): ', bold: true }), 
                      new TextRun('Represents major trend direction and acts as support/resistance.')]
        }),
        
        new Paragraph({
            heading: HeadingLevel.HEADING_3,
            children: [new TextRun('Golden Cross & Death Cross')]
        }),
        new Paragraph({
            numbering: { reference: 'bullet-list', level: 0 },
            children: [new TextRun({ text: 'Golden Cross: ', bold: true, color: '27AE60' }), 
                      new TextRun('When the 50-day MA crosses above the 200-day MA, signaling potential uptrend.')]
        }),
        new Paragraph({
            numbering: { reference: 'bullet-list', level: 0 },
            spacing: { after: 240 },
            children: [new TextRun({ text: 'Death Cross: ', bold: true, color: 'E74C3C' }), 
                      new TextRun('When the 50-day MA crosses below the 200-day MA, signaling potential downtrend.')]
        }),
        
        new Paragraph({
            heading: HeadingLevel.HEADING_2,
            children: [new TextRun('Relative Strength Index (RSI)')]
        }),
        new Paragraph({
            spacing: { after: 120 },
            children: [new TextRun('RSI measures momentum on a scale of 0-100 to identify overbought or oversold conditions.')]
        }),
        new Paragraph({
            numbering: { reference: 'bullet-list', level: 0 },
            children: [new TextRun({ text: 'RSI > 70: ', bold: true, color: 'E74C3C' }), 
                      new TextRun('Overbought territory. Price may be due for a pullback.')]
        }),
        new Paragraph({
            numbering: { reference: 'bullet-list', level: 0 },
            children: [new TextRun({ text: 'RSI < 30: ', bold: true, color: '27AE60' }), 
                      new TextRun('Oversold territory. Price may be due for a bounce.')]
        }),
        new Paragraph({
            numbering: { reference: 'bullet-list', level: 0 },
            spacing: { after: 240 },
            children: [new TextRun({ text: 'RSI 30-70: ', bold: true }), 
                      new TextRun('Neutral zone with balanced momentum.')]
        }),
        
        new Paragraph({
            heading: HeadingLevel.HEADING_2,
            children: [new TextRun('Bollinger Bands')]
        }),
        new Paragraph({
            spacing: { after: 120 },
            children: [new TextRun('Bollinger Bands measure volatility using standard deviations from a moving average.')]
        }),
        new Paragraph({
            numbering: { reference: 'bullet-list', level: 0 },
            children: [new TextRun({ text: 'Upper Band: ', bold: true }), 
                      new TextRun('When price touches the upper band, it may indicate overbought conditions.')]
        }),
        new Paragraph({
            numbering: { reference: 'bullet-list', level: 0 },
            children: [new TextRun({ text: 'Lower Band: ', bold: true }), 
                      new TextRun('When price touches the lower band, it may indicate oversold conditions.')]
        }),
        new Paragraph({
            numbering: { reference: 'bullet-list', level: 0 },
            spacing: { after: 240 },
            children: [new TextRun({ text: 'Band Squeeze: ', bold: true }), 
                      new TextRun('When bands narrow, it signals low volatility and potential breakout.')]
        }),
        
        new Paragraph({
            heading: HeadingLevel.HEADING_2,
            children: [new TextRun('Risk Management Tips')]
        }),
        new Paragraph({
            numbering: { reference: 'bullet-list', level: 0 },
            children: [new TextRun('Never invest more than you can afford to lose')]
        }),
        new Paragraph({
            numbering: { reference: 'bullet-list', level: 0 },
            children: [new TextRun('Diversify across multiple assets to reduce risk')]
        }),
        new Paragraph({
            numbering: { reference: 'bullet-list', level: 0 },
            children: [new TextRun('Use stop-loss orders to limit potential losses')]
        }),
        new Paragraph({
            numbering: { reference: 'bullet-list', level: 0 },
            children: [new TextRun('Technical analysis is not foolproof - always do your own research')]
        }),
        new Paragraph({
            numbering: { reference: 'bullet-list', level: 0 },
            children: [new TextRun('Consider market sentiment, news, and fundamental analysis alongside technical signals')]
        })
    ];
}

async function generateWordReport(data, outputPath) {
    const doc = createDocument(data);
    const buffer = await Packer.toBuffer(doc);
    fs.writeFileSync(outputPath, buffer);
    console.log(`Word document generated: ${outputPath}`);
}

module.exports = { generateWordReport };
