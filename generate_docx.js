
const { Document, Packer, Paragraph, TextRun, Table, TableRow, TableCell, AlignmentType, 
        HeadingLevel, BorderStyle, WidthType, ShadingType, VerticalAlign, LevelFormat, PageBreak } = require('docx');
const fs = require('fs');

const signal_color = {'BUY': 'GREEN', 'SELL': 'RED', 'HOLD': 'ORANGE'};
const signal = 'HOLD';
const confidence = 'MEDIUM';

const tableBorder = { style: BorderStyle.SINGLE, size: 1, color: "CCCCCC" };

const doc = new Document({
    styles: {
        default: { 
            document: { run: { font: "Arial", size: 24 } } 
        },
        paragraphStyles: [
            { id: "Title", name: "Title", basedOn: "Normal",
               run: { size: 56, bold: true, color: "1C2833", font: "Arial" },
               paragraph: { spacing: { before: 240, after: 120 }, alignment: AlignmentType.CENTER } },
            { id: "Heading1", name: "Heading 1", basedOn: "Normal", quickFormat: true,
               run: { size: 32, bold: true, color: "2E4053", font: "Arial" },
               paragraph: { spacing: { before: 240, after: 180 }, outlineLevel: 0 } },
            { id: "Heading2", name: "Heading 2", basedOn: "Normal", quickFormat: true,
               run: { size: 28, bold: true, color: "2E4053", font: "Arial" },
               paragraph: { spacing: { before: 180, after: 120 }, outlineLevel: 1 } },
            { id: "Heading3", name: "Heading 3", basedOn: "Normal", quickFormat: true,
               run: { size: 24, bold: true, color: "34495E", font: "Arial" },
               paragraph: { spacing: { before: 120, after: 80 }, outlineLevel: 2 } }
        ]
    },
    numbering: {
        config: [
            { reference: "bullet-list",
               levels: [{ level: 0, format: LevelFormat.BULLET, text: "•", alignment: AlignmentType.LEFT,
                          style: { paragraph: { indent: { left: 720, hanging: 360 } } } }] },
            { reference: "numbered-list",
               levels: [{ level: 0, format: LevelFormat.DECIMAL, text: "%1.", alignment: AlignmentType.LEFT,
                          style: { paragraph: { indent: { left: 720, hanging: 360 } } } }] }
        ]
    },
    sections: [{
        properties: {
            page: { margin: { top: 1440, right: 1440, bottom: 1440, left: 1440 } }
        },
        children: [
            // Title
            new Paragraph({
                heading: HeadingLevel.TITLE,
                children: [new TextRun("Cryptocurrency Market Analysis Report")]
            }),
            new Paragraph({
                alignment: AlignmentType.CENTER,
                spacing: { after: 400 },
                children: [new TextRun({
                    text: "BTC/USDT",
                    size: 32,
                    color: "7F8C8D"
                })]
            }),
            
            // Executive Summary
            new Paragraph({
                heading: HeadingLevel.HEADING_1,
                children: [new TextRun("Executive Summary")]
            }),
            new Paragraph({
                spacing: { after: 200 },
                children: [
                    new TextRun("Trading Recommendation: "),
                    new TextRun({
                        text: signal,
                        bold: true,
                        size: 32,
                        color: signal === 'BUY' ? '2ECC71' : signal === 'SELL' ? 'E74C3C' : 'F39C12'
                    })
                ]
            }),
            new Paragraph({
                spacing: { after: 200 },
                children: [
                    new TextRun("Confidence Level: "),
                    new TextRun({ text: confidence, bold: true, size: 24 })
                ]
            }),
            new Paragraph({
                spacing: { after: 300 },
                children: [new TextRun({
                    text: "Uptrend confirmed by moving averages | Neutral momentum (RSI: 45.0)",
                    italics: true
                })]
            }),
            
            new Paragraph({
                spacing: { after: 200 },
                children: [
                    new TextRun("Current Price: "),
                    new TextRun({
                        text: "$110,059.99",
                        bold: true,
                        size: 28
                    })
                ]
            }),
            new Paragraph({
                spacing: { after: 400 },
                children: [
                    new TextRun("RSI: "),
                    new TextRun({
                        text: "45.0",
                        bold: true,
                        size: 28
                    })
                ]
            }),
            
            // Page Break
            new Paragraph({ children: [new PageBreak()] }),
            
            // Exchange Analysis
            new Paragraph({
                heading: HeadingLevel.HEADING_1,
                children: [new TextRun("Exchange Comparison")]
            }),
            new Paragraph({
                spacing: { after: 200 },
                children: [new TextRun("Below is a comparison of exchanges available for trading. Choose exchanges with good USDT pair availability and OHLCV support for technical analysis.")]
            }),

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

                    new TableRow({
                        children: [
                            new TableCell({
                                borders: { top: tableBorder, bottom: tableBorder, left: tableBorder, right: tableBorder },
                                width: { size: 2340, type: WidthType.DXA },
                                children: [new Paragraph({ children: [new TextRun("binance")] })]
                            }),
                            new TableCell({
                                borders: { top: tableBorder, bottom: tableBorder, left: tableBorder, right: tableBorder },
                                width: { size: 2340, type: WidthType.DXA },
                                children: [new Paragraph({ alignment: AlignmentType.CENTER, children: [new TextRun("1593")] })]
                            }),
                            new TableCell({
                                borders: { top: tableBorder, bottom: tableBorder, left: tableBorder, right: tableBorder },
                                width: { size: 2340, type: WidthType.DXA },
                                children: [new Paragraph({ alignment: AlignmentType.CENTER, children: [new TextRun("434")] })]
                            }),
                            new TableCell({
                                borders: { top: tableBorder, bottom: tableBorder, left: tableBorder, right: tableBorder },
                                width: { size: 2340, type: WidthType.DXA },
                                children: [new Paragraph({ alignment: AlignmentType.CENTER, children: [new TextRun("Yes")] })]
                            })
                        ]
                    }),

                    new TableRow({
                        children: [
                            new TableCell({
                                borders: { top: tableBorder, bottom: tableBorder, left: tableBorder, right: tableBorder },
                                width: { size: 2340, type: WidthType.DXA },
                                children: [new Paragraph({ children: [new TextRun("kucoin")] })]
                            }),
                            new TableCell({
                                borders: { top: tableBorder, bottom: tableBorder, left: tableBorder, right: tableBorder },
                                width: { size: 2340, type: WidthType.DXA },
                                children: [new Paragraph({ alignment: AlignmentType.CENTER, children: [new TextRun("1311")] })]
                            }),
                            new TableCell({
                                borders: { top: tableBorder, bottom: tableBorder, left: tableBorder, right: tableBorder },
                                width: { size: 2340, type: WidthType.DXA },
                                children: [new Paragraph({ alignment: AlignmentType.CENTER, children: [new TextRun("1081")] })]
                            }),
                            new TableCell({
                                borders: { top: tableBorder, bottom: tableBorder, left: tableBorder, right: tableBorder },
                                width: { size: 2340, type: WidthType.DXA },
                                children: [new Paragraph({ alignment: AlignmentType.CENTER, children: [new TextRun("Yes")] })]
                            })
                        ]
                    }),

                    new TableRow({
                        children: [
                            new TableCell({
                                borders: { top: tableBorder, bottom: tableBorder, left: tableBorder, right: tableBorder },
                                width: { size: 2340, type: WidthType.DXA },
                                children: [new Paragraph({ children: [new TextRun("bybit")] })]
                            }),
                            new TableCell({
                                borders: { top: tableBorder, bottom: tableBorder, left: tableBorder, right: tableBorder },
                                width: { size: 2340, type: WidthType.DXA },
                                children: [new Paragraph({ alignment: AlignmentType.CENTER, children: [new TextRun("685")] })]
                            }),
                            new TableCell({
                                borders: { top: tableBorder, bottom: tableBorder, left: tableBorder, right: tableBorder },
                                width: { size: 2340, type: WidthType.DXA },
                                children: [new Paragraph({ alignment: AlignmentType.CENTER, children: [new TextRun("510")] })]
                            }),
                            new TableCell({
                                borders: { top: tableBorder, bottom: tableBorder, left: tableBorder, right: tableBorder },
                                width: { size: 2340, type: WidthType.DXA },
                                children: [new Paragraph({ alignment: AlignmentType.CENTER, children: [new TextRun("Yes")] })]
                            })
                        ]
                    }),

                    new TableRow({
                        children: [
                            new TableCell({
                                borders: { top: tableBorder, bottom: tableBorder, left: tableBorder, right: tableBorder },
                                width: { size: 2340, type: WidthType.DXA },
                                children: [new Paragraph({ children: [new TextRun("gate")] })]
                            }),
                            new TableCell({
                                borders: { top: tableBorder, bottom: tableBorder, left: tableBorder, right: tableBorder },
                                width: { size: 2340, type: WidthType.DXA },
                                children: [new Paragraph({ alignment: AlignmentType.CENTER, children: [new TextRun("2602")] })]
                            }),
                            new TableCell({
                                borders: { top: tableBorder, bottom: tableBorder, left: tableBorder, right: tableBorder },
                                width: { size: 2340, type: WidthType.DXA },
                                children: [new Paragraph({ alignment: AlignmentType.CENTER, children: [new TextRun("2450")] })]
                            }),
                            new TableCell({
                                borders: { top: tableBorder, bottom: tableBorder, left: tableBorder, right: tableBorder },
                                width: { size: 2340, type: WidthType.DXA },
                                children: [new Paragraph({ alignment: AlignmentType.CENTER, children: [new TextRun("Yes")] })]
                            })
                        ]
                    }),

                ]
            }),
            
            // Page Break
            new Paragraph({ children: [new PageBreak()] }),
            
            // Detailed Analysis
            new Paragraph({
                heading: HeadingLevel.HEADING_1,
                children: [new TextRun("Technical Analysis Details")]
            }),

            new Paragraph({
                spacing: { after: 150 },
                numbering: { reference: "bullet-list", level: 0 },
                children: [new TextRun("CURRENT TREND: The asset appears to be in a long-term uptrend, as the short-term moving average is currently above the long-term average.")]
            }),

            new Paragraph({
                spacing: { after: 150 },
                numbering: { reference: "bullet-list", level: 0 },
                children: [new TextRun("MOMENTUM: The RSI is currently neutral (RSI = 45.05), not indicating extreme overbought or oversold conditions.")]
            }),

            
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
