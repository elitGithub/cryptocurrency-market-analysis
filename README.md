# Crypto Trading Bot - Market Analysis Tool

## Overview
This bot analyzes cryptocurrency markets and generates comprehensive reports in both PowerPoint and Word formats with clear BUY/SELL/HOLD recommendations.

## What You Get
When you run this bot, you'll receive:

1. **PowerPoint Presentation** (`Crypto_Market_Analysis.pptx`)
   - Clear trading signal (BUY/SELL/HOLD) with confidence level
   - Exchange comparison
   - Technical analysis chart
   - Explanations of what the indicators mean
   - Educational content about trading concepts

2. **Word Document** (`Crypto_Market_Analysis.docx`)
   - Executive summary with recommendation
   - Exchange analysis table
   - Detailed technical analysis
   - Educational guide for beginners
   - Getting started tips

## Installation

### 1. Install Python Dependencies
```bash
pip install -r requirements.txt --break-system-packages
```

### 2. Install Node.js Dependencies (for report generation)
```bash
# Install docx library for Word documents
npm install -g docx

# Install html2pptx for PowerPoint presentations
npm install -g /mnt/skills/public/pptx/html2pptx.tgz

# Install pptxgenjs
npm install -g pptxgenjs

# Install playwright (needed for html2pptx)
npm install -g playwright
```

## Running the Bot

Simply run:
```bash
python main.py
```

The bot will:
1. Connect to multiple exchanges (Binance, KuCoin, Bybit, Gate)
2. Analyze BTC/USDT (or your configured symbol)
3. Generate technical analysis
4. Create both PowerPoint and Word reports
5. Save them to `/mnt/user-data/outputs/`

## Understanding the Reports

### Trading Signals
- **BUY**: Multiple indicators suggest this is a good entry point
- **SELL**: Indicators suggest taking profits or avoiding entry
- **HOLD**: Mixed signals, wait for clearer direction

### Confidence Levels
- **HIGH**: 3+ indicators align
- **MEDIUM**: 2 indicators align
- **LOW**: 1 or fewer indicators align

### Key Indicators Explained

**Moving Averages (MA)**
- Golden Cross = Bullish (BUY signal)
- Death Cross = Bearish (SELL signal)

**RSI (Relative Strength Index)**
- Below 30 = Oversold (potential BUY)
- Above 70 = Overbought (potential SELL)
- 30-70 = Neutral

**Bollinger Bands**
- Price at lower band = Potential bounce
- Price at upper band = Potential correction

## Configuration

Edit `config.ini` to customize:

```ini
[exchanges]
target_exchanges = binance, kucoin, bybit, gate

[analysis]
symbol = BTC/USDT          # Change this to analyze other coins
timeframe = 1d             # 1d = daily, 4h = 4-hour, 1h = hourly
history_days = 730         # How many days of history to analyze
short_ma = 50              # Short-term moving average period
long_ma = 200              # Long-term moving average period
```

## For Beginners

### Step 1: Run the Bot
Just execute `python main.py` and let it generate the reports.

### Step 2: Open the Reports
- Open the PowerPoint for a visual overview
- Open the Word document for detailed reading

### Step 3: Understand the Recommendation
- Look at the trading signal (BUY/SELL/HOLD)
- Check the confidence level
- Read the reasoning provided

### Step 4: Learn More
Both documents include educational sections explaining:
- What each indicator means
- How to read the signals
- Tips for getting started

## Important Disclaimers

⚠️ **This tool is for educational purposes only**
- Not financial advice
- Always do your own research (DYOR)
- Never invest more than you can afford to lose
- Cryptocurrency trading carries substantial risk
- Consider consulting a financial advisor

## Troubleshooting

### "Module not found" errors
Run: `pip install -r requirements.txt --break-system-packages`

### Report generation fails
Ensure Node.js packages are installed:
```bash
npm install -g docx pptxgenjs playwright
npm install -g /mnt/skills/public/pptx/html2pptx.tgz
```

### Exchange connection issues
Some exchanges may require API keys for certain features. The bot works with public data by default.

## Support

If you encounter issues:
1. Check that all dependencies are installed
2. Verify your internet connection (needed to fetch market data)
3. Review the console output for specific error messages

## Next Steps

After running the bot successfully:
1. Review the generated reports
2. Learn about the indicators
3. Practice reading the signals
4. Start with paper trading (simulated trading)
5. Only trade with real money once you're comfortable

Good luck with your trading journey! 🚀
