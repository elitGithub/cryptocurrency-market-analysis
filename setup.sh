#!/bin/bash
# Setup script for Crypto Trading Bot v2.0

echo "=========================================="
echo "Crypto Trading Bot - Setup Script"
echo "=========================================="
echo ""

# Check Python
if ! command -v python3 &> /dev/null; then
    echo "❌ Python 3 is not installed. Please install Python 3.8 or higher."
    exit 1
fi
echo "✓ Python 3 found: $(python3 --version)"

# Check Node.js
if ! command -v node &> /dev/null; then
    echo "❌ Node.js is not installed. Please install Node.js 14 or higher."
    exit 1
fi
echo "✓ Node.js found: $(node --version)"

# Check npm
if ! command -v npm &> /dev/null; then
    echo "❌ npm is not installed. Please install npm."
    exit 1
fi
echo "✓ npm found: $(npm --version)"

echo ""
echo "Installing Python dependencies..."
pip install -r requirements.txt --break-system-packages

if [ $? -ne 0 ]; then
    echo "❌ Failed to install Python dependencies"
    exit 1
fi
echo "✓ Python dependencies installed"

echo ""
echo "Installing Node.js dependencies..."

# Install docx
echo "  → Installing docx..."
npm install -g docx
if [ $? -ne 0 ]; then
    echo "⚠️  Warning: docx installation failed"
fi

# Install pptxgenjs
echo "  → Installing pptxgenjs..."
npm install -g pptxgenjs
if [ $? -ne 0 ]; then
    echo "⚠️  Warning: pptxgenjs installation failed"
fi

# Install playwright
echo "  → Installing playwright..."
npm install -g playwright
if [ $? -ne 0 ]; then
    echo "⚠️  Warning: playwright installation failed"
fi

# Install html2pptx (if available)
if [ -f "/mnt/skills/public/pptx/html2pptx.tgz" ]; then
    echo "  → Installing html2pptx..."
    npm install -g /mnt/skills/public/pptx/html2pptx.tgz
    if [ $? -ne 0 ]; then
        echo "⚠️  Warning: html2pptx installation failed"
    fi
else
    echo "⚠️  Warning: html2pptx.tgz not found. PowerPoint generation may not work."
fi

echo "✓ Node.js dependencies installed"

echo ""
echo "=========================================="
echo "✅ Setup Complete!"
echo "=========================================="
echo ""
echo "To run the bot:"
echo "  python main.py"
echo ""
echo "Configuration:"
echo "  Edit config.ini to customize settings"
echo ""
echo "Documentation:"
echo "  • README_NEW.md       - Usage guide"
echo "  • MIGRATION_GUIDE.md  - Migration from old code"
echo "  • ARCHITECTURE.md     - Technical details"
echo ""
echo "Output:"
echo "  Reports will be saved to: output/reports/"
echo ""
echo "=========================================="
