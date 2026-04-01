#!/usr/bin/env bash
set -euo pipefail

# ============================================================
# OpenCode Corporate PPTX Skill — Installer
# ============================================================

SCRIPT_DIR="$(cd "$(dirname "${BASH_SOURCE[0]}")" && pwd)"

SKILL_DIR="$HOME/.config/opencode/skills/corporate-pptx"
TOOLS_DIR="$HOME/.config/opencode/tools"
ASSETS_DIR="$TOOLS_DIR/assets"

echo "=== OpenCode Corporate PPTX Skill — Installer ==="
echo ""

# ---- Check prerequisites ----

echo "[1/5] Checking prerequisites..."

if ! command -v python3 &>/dev/null; then
    echo "  ERROR: python3 not found. Please install Python 3.8+."
    exit 1
fi
echo "  python3: $(python3 --version)"

# Check python-pptx
if python3 -c "import pptx" 2>/dev/null; then
    echo "  python-pptx: installed"
else
    echo "  python-pptx: not found, installing..."
    pip3 install python-pptx
    echo "  python-pptx: installed"
fi

# ---- Install Skill ----

echo ""
echo "[2/5] Installing Agent Skill..."

mkdir -p "$SKILL_DIR"
cp "$SCRIPT_DIR/skill/SKILL.md" "$SKILL_DIR/SKILL.md"
echo "  -> $SKILL_DIR/SKILL.md"

# ---- Install Custom Tool ----

echo ""
echo "[3/5] Installing Custom Tool..."

mkdir -p "$TOOLS_DIR"
cp "$SCRIPT_DIR/tools/corporate-pptx.ts" "$TOOLS_DIR/corporate-pptx.ts"
cp "$SCRIPT_DIR/tools/generate_pptx.py" "$TOOLS_DIR/generate_pptx.py"
echo "  -> $TOOLS_DIR/corporate-pptx.ts"
echo "  -> $TOOLS_DIR/generate_pptx.py"

# ---- Install Assets ----

echo ""
echo "[4/5] Installing logo assets..."

mkdir -p "$ASSETS_DIR"
cp "$SCRIPT_DIR/tools/assets/"*.png "$ASSETS_DIR/"
echo "  -> $ASSETS_DIR/bankiru_logo.png"
echo "  -> $ASSETS_DIR/bankiru_logo_white.png"
echo "  -> $ASSETS_DIR/bankiru_logo_dark.png"

# ---- Verify ----

echo ""
echo "[5/5] Verifying installation..."

OK=true

if [ ! -f "$SKILL_DIR/SKILL.md" ]; then
    echo "  FAIL: SKILL.md not found"
    OK=false
fi

if [ ! -f "$TOOLS_DIR/corporate-pptx.ts" ]; then
    echo "  FAIL: corporate-pptx.ts not found"
    OK=false
fi

if [ ! -f "$TOOLS_DIR/generate_pptx.py" ]; then
    echo "  FAIL: generate_pptx.py not found"
    OK=false
fi

if [ ! -f "$ASSETS_DIR/bankiru_logo_white.png" ]; then
    echo "  FAIL: bankiru_logo_white.png not found"
    OK=false
fi

if $OK; then
    echo ""
    echo "=== Installation complete! ==="
    echo ""
    echo "Installed files:"
    echo "  Skill:  $SKILL_DIR/SKILL.md"
    echo "  Tool:   $TOOLS_DIR/corporate-pptx.ts"
    echo "  Script: $TOOLS_DIR/generate_pptx.py"
    echo "  Assets: $ASSETS_DIR/bankiru_logo*.png"
    echo ""
    echo "Restart OpenCode to activate the skill."
    echo "Then ask: \"Создай презентацию про ...\""
else
    echo ""
    echo "=== Installation failed. Check errors above. ==="
    exit 1
fi
