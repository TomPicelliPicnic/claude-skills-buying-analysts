#!/usr/bin/env bash
# setup.sh — one-command onboarding for NT-slides-check
# Run once after cloning: ./setup.sh

set -euo pipefail

REPO_DIR="$(cd "$(dirname "${BASH_SOURCE[0]}")" && pwd)"
SKILLS_DIR="$HOME/.claude/skills"
LINK_NAME="NT-slides-check"
LINK_PATH="$SKILLS_DIR/$LINK_NAME"
CREDS_PATH="$HOME/.cache/picnic-google-sheets/authorized_user.json"

echo "━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━"
echo "  NT-slides-check setup"
echo "━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━"

# ── 1. Poetry ─────────────────────────────────────────
if ! command -v poetry &>/dev/null; then
    echo ""
    echo "ERROR: Poetry is not installed."
    echo "Install it from: https://python-poetry.org/docs/#installation"
    echo "Then re-run this script."
    exit 1
fi
echo ""
echo "✅ Poetry: $(poetry --version)"

# ── 2. Dependencies ───────────────────────────────────
echo ""
echo "Installing Python dependencies..."
cd "$REPO_DIR"
poetry install
echo "✅ Dependencies ready."

# ── 3. Claude Code symlink ────────────────────────────
echo ""
mkdir -p "$SKILLS_DIR"
if [ -L "$LINK_PATH" ]; then
    current="$(readlink "$LINK_PATH")"
    if [ "$current" = "$REPO_DIR" ]; then
        echo "✅ Symlink already correct:"
        echo "   $LINK_PATH → $REPO_DIR"
    else
        ln -sf "$REPO_DIR" "$LINK_PATH"
        echo "✅ Symlink updated:"
        echo "   $LINK_PATH → $REPO_DIR  (was → $current)"
    fi
elif [ -e "$LINK_PATH" ]; then
    echo "ERROR: $LINK_PATH exists but is not a symlink."
    echo "Remove it manually and re-run this script:"
    echo "  rm -rf \"$LINK_PATH\""
    exit 1
else
    ln -s "$REPO_DIR" "$LINK_PATH"
    echo "✅ Symlink created:"
    echo "   $LINK_PATH → $REPO_DIR"
fi

# ── 4. Google Sheets credentials ──────────────────────
echo ""
if [ -f "$CREDS_PATH" ]; then
    echo "✅ Google Sheets credentials found."
else
    echo "⚠️  Credentials not found at:"
    echo "   $CREDS_PATH"
    echo ""
    echo "   Ask a teammate to walk you through the auth setup:"
    echo "   cd ~/.claude/skills/picnic-gsheet && poetry run python gsheet_auth_setup.py"
fi

# ── Done ──────────────────────────────────────────────
echo ""
echo "━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━"
echo "  Setup complete!"
echo "━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━"
echo ""
echo "  Run a check from Claude Code:"
echo "    /NT-slides-check <google_sheet_url>"
echo ""
echo "  Or run directly:"
echo "    cd $REPO_DIR"
echo "    poetry run python audit.py <SHEET_ID>"
echo ""
