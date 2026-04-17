#!/usr/bin/env bash
# setup.sh — one-command onboarding for all Picnic Claude skills
# Run once after cloning: ./setup.sh

set -euo pipefail

REPO_DIR="$(cd "$(dirname "${BASH_SOURCE[0]}")" && pwd)"
SKILLS_DIR="$HOME/.claude/skills"
CREDS_PATH="$HOME/.cache/picnic-google-sheets/authorized_user.json"

echo "━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━"
echo "  Picnic Claude Skills — setup"
echo "━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━"

# ── 1. Poetry ─────────────────────────────────────────
if ! command -v poetry &>/dev/null; then
    echo ""
    echo "ERROR: Poetry is not installed."
    echo "Install from: https://python-poetry.org/docs/#installation"
    exit 1
fi
echo ""
echo "✅ Poetry: $(poetry --version)"

# ── 2. Install + symlink each skill ───────────────────
# A skill is any subfolder containing a SKILL.md file.
# If it also has a pyproject.toml, poetry install is run inside it.

mkdir -p "$SKILLS_DIR"
SKILL_COUNT=0

for skill_dir in "$REPO_DIR"/*/; do
    [[ -f "$skill_dir/SKILL.md" ]] || continue

    skill_name="$(basename "$skill_dir")"
    echo ""
    echo "── $skill_name ──────────────────────────────────"

    if [[ -f "$skill_dir/pyproject.toml" ]]; then
        echo "  Installing dependencies..."
        (cd "$skill_dir" && poetry install)
        echo "  ✅ Dependencies ready."
    fi

    link="$SKILLS_DIR/$skill_name"
    if [[ -L "$link" ]]; then
        current="$(readlink "$link")"
        if [[ "$current" == "$skill_dir" || "$current" == "${skill_dir%/}" ]]; then
            echo "  ✅ Symlink already correct."
        else
            ln -sfn "$skill_dir" "$link"
            echo "  ✅ Symlink updated (was → $current)."
        fi
    elif [[ -e "$link" ]]; then
        echo "  ERROR: $link exists but is not a symlink. Remove it manually and re-run."
        exit 1
    else
        ln -s "$skill_dir" "$link"
        echo "  ✅ Symlink created: $link"
    fi

    SKILL_COUNT=$((SKILL_COUNT + 1))
done

# ── 3. Google Sheets credentials ──────────────────────
echo ""
echo "── Google Sheets credentials ────────────────────"
if [[ -f "$CREDS_PATH" ]]; then
    echo "  ✅ Credentials found."
else
    echo "  ⚠️  Not found at: $CREDS_PATH"
    echo "     Run the one-time auth flow:"
    echo "     cd ~/.claude/skills/picnic-gsheet && poetry run python gsheet_auth_setup.py"
fi

# ── Done ──────────────────────────────────────────────
echo ""
echo "━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━"
echo "  Setup complete — $SKILL_COUNT skill(s) installed."
echo "━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━"
echo ""
echo "  Use any skill in Claude Code: /<skill-name> <args>"
echo ""
