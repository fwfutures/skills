#!/usr/bin/env bash
set -euo pipefail

# Install skills from the public skills repo into user-level directories.
#
# Default (global) mode:
#   - Symlinks public skills into ~/.agents/skills/
#   - Creates per-skill links in ~/.claude/skills/ and ~/.codex/skills/
#
# Project mode (--project):
#   - Copies all public skills into ./.claude/skills/ for portable use
#
# Usage:
#   ./scripts/install-skills.sh [--clean] [--project]

CLEAN=0
PROJECT=0
for arg in "$@"; do
  case "$arg" in
    --clean) CLEAN=1 ;;
    --project) PROJECT=1 ;;
  esac
done

ROOT_DIR="$(cd "$(dirname "${BASH_SOURCE[0]}")/.." && pwd)"
PUBLIC_SKILLS_DIR="$ROOT_DIR/plugins/fwf-public-skills/skills"

if [ ! -d "$PUBLIC_SKILLS_DIR" ]; then
  echo "ERROR: $PUBLIC_SKILLS_DIR not found."
  exit 1
fi

ensure_dir() {
  local dir="$1"

  if [ -L "$dir" ]; then
    local current
    current="$(readlink "$dir")"
    echo "NOTICE: $dir is a symlink -> $current"
    echo "        Replacing with a real directory for per-skill links."
    rm -f "$dir"
  fi

  if [ -e "$dir" ] && [ ! -d "$dir" ]; then
    echo "WARNING: $dir exists and is not a directory. Skipping."
    return 1
  fi

  mkdir -p "$dir"
}

clean_dir() {
  local dir="$1"
  local skip="${2:-}"

  if [ -L "$dir" ]; then
    rm -f "$dir"
  fi

  mkdir -p "$dir"

  if [ -n "$skip" ]; then
    find "$dir" -mindepth 1 -maxdepth 1 ! -name "$skip" -exec rm -rf {} +
  else
    find "$dir" -mindepth 1 -maxdepth 1 -exec rm -rf {} +
  fi
}

link_skill() {
  local target="$1"
  local source="$2"

  if [ -L "$target" ]; then
    local current
    current="$(readlink "$target")"
    if [ "$current" = "$source" ]; then
      return 0
    fi

    local current_resolved source_resolved
    current_resolved="$(python3 -c "import os,sys; print(os.path.realpath(sys.argv[1]))" "$target" 2>/dev/null || echo "$current")"
    source_resolved="$(python3 -c "import os,sys; print(os.path.realpath(sys.argv[1]))" "$source" 2>/dev/null || echo "$source")"
    if [ "$current_resolved" = "$source_resolved" ]; then
      rm -f "$target"
      ln -s "$source" "$target"
      return 0
    fi

    echo "WARNING: $target points to $current (expected $source). Replacing."
    rm -f "$target"
  fi

  if [ -e "$target" ]; then
    if [ -d "$target" ]; then
      echo "NOTICE: replacing existing directory at $target with symlink -> $source"
      rm -rf "$target"
    else
      echo "WARNING: $target exists and is not a directory or symlink. Skipping."
      return 0
    fi
  fi

  ln -s "$source" "$target"
}

if [ "$PROJECT" = "1" ]; then
  PROJECT_DIR="$(pwd)"
  PROJECT_REAL="$(cd "$PROJECT_DIR" && pwd -P)"
  ROOT_REAL="$(cd "$ROOT_DIR" && pwd -P)"

  case "$PROJECT_REAL" in
    "$ROOT_REAL"|"$ROOT_REAL"/*)
      echo "ERROR: --project must be run from outside the public skills repo."
      echo "       cd to your project directory first, then run:"
      echo "       $ROOT_DIR/scripts/install-skills.sh --project"
      exit 1
      ;;
  esac

  PROJECT_SKILLS_DIR="$PROJECT_DIR/.claude/skills"

  if [ "$CLEAN" = "1" ]; then
    echo "Cleaning $PROJECT_SKILLS_DIR ..."
    rm -rf "$PROJECT_SKILLS_DIR"
  fi

  mkdir -p "$PROJECT_SKILLS_DIR"

  echo "Copying public skills into $PROJECT_SKILLS_DIR ..."
  copied=0
  for skill_dir in "$PUBLIC_SKILLS_DIR"/*; do
    if [ ! -d "$skill_dir" ]; then
      continue
    fi
    skill_name="$(basename "$skill_dir")"
    target="$PROJECT_SKILLS_DIR/$skill_name"
    if [ -e "$target" ] && [ "$CLEAN" != "1" ]; then
      echo "  ~ $skill_name (exists, skipping)"
    else
      cp -R "$skill_dir" "$target"
      echo "  + $skill_name"
      copied=$((copied + 1))
    fi
  done

  echo "  ($copied public skills copied)"
  echo ""
  echo "Done. Skills installed into: $PROJECT_SKILLS_DIR"
  exit 0
fi

if [ "$CLEAN" = "1" ]; then
  echo "Cleaning skill directories..."
  clean_dir "$HOME/.agents/skills"
  clean_dir "$HOME/.claude/skills"
  if [ "${CLEAN_CODEX_SYSTEM:-}" = "1" ]; then
    clean_dir "$HOME/.codex/skills"
  else
    clean_dir "$HOME/.codex/skills" ".system"
  fi
  clean_dir "$HOME/.picoclaw/skills"
fi

ensure_dir "$HOME/.agents/skills"
ensure_dir "$HOME/.claude/skills"
ensure_dir "$HOME/.codex/skills"
ensure_dir "$HOME/.picoclaw/skills"

echo ""
echo "Linking public skills..."
public_count=0
for skill_dir in "$PUBLIC_SKILLS_DIR"/*; do
  if [ ! -d "$skill_dir" ]; then
    continue
  fi
  skill_name="$(basename "$skill_dir")"
  link_skill "$HOME/.agents/skills/$skill_name" "$skill_dir"
  public_count=$((public_count + 1))
done
echo "  ($public_count public skills)"

echo "Linking into agent directories..."
for skill_dir in "$HOME/.agents/skills"/*; do
  if [ ! -d "$skill_dir" ] && [ ! -L "$skill_dir" ]; then
    continue
  fi
  skill_name="$(basename "$skill_dir")"
  link_skill "$HOME/.claude/skills/$skill_name" "$skill_dir"
  link_skill "$HOME/.codex/skills/$skill_name" "$skill_dir"
done

echo "Linking into picoclaw directory..."
for skill_dir in "$HOME/.agents/skills"/*; do
  if [ ! -d "$skill_dir" ] && [ ! -L "$skill_dir" ]; then
    continue
  fi
  skill_name="$(basename "$skill_dir")"
  link_skill "$HOME/.picoclaw/skills/$skill_name" "$skill_dir"
done

echo ""
echo "Done. Public skills installed."
