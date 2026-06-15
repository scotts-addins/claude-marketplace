# Scott's Add-ins — Claude Plugins

This repository is a **marketplace** containing Claude plugins for users of [Scott's Add-ins](https://scottsaddins.com). Install one of the plugins below in Cowork or Claude Code, then talk to Claude in plain English to build Excel workbooks that pull live data from your Xero or QuickBooks org.

> **This repo is the single source of truth for the skill files.** All edits to `SKILL.md`, references (`pnl.md`, `balance-sheet.md`, `formula-reference.md`, etc.), and `scripts/build_pattern.py` should be made here, under `plugins/scotts-addins/skills/`. Any other copies elsewhere on disk are stale and should not be edited.

## Plugins in this marketplace

| Plugin | What it does |
|--------|--------------|
| [`scotts-addins`](./plugins/scotts-addins) | Builds Excel workbooks (P&L, Balance Sheet, Budget vs. Actual, transaction registers, sales reports, custom layouts) populated with `=SCOTT.*` formulas. Includes a companion skill that answers formula-syntax questions. |

## Installation

### In Cowork

1. Open Cowork's plugin settings.
2. Choose **Add custom marketplace**.
3. Paste this repo's URL: `https://github.com/scotts-addins/claude-marketplace`.
4. Install `scotts-addins`.

### In Claude Code

```
/plugin marketplace add scotts-addins/claude-marketplace
/plugin install scotts-addins
```

## What you get after installing

Two skills become available to Claude:

- **`scott-workbook-builder`** — triggers when you ask Claude to build, create, or generate an Excel workbook using Scott's Add-ins.
- **`scott-formulas`** — triggers when you ask Claude formula-syntax questions ("how do I write `SCOTT.GL`?", "what arguments does `SCOTT.BUDGET` take?", etc.).

You don't trigger them by name — Claude routes to the right one based on what you ask.

## Requirements

- Cowork or Claude Code installed and signed in.
- Excel with [Scott's Add-ins](https://scottsaddins.com) installed and authorized for at least one Xero or QuickBooks org.

## Updates

We push updates as we improve the skills. Run `/plugin update scotts-addins` (Claude Code) or use Cowork's plugin sync to pick them up.

## Support

Issues, suggestions, and bug reports: file an issue on this repo, or contact scott@scottsaddins.com.
