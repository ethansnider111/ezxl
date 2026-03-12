# EZXL

**Your Excel helper — formulas, shortcuts, and quick actions all in one place.**

EZXL is a Microsoft Excel task pane add-in that puts 80+ formula builders, keyboard shortcuts, and one-click actions directly inside Excel so you never have to leave your spreadsheet.

---

## Features

**⚡ Formula Builder**
Build complex Excel formulas without memorizing syntax. Pick a formula, fill in the inputs, and insert it directly into your selected cell. Supports chaining formulas together as inputs to other formulas.

**⇄ Lookup & Text**
VLOOKUP, XLOOKUP, INDEX/MATCH, text manipulation, date functions, and more — all with guided inputs and plain-language descriptions.

**⌨ Shortcuts**
A full reference of Excel keyboard shortcuts, organized and searchable. Expandable cards with detailed descriptions.

**⚡ GO TO Panel**
Jump to A1, the last used cell, or select all data in one tap. Type any cell address to jump there instantly.

**🛠 Tools Panel**
Toggle AutoFilter, freeze/unfreeze rows, create tables, and more — one tap, no ribbon digging.

**🎨 Themes**
9 themes across light, dark, and monochrome. Auto-detects Excel's dark/light mode on launch and picks the matching green theme.

**Sort**
Sort formulas and shortcuts by default order, most used, or alphabetically — per tab, with direction toggle.

---

## Installation

1. Download `manifest.xml` from this repository and save it to a folder on your computer (e.g. `C:\Users\you\Desktop\EZXL`)
2. In Excel go to **File → Options → Trust Center → Trust Center Settings → Trusted Add-in Catalogs**
3. In the catalog address box enter the path to your folder (e.g. `\\localhost\C$\Users\you\Desktop\EZXL`)
4. Check **Show in Menu** and click OK
5. Restart Excel
6. Go to **Insert → Add-ins → Shared Folder** — EZXL will appear

---

## Files

| File | Description |
|------|-------------|
| `index.html` | The full add-in (single file) |
| `manifest.xml` | Office add-in manifest for Excel |
| `icon-16.png` | Ribbon icon 16×16 |
| `icon-32.png` | Ribbon icon 32×32 |
| `icon-80.png` | Ribbon icon 80×80 |

---

## Development

Everything lives in `index.html` — no build step, no dependencies, no framework. Just open it in a browser to preview the UI, or sideload via the manifest to run inside Excel with full Office JS API access.

---

## Tech Stack

- Vanilla HTML/CSS/JS
- [Office JS API](https://learn.microsoft.com/en-us/office/dev/add-ins/reference/overview/excel-add-ins-reference-overview) for Excel integration
- [JetBrains Mono](https://www.jetbrains.com/legalnotices/terms_of_use_for_jetbrains_mono_typeface/) + [Bebas Neue](https://fonts.google.com/specimen/Bebas+Neue) fonts
- Hosted on GitHub Pages
