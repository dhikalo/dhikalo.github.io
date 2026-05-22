# Messstellen Manager Pro v6.0

Professional field measurement tool for soil resistance (Bodenwiderstandsmessung).

## Features

- **Data Table** — Full measurement grid with live ρ/MW/SD calculations
- **Interactive Map** — GPS tracking, markers, lines, numbered pins, offline plans
- **Excel Export** — Professional reports with images and statistics
- **Offline PWA** — Works without internet, auto-caches all assets
- **Undo/Redo** — Full history stack for data changes
- **Multi-Project** — Save and switch between measurement sites

## Quick Start

1. Open `index.html` in any modern browser
2. Choose "Datei öffnen" (import Excel) or "Neues Projekt"
3. Start measuring

## Tech Stack

- Pure HTML/CSS/JS (no build tools needed)
- Leaflet for maps
- ExcelJS + SheetJS for Excel I/O
- html2canvas for map screenshots
- Service Worker for offline capability

## File Structure

```
messstellen-manager/
├── index.html       ← Main entry point
├── app.js           ← Application logic (v6.0)
├── formatting.js    ← Cell formatting module
├── styles.css       ← Professional design system
├── sw.js            ← Service worker (offline)
├── manifest.json    ← PWA manifest
├── app-icon.png     ← App icon
└── idb_logo.jpg     ← Company logo
```

## What's New in v6.0

- Complete rewrite — zero known bugs
- Professional, clean UI design (sell-ready)
- Proper undo/redo implementation
- Fixed formatting module (was broken in v5.x)
- Touch-optimized for tablets (44px touch targets)
- Better GPS error handling
- Improved offline caching strategy
- Cleaner code structure for maintainability

## Browser Support

- Chrome/Edge 90+
- Safari 15+ (iOS/macOS)
- Firefox 90+

---

NRW-IDB | Messtechnik & Bauservice
