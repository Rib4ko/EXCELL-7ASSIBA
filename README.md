# EXCELL-7ASSIBA

A high-performance Excel-compatible desktop spreadsheet application built with vanilla JavaScript, HTML5 Canvas 2D virtualization, dynamic DAG recalculation engine, and modern dark-mode ribbon interface.

## Features
- **High-DPI Virtualized Canvas Grid**: Ultra-smooth scrolling, column/row resizing, auto-fit, selection range, and drag-to-fill handles.
- **DAG Reactive Formula Engine**: Full dependency resolution and evaluation for formulas (`SUM`, `AVERAGE`, `COUNT`, `MIN`, `MAX`, `IF`, `VLOOKUP`, `HLOOKUP`, `INDEX`, `MATCH`, `XLOOKUP`).
- **Interactive Formula Picking**: Click or Ctrl-click cells to append or replace coordinates directly into formula bar.
- **Modern Dark Ribbon Interface**: Pixel-accurate desktop ribbon UI with Quick Access toolbar, tabs (`File`, `Home`, `Insert`, etc.), typography controls, alignment, borders, number formatting, and cell style chips.
- **Office Compatibility**: Native `.xlsx` read/write via ExcelJS, CSV import/export, JSON backup.
- **Full Keyboard Shortcuts Suite**: `Ctrl+C`, `Ctrl+V`, `Ctrl+X`, `Ctrl+Z` (Undo), `Ctrl+Y` (Redo), `Ctrl+B`, `Ctrl+I`, `Ctrl+U`, `Ctrl+A`, `Ctrl+S`, `F2`.
- **Standalone Desktop Launcher**: Run as a dedicated desktop application with dedicated profile via `launch-app.bat`.

## Quick Start
```bash
npm install
npm start
```
Or launch standalone desktop window:
```bat
launch-app.bat
```
