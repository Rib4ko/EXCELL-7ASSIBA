import { Workbook, CellChangeAction } from '../core/workbook.js';
import { CanvasGridRenderer } from '../render/gridRenderer.js';
import { formatCoordinate, colToName, shiftFormula } from '../core/cell.js';

export class SpreadsheetApp {
    constructor() {
        this.workbook = new Workbook('My Workbook');
        this.canvas = document.getElementById('grid-canvas');
        this.renderer = new CanvasGridRenderer(this.canvas, this.workbook);

        // UI elements
        this.cellAddressBox = document.getElementById('cell-address');
        this.formulaInput = document.getElementById('formula-input');
        this.inCellEditor = document.getElementById('in-cell-editor');
        this.sheetTabsContainer = document.getElementById('sheet-tabs');
        this.btnAddSheet = document.getElementById('btn-add-sheet');

        // Filter popup element
        this.filterPopup = document.getElementById('filter-popup');

        this.populateSampleData();
        this.initEventListeners();
        this.renderSheetTabs();
        this.updateActiveCellDisplay();
        this.renderer.render();
    }

    populateSampleData() {
        this.workbook.title = 'L7ASSIBA - Business_Financial_Model.xlsx';
        const sheet = this.workbook.getActiveSheet();
        sheet.name = 'Financial Overview';
        
        // --- 1. Main Executive Header ---
        sheet.setCellValue(0, 0, '📊 L7ASSIBA — ENTERPRISE FINANCIAL & OPERATIONAL INTELLIGENCE');
        sheet.setCellStyle(0, 0, { bold: true, fontSize: 16, color: '#15803d', align: 'left' });
        sheet.mergeCells(0, 0, 0, 7);

        sheet.setCellValue(1, 0, 'Fiscal Year 2026 Consolidated P&L Analysis | Live Formula Calculation & Matrix Engine');
        sheet.setCellStyle(1, 0, { italic: true, fontSize: 11, color: '#64748b', align: 'left' });
        sheet.mergeCells(1, 0, 1, 7);

        // --- 2. Primary Revenue & Performance Table ---
        const headers = ['Business Unit', 'Q1 Revenue', 'Q2 Revenue', 'Q3 Revenue', 'Q4 Revenue', 'Annual Total', 'Target', 'KPI Status'];
        const colWidths = [170, 120, 120, 120, 120, 130, 120, 110];
        headers.forEach((h, c) => {
            sheet.setCellValue(3, c, h);
            sheet.setCellStyle(3, c, { bold: true, background: '#1e293b', color: '#ffffff', align: c === 0 ? 'left' : (c === 7 ? 'center' : 'right') });
            sheet.setColWidth(c, colWidths[c]);
        });

        const unitsData = [
            ['Cloud Solutions', 125000, 142000, 158000, 185000, 580000],
            ['Hardware Systems', 95000, 89000, 102000, 115000, 420000],
            ['AI & Automation', 68000, 84000, 115000, 148000, 390000],
            ['Security Services', 54000, 62000, 71000, 89000, 260000],
            ['Digital Media', 41000, 45000, 48000, 52000, 190000],
            ['Consulting & Training', 32000, 37000, 42000, 48000, 150000]
        ];

        unitsData.forEach((row, idx) => {
            const r = 4 + idx;
            sheet.setCellValue(r, 0, row[0]);
            sheet.setCellStyle(r, 0, { bold: true, color: '#1e293b' });

            // Q1 to Q4 values
            for (let c = 1; c <= 4; c++) {
                sheet.setCellValue(r, c, row[c]);
                sheet.setCellStyle(r, c, { numFormat: 'currency', align: 'right' });
            }

            // Column 5: Annual Total Formula =SUM(B:E)
            sheet.setCellValue(r, 5, `=SUM(B${r + 1}:E${r + 1})`);
            sheet.setCellStyle(r, 5, { bold: true, numFormat: 'currency', align: 'right', background: '#f1f5f9' });

            // Column 6: Target
            sheet.setCellValue(r, 6, row[5]);
            sheet.setCellStyle(r, 6, { numFormat: 'currency', align: 'right' });

            // Column 7: KPI Status Formula =IF(F >= G, "EXCEEDED", "BEHIND")
            sheet.setCellValue(r, 7, `=IF(F${r + 1}>=G${r + 1}, "EXCEEDED", "BEHIND")`);
            sheet.setCellStyle(r, 7, { bold: true, align: 'center', color: '#15803d', background: '#ecfdf5' });
        });

        // Total Summary Row
        const totalR = 4 + unitsData.length;
        sheet.setCellValue(totalR, 0, 'TOTAL CONSOLIDATED');
        sheet.setCellStyle(totalR, 0, { bold: true, background: '#e2e8f0', color: '#0f172a' });

        for (let c = 1; c <= 6; c++) {
            const colL = colToName(c);
            sheet.setCellValue(totalR, c, `=SUM(${colL}5:${colL}${totalR})`);
            sheet.setCellStyle(totalR, c, { bold: true, background: '#e2e8f0', color: '#0f172a', numFormat: 'currency', align: 'right' });
        }
        sheet.setCellValue(totalR, 7, '100% AUDITED');
        sheet.setCellStyle(totalR, 7, { bold: true, background: '#e2e8f0', color: '#0f172a', align: 'center' });

        // Average Row
        const avgR = totalR + 1;
        sheet.setCellValue(avgR, 0, 'AVERAGE PER DIVISION');
        sheet.setCellStyle(avgR, 0, { bold: true, italic: true, background: '#f8fafc' });
        for (let c = 1; c <= 6; c++) {
            const colL = colToName(c);
            sheet.setCellValue(avgR, c, `=AVERAGE(${colL}5:${colL}${totalR})`);
            sheet.setCellStyle(avgR, c, { italic: true, background: '#f8fafc', numFormat: 'currency', align: 'right' });
        }
        sheet.setCellValue(avgR, 7, '-');
        sheet.setCellStyle(avgR, 7, { align: 'center' });

        // Enable AutoFilter on the primary table headers
        sheet.toggleAutoFilter(3, 0, totalR - 1, 7);

        // --- 3. Dynamic Interactive Lookup Dashboard Card ---
        const cardTop = avgR + 3;
        sheet.setCellValue(cardTop, 0, '⚡ EXECUTIVE INTELLIGENCE SEARCH (VLOOKUP & XLOOKUP)');
        sheet.setCellStyle(cardTop, 0, { bold: true, fontSize: 13, color: '#0369a1' });
        sheet.mergeCells(cardTop, 0, cardTop, 5);

        sheet.setCellValue(cardTop + 1, 0, 'Select or type Division name in B' + (cardTop + 2) + ' to test real-time recalculation:');
        sheet.setCellStyle(cardTop + 1, 0, { italic: true, color: '#64748b' });
        sheet.mergeCells(cardTop + 1, 0, cardTop + 1, 5);

        sheet.setCellValue(cardTop + 2, 0, 'Query Unit:');
        sheet.setCellStyle(cardTop + 2, 0, { bold: true, background: '#f1f5f9' });
        sheet.setCellValue(cardTop + 2, 1, 'AI & Automation');
        sheet.setCellStyle(cardTop + 2, 1, { bold: true, color: '#0284c7', background: '#e0f2fe' });

        sheet.setCellValue(cardTop + 3, 0, 'VLOOKUP Total:');
        sheet.setCellStyle(cardTop + 3, 0, { bold: true });
        sheet.setCellValue(cardTop + 3, 1, `=VLOOKUP(B${cardTop + 3}, A5:H10, 6, FALSE)`);
        sheet.setCellStyle(cardTop + 3, 1, { bold: true, numFormat: 'currency', color: '#166534', background: '#dcfce7' });

        sheet.setCellValue(cardTop + 4, 0, 'XLOOKUP Target:');
        sheet.setCellStyle(cardTop + 4, 0, { bold: true });
        sheet.setCellValue(cardTop + 4, 1, `=XLOOKUP(B${cardTop + 3}, A5:A10, G5:G10)`);
        sheet.setCellStyle(cardTop + 4, 1, { bold: true, numFormat: 'currency' });

        sheet.setCellValue(cardTop + 5, 0, 'INDEX / MATCH Status:');
        sheet.setCellStyle(cardTop + 5, 0, { bold: true });
        sheet.setCellValue(cardTop + 5, 1, `=INDEX(H5:H10, MATCH(B${cardTop + 3}, A5:A10, 0), 1)`);
        sheet.setCellStyle(cardTop + 5, 1, { bold: true, color: '#15803d', align: 'center', background: '#f0fdf4' });

        // --- 4. Populate Sheet 2: Inventory & Catalog ---
        const sheet2 = this.workbook.addSheet('Inventory & Logistics');
        sheet2.setCellValue(0, 0, '📦 WAREHOUSE INVENTORY & STOCK MANAGEMENT');
        sheet2.setCellStyle(0, 0, { bold: true, fontSize: 15, color: '#0284c7' });
        sheet2.mergeCells(0, 0, 0, 5);

        const invHeaders = ['SKU Code', 'Product Description', 'Unit Price', 'In Stock', 'Inventory Value', 'Reorder Alert'];
        const invColWidths = [110, 210, 110, 95, 140, 120];
        invHeaders.forEach((h, c) => {
            sheet2.setCellValue(2, c, h);
            sheet2.setCellStyle(2, c, { bold: true, background: '#0f172a', color: '#ffffff', align: c < 2 ? 'left' : (c === 5 ? 'center' : 'right') });
            sheet2.setColWidth(c, invColWidths[c]);
        });

        const inventory = [
            ['SKU-1001', 'Enterprise Rack Server 2U', 2499.00, 45],
            ['SKU-1002', 'High-Density Switch 48-Port', 850.50, 120],
            ['SKU-1003', 'NVMe SSD Storage Array 16TB', 1290.00, 18],
            ['SKU-1004', 'Edge AI Acceleration Card', 640.00, 85],
            ['SKU-1005', 'Fiber Optic Transceiver 100G', 175.25, 340],
            ['SKU-1006', 'Redundant Power Supply 1200W', 215.00, 12]
        ];

        inventory.forEach((item, i) => {
            const r = 3 + i;
            sheet2.setCellValue(r, 0, item[0]);
            sheet2.setCellValue(r, 1, item[1]);
            sheet2.setCellValue(r, 2, item[2]);
            sheet2.setCellStyle(r, 2, { numFormat: 'currency', align: 'right' });
            sheet2.setCellValue(r, 3, item[3]);
            sheet2.setCellStyle(r, 3, { align: 'right' });

            // Formula: Value = Unit Price * In Stock
            sheet2.setCellValue(r, 4, `=C${r + 1} * D${r + 1}`);
            sheet2.setCellStyle(r, 4, { bold: true, numFormat: 'currency', align: 'right' });

            // Formula: Alert = IF(In Stock < 25, "LOW STOCK", "OK")
            sheet2.setCellValue(r, 5, `=IF(D${r + 1} < 25, "REORDER", "IN STOCK")`);
            const isLow = item[3] < 25;
            sheet2.setCellStyle(r, 5, { bold: true, align: 'center', color: isLow ? '#b91c1c' : '#15803d', background: isLow ? '#fef2f2' : '#f0fdf4' });
        });

        sheet2.setCellValue(10, 1, 'TOTAL INVENTORY VALUE:');
        sheet2.setCellStyle(10, 1, { bold: true, align: 'right' });
        sheet2.setCellValue(10, 4, '=SUM(E4:E9)');
        sheet2.setCellStyle(10, 4, { bold: true, fontSize: 13, numFormat: 'currency', background: '#e2e8f0', align: 'right' });

        // Switch back to Sheet 1 as active
        this.workbook.setActiveSheet(0);
    }

    initEventListeners() {
        window.addEventListener('resize', () => this.renderer.onResize());

        this.canvas.addEventListener('mousedown', (e) => this.handleMouseDown(e));
        window.addEventListener('mousemove', (e) => this.handleMouseMove(e));
        window.addEventListener('mouseup', (e) => this.handleMouseUp(e));

        this.canvas.addEventListener('dblclick', (e) => this.handleDoubleClick(e));
        this.canvas.addEventListener('wheel', (e) => this.handleWheel(e), { passive: false });
        window.addEventListener('keydown', (e) => this.handleKeyDown(e));

        // Clipboard paste event listener (supports system clipboard paste)
        window.addEventListener('paste', (e) => {
            if (this.isEditingInCell()) return;
            if (document.activeElement === this.formulaInput) return;
            e.preventDefault();
            const text = e.clipboardData?.getData('text/plain') || '';
            if (text) {
                this.pasteData(text);
            }
        });

        // Click outside closes filter popup
        window.addEventListener('click', (e) => {
            if (this.filterPopup && !this.filterPopup.contains(e.target) && !e.target.closest('#grid-canvas')) {
                this.closeFilterPopup();
            }
        });

        this.formulaInput.addEventListener('input', () => {
            if (this.isEditingInCell()) return;
            const r = this.renderer.selection.activeRow;
            const c = this.renderer.selection.activeCol;
            this.workbook.getActiveSheet().setCellValue(r, c, this.formulaInput.value);
            this.renderer.render();
        });

        this.formulaInput.addEventListener('keydown', (e) => {
            if (e.key === 'Enter') {
                e.preventDefault();
                this.finishEditing();
                this.moveSelection(1, 0);
            }
        });

        this.inCellEditor.addEventListener('input', () => {
            this.formulaInput.value = this.inCellEditor.value;
        });

        this.inCellEditor.addEventListener('keydown', (e) => {
            if (e.key === 'Enter') {
                e.preventDefault();
                this.finishInCellEdit();
                this.moveSelection(1, 0);
            } else if (e.key === 'Escape') {
                e.preventDefault();
                this.cancelInCellEdit();
            } else if (e.key === 'Tab') {
                e.preventDefault();
                this.finishInCellEdit();
                this.moveSelection(0, 1);
            }
        });

        this.btnAddSheet.addEventListener('click', () => {
            this.workbook.addSheet();
            this.workbook.setActiveSheet(this.workbook.sheets.length - 1);
            this.renderSheetTabs();
            this.renderer.render();
        });

        this.setupToolbarActions();
    }

    setupToolbarActions() {
        const toggleStyle = (prop, valueGetter) => {
            const sel = this.renderer.getSelectionRange();
            const sheet = this.workbook.getActiveSheet();
            const changes = [];

            for (let r = sel.minRow; r <= sel.maxRow; r++) {
                for (let c = sel.minCol; c <= sel.maxCol; c++) {
                    const cell = sheet.getOrCreateCell(r, c);
                    const oldStyle = { ...cell.style };
                    const current = cell.style[prop];
                    const next = valueGetter ? valueGetter(current) : !current;
                    sheet.setCellStyle(r, c, { [prop]: next });
                    changes.push({
                        row: r,
                        col: c,
                        oldVal: cell.value,
                        oldFormula: cell.formula,
                        oldStyle,
                        newVal: cell.value,
                        newFormula: cell.formula,
                        newStyle: { ...cell.style }
                    });
                }
            }
            if (changes.length > 0) {
                this.workbook.pushUndo(new CellChangeAction(sheet, changes));
            }
            this.renderer.render();
            this.updateToolbarState();
        };
        this.toggleStyle = toggleStyle;

        // Clipboard & History Actions
        document.getElementById('btn-undo')?.addEventListener('click', () => this.undo());
        document.getElementById('btn-redo')?.addEventListener('click', () => this.redo());
        document.getElementById('btn-cut')?.addEventListener('click', () => this.cutSelectionToClipboard());
        document.getElementById('btn-copy')?.addEventListener('click', () => this.copySelectionToClipboard());
        document.getElementById('btn-paste')?.addEventListener('click', () => this.pasteFromClipboard());

        document.getElementById('btn-bold')?.addEventListener('click', () => toggleStyle('bold'));
        document.getElementById('btn-italic')?.addEventListener('click', () => toggleStyle('italic'));
        document.getElementById('btn-underline')?.addEventListener('click', () => toggleStyle('underline'));

        document.getElementById('btn-align-left')?.addEventListener('click', () => toggleStyle('align', () => 'left'));
        document.getElementById('btn-align-center')?.addEventListener('click', () => toggleStyle('align', () => 'center'));
        document.getElementById('btn-align-right')?.addEventListener('click', () => toggleStyle('align', () => 'right'));

        document.getElementById('font-color-picker')?.addEventListener('input', (e) => toggleStyle('color', () => e.target.value));
        document.getElementById('bg-color-picker')?.addEventListener('input', (e) => toggleStyle('background', () => e.target.value));

        document.getElementById('btn-format-currency')?.addEventListener('click', () => toggleStyle('numFormat', () => 'currency'));
        document.getElementById('btn-format-percent')?.addEventListener('click', () => toggleStyle('numFormat', () => 'percent'));
        document.getElementById('btn-format-number')?.addEventListener('click', () => toggleStyle('numFormat', () => 'general'));

        document.getElementById('btn-formula-sum')?.addEventListener('click', () => {
            const sel = this.renderer.getSelectionRange();
            const r = sel.maxRow + 1;
            const c = sel.activeCol;
            const colLetter = colToName(c);
            const formula = `=SUM(${colLetter}${sel.minRow + 1}:${colLetter}${sel.maxRow + 1})`;
            this.workbook.getActiveSheet().setCellValue(r, c, formula);
            this.renderer.selection.activeRow = r;
            this.renderer.render();
            this.updateActiveCellDisplay();
        });

        document.getElementById('btn-formula-avg')?.addEventListener('click', () => {
            const sel = this.renderer.getSelectionRange();
            const r = sel.maxRow + 1;
            const c = sel.activeCol;
            const colLetter = colToName(c);
            const formula = `=AVERAGE(${colLetter}${sel.minRow + 1}:${colLetter}${sel.maxRow + 1})`;
            this.workbook.getActiveSheet().setCellValue(r, c, formula);
            this.renderer.selection.activeRow = r;
            this.renderer.render();
            this.updateActiveCellDisplay();
        });

        document.getElementById('btn-clear')?.addEventListener('click', () => this.clearSelectionValues());

        document.getElementById('btn-merge-center')?.addEventListener('click', () => {
            const sel = this.renderer.getSelectionRange();
            const sheet = this.workbook.getActiveSheet();
            const merge = sheet.getMerge(sel.minRow, sel.minCol);
            if (merge) {
                sheet.unmergeCells(sel.minRow, sel.minCol);
            } else {
                sheet.mergeCells(sel.minRow, sel.minCol, sel.maxRow, sel.maxCol);
            }
            this.renderer.render();
            this.updateToolbarState();
        });

        // Phase 2: Sort Ascending
        document.getElementById('btn-sort-asc')?.addEventListener('click', () => {
            const sel = this.renderer.getSelectionRange();
            this.workbook.getActiveSheet().sortRange(sel.minRow, sel.maxRow, sel.minCol, sel.maxCol, sel.minCol, true);
            this.renderer.render();
            this.updateActiveCellDisplay();
        });

        // Phase 2: Sort Descending
        document.getElementById('btn-sort-desc')?.addEventListener('click', () => {
            const sel = this.renderer.getSelectionRange();
            this.workbook.getActiveSheet().sortRange(sel.minRow, sel.maxRow, sel.minCol, sel.maxCol, sel.minCol, false);
            this.renderer.render();
            this.updateActiveCellDisplay();
        });

        // Phase 2: AutoFilter Toggle
        document.getElementById('btn-filter')?.addEventListener('click', () => {
            const sel = this.renderer.getSelectionRange();
            this.workbook.getActiveSheet().toggleAutoFilter(sel.minRow, sel.minCol, sel.maxRow, sel.maxCol);
            const isFilterActive = !!this.workbook.getActiveSheet().autoFilter;
            document.getElementById('btn-filter')?.classList.toggle('active', isFilterActive);
            this.renderer.render();
        });

        // Phase 2: Freeze Panes Toggle
        document.getElementById('btn-freeze-pane')?.addEventListener('click', () => {
            const sheet = this.workbook.getActiveSheet();
            if (sheet.freezeRow > 0) {
                sheet.freezeRow = 0;
            } else {
                const sel = this.renderer.getSelectionRange();
                sheet.freezeRow = sel.minRow > 0 ? sel.minRow : 1;
            }
            const isFrozen = sheet.freezeRow > 0;
            document.getElementById('btn-freeze-pane')?.classList.toggle('active', isFrozen);
            this.renderer.render();
        });

        // Quick Access bar
        document.getElementById('btn-quick-save')?.addEventListener('click', () => this.exportXLSX());
        document.getElementById('btn-quick-print')?.addEventListener('click', () => window.print());
        document.getElementById('btn-undo-qa')?.addEventListener('click', () => this.undo());
        document.getElementById('btn-redo-qa')?.addEventListener('click', () => this.redo());

        // Ribbon tabs navigation
        document.querySelectorAll('.ribbon-tab-btn').forEach(tabBtn => {
            tabBtn.addEventListener('click', () => {
                document.querySelectorAll('.ribbon-tab-btn').forEach(b => b.classList.remove('active'));
                tabBtn.classList.add('active');
                if (tabBtn.dataset.tab === 'file') {
                    document.getElementById('file-input')?.click();
                }
            });
        });

        // Font family & size
        document.getElementById('font-family-select')?.addEventListener('change', (e) => toggleStyle('fontFamily', () => e.target.value));
        document.getElementById('font-size-select')?.addEventListener('change', (e) => toggleStyle('fontSize', () => parseInt(e.target.value, 10)));
        document.getElementById('btn-font-increase')?.addEventListener('click', () => toggleStyle('fontSize', (s = 11) => Math.min(72, (parseInt(s, 10) || 11) + 1)));
        document.getElementById('btn-font-decrease')?.addEventListener('click', () => toggleStyle('fontSize', (s = 11) => Math.max(6, (parseInt(s, 10) || 11) - 1)));
        document.getElementById('btn-strikethrough')?.addEventListener('click', () => toggleStyle('strikethrough'));
        document.getElementById('btn-clear-format')?.addEventListener('click', () => {
            const sel = this.renderer.getSelectionRange();
            const sheet = this.workbook.getActiveSheet();
            for (let r = sel.minRow; r <= sel.maxRow; r++) {
                for (let c = sel.minCol; c <= sel.maxCol; c++) {
                    const cell = sheet.getCell(r, c);
                    if (cell) cell.style = {};
                }
            }
            this.renderer.render();
            this.updateToolbarState();
        });

        // Colors
        document.getElementById('font-color-picker')?.addEventListener('input', (e) => {
            toggleStyle('color', () => e.target.value);
            const bar = document.getElementById('font-color-bar');
            if (bar) bar.style.backgroundColor = e.target.value;
        });
        document.getElementById('bg-color-picker')?.addEventListener('input', (e) => {
            toggleStyle('background', () => e.target.value);
            const bar = document.getElementById('bg-color-bar');
            if (bar) bar.style.backgroundColor = e.target.value;
        });

        // Number format select
        document.getElementById('num-format-select')?.addEventListener('change', (e) => toggleStyle('numFormat', () => e.target.value));
        document.getElementById('btn-format-comma')?.addEventListener('click', () => toggleStyle('numFormat', () => 'number'));
        document.getElementById('btn-clear-filter')?.addEventListener('click', () => {
            const sheet = this.workbook.getActiveSheet();
            if (sheet.autoFilter) {
                sheet.autoFilter.filters.clear();
                this.renderer.render();
            }
        });

        // Cell style chips
        document.querySelectorAll('.cell-style-chip').forEach(chip => {
            chip.addEventListener('click', () => {
                const styleType = chip.dataset.style;
                if (styleType === 'neutral') {
                    toggleStyle('background', () => '#fef3c7');
                    toggleStyle('color', () => '#92400e');
                } else if (styleType === 'accent') {
                    toggleStyle('background', () => '#dbeafe');
                    toggleStyle('color', () => '#1e40af');
                    toggleStyle('bold', () => true);
                } else {
                    toggleStyle('background', () => '#f1f5f9');
                    toggleStyle('color', () => '#0f172a');
                }
            });
        });

        document.getElementById('file-input')?.addEventListener('change', (e) => this.handleFileImport(e));
    }

    handleMouseDown(e) {
        const rect = this.canvas.getBoundingClientRect();
        const mouseX = e.clientX - rect.left;
        const mouseY = e.clientY - rect.top;

        // Interactive Formula Building: Clicking a cell (or Ctrl-clicking) while editing a formula adds its reference
        if (this.isEditingFormula()) {
            const hit = this.renderer.screenToCell(mouseX, mouseY);
            if (!hit.inHeader) {
                this.insertFormulaCellReference(hit.row, hit.col);
                return;
            }
        }

        if (this.isEditingInCell()) {
            this.finishInCellEdit();
        }

        // Check for Filter Button Click
        const filterBtn = this.renderer.getFilterButtonAt(mouseX, mouseY);
        if (filterBtn) {
            this.openFilterPopup(filterBtn.col, e.clientX, e.clientY);
            return;
        }

        // Check for Drag-to-Fill Handle
        if (this.renderer.isOverFillHandle(mouseX, mouseY)) {
            const sel = this.renderer.getSelectionRange();
            this.renderer.fillState = {
                active: true,
                startRow: sel.minRow,
                startCol: sel.minCol,
                targetRow: sel.maxRow,
                targetCol: sel.maxCol
            };
            return;
        }

        // Check for Column Resizing
        if (mouseY < this.renderer.headerRowHeight && mouseX >= this.renderer.headerColWidth) {
            const gridX = mouseX - this.renderer.headerColWidth + this.renderer.scrollX;
            const col = this.renderer.sheet.getColAtX(gridX);
            const colRightEdge = this.renderer.headerColWidth + this.renderer.sheet.getColOffset(col + 1) - this.renderer.scrollX;

            if (Math.abs(mouseX - colRightEdge) <= 4) {
                this.renderer.resizeState = {
                    type: 'col',
                    index: col,
                    startPos: e.clientX,
                    originalSize: this.renderer.sheet.getColWidth(col),
                    active: true
                };
                return;
            }
        }

        // Check for Row Resizing
        if (mouseX < this.renderer.headerColWidth && mouseY >= this.renderer.headerRowHeight) {
            const gridY = mouseY - this.renderer.headerRowHeight + this.renderer.scrollY;
            const row = this.renderer.sheet.getRowAtY(gridY);
            const rowBottomEdge = this.renderer.headerRowHeight + this.renderer.sheet.getRowOffset(row + 1) - this.renderer.scrollY;

            if (Math.abs(mouseY - rowBottomEdge) <= 4) {
                this.renderer.resizeState = {
                    type: 'row',
                    index: row,
                    startPos: e.clientY,
                    originalSize: this.renderer.sheet.getRowHeight(row),
                    active: true
                };
                return;
            }
        }

        const hit = this.renderer.screenToCell(mouseX, mouseY);
        if (hit.inHeader) {
            if (hit.isColHeader && hit.col >= 0) {
                this.renderer.selection = {
                    activeRow: 0,
                    activeCol: hit.col,
                    startRow: 0,
                    startCol: hit.col,
                    endRow: 1048575,
                    endCol: hit.col,
                    isSelecting: false
                };
            } else if (hit.isRowHeader && hit.row >= 0) {
                this.renderer.selection = {
                    activeRow: hit.row,
                    activeCol: 0,
                    startRow: hit.row,
                    startCol: 0,
                    endRow: hit.row,
                    endCol: 16383,
                    isSelecting: false
                };
            }
            this.renderer.render();
            this.updateActiveCellDisplay();
            return;
        }

        if (e.shiftKey) {
            this.renderer.selection.endRow = hit.row;
            this.renderer.selection.endCol = hit.col;
        } else {
            this.renderer.selection = {
                activeRow: hit.row,
                activeCol: hit.col,
                startRow: hit.row,
                startCol: hit.col,
                endRow: hit.row,
                endCol: hit.col,
                isSelecting: true
            };
        }

        this.updateActiveCellDisplay();
        this.renderer.render();
    }

    handleMouseMove(e) {
        const rect = this.canvas.getBoundingClientRect();
        const mouseX = e.clientX - rect.left;
        const mouseY = e.clientY - rect.top;

        if (this.renderer.fillState.active) {
            const hit = this.renderer.screenToCell(mouseX, mouseY);
            if (!hit.inHeader) {
                this.renderer.fillState.targetRow = hit.row;
                this.renderer.fillState.targetCol = hit.col;
                this.renderer.render();
            }
            return;
        }

        if (this.renderer.resizeState.active) {
            const delta = (this.renderer.resizeState.type === 'col')
                ? e.clientX - this.renderer.resizeState.startPos
                : e.clientY - this.renderer.resizeState.startPos;

            const newSize = Math.max(20, this.renderer.resizeState.originalSize + delta);
            if (this.renderer.resizeState.type === 'col') {
                this.renderer.sheet.setColWidth(this.renderer.resizeState.index, newSize);
            } else {
                this.renderer.sheet.setRowHeight(this.renderer.resizeState.index, newSize);
            }
            this.renderer.render();
            return;
        }

        if (this.renderer.getFilterButtonAt(mouseX, mouseY)) {
            this.canvas.style.cursor = 'pointer';
            return;
        }

        if (this.renderer.isOverFillHandle(mouseX, mouseY)) {
            this.canvas.style.cursor = 'crosshair';
            return;
        }

        if (mouseY < this.renderer.headerRowHeight && mouseX >= this.renderer.headerColWidth) {
            const gridX = mouseX - this.renderer.headerColWidth + this.renderer.scrollX;
            const col = this.renderer.sheet.getColAtX(gridX);
            const colRightEdge = this.renderer.headerColWidth + this.renderer.sheet.getColOffset(col + 1) - this.renderer.scrollX;
            if (Math.abs(mouseX - colRightEdge) <= 4) {
                this.canvas.style.cursor = 'col-resize';
                return;
            }
        } else if (mouseX < this.renderer.headerColWidth && mouseY >= this.renderer.headerRowHeight) {
            const gridY = mouseY - this.renderer.headerRowHeight + this.renderer.scrollY;
            const row = this.renderer.sheet.getRowAtY(gridY);
            const rowBottomEdge = this.renderer.headerRowHeight + this.renderer.sheet.getRowOffset(row + 1) - this.renderer.scrollY;
            if (Math.abs(mouseY - rowBottomEdge) <= 4) {
                this.canvas.style.cursor = 'row-resize';
                return;
            }
        }
        this.canvas.style.cursor = 'cell';

        if (this.renderer.selection.isSelecting) {
            const hit = this.renderer.screenToCell(mouseX, mouseY);
            if (!hit.inHeader) {
                this.renderer.selection.endRow = hit.row;
                this.renderer.selection.endCol = hit.col;
                this.renderer.render();
            }
        }
    }

    handleMouseUp() {
        if (this.renderer.fillState.active) {
            this.executeFillDrag();
            this.renderer.fillState.active = false;
            this.renderer.render();
            return;
        }

        if (this.renderer.resizeState.active) {
            this.renderer.resizeState.active = false;
        }
        if (this.renderer.selection.isSelecting) {
            this.renderer.selection.isSelecting = false;
        }
    }

    executeFillDrag() {
        const sel = this.renderer.getSelectionRange();
        const sheet = this.renderer.sheet;
        const targetRow = this.renderer.fillState.targetRow;
        const targetCol = this.renderer.fillState.targetCol;
        const changes = [];

        if (targetRow > sel.maxRow) {
            const sourceHeight = sel.maxRow - sel.minRow + 1;
            for (let r = sel.maxRow + 1; r <= targetRow; r++) {
                const srcR = sel.minRow + ((r - sel.maxRow - 1) % sourceHeight);
                const dRow = r - srcR;

                for (let c = sel.minCol; c <= sel.maxCol; c++) {
                    const srcCell = sheet.getCell(srcR, c);
                    if (!srcCell) continue;

                    const prev = sheet.getCell(r, c);
                    const oldVal = prev ? prev.value : '';
                    const oldFormula = prev ? prev.formula : '';
                    const oldStyle = prev ? { ...prev.style } : {};

                    let newVal = srcCell.value;
                    let newFormula = srcCell.formula;
                    let newCell;

                    if (newFormula) {
                        newFormula = shiftFormula(newFormula, dRow, 0);
                        newCell = sheet.setCellValue(r, c, newFormula);
                    } else if (typeof srcCell.value === 'number') {
                        if (sourceHeight >= 2) {
                            const c0 = sheet.getCell(sel.minRow, c)?.value || 0;
                            const c1 = sheet.getCell(sel.minRow + 1, c)?.value || 0;
                            const step = (c1 - c0);
                            newVal = c0 + (r - sel.minRow) * step;
                        }
                        newCell = sheet.setCellValue(r, c, newVal);
                    } else {
                        newCell = sheet.setCellValue(r, c, newVal);
                    }

                    if (srcCell.style) {
                        sheet.setCellStyle(r, c, { ...srcCell.style });
                    }

                    changes.push({
                        row: r,
                        col: c,
                        oldVal,
                        oldFormula,
                        oldStyle,
                        newVal: newCell.value,
                        newFormula: newCell.formula,
                        newStyle: { ...newCell.style }
                    });
                }
            }
            this.renderer.selection.endRow = targetRow;
        } else if (targetCol > sel.maxCol) {
            const sourceWidth = sel.maxCol - sel.minCol + 1;
            for (let c = sel.maxCol + 1; c <= targetCol; c++) {
                const srcC = sel.minCol + ((c - sel.maxCol - 1) % sourceWidth);
                const dCol = c - srcC;

                for (let r = sel.minRow; r <= sel.maxRow; r++) {
                    const srcCell = sheet.getCell(r, srcC);
                    if (!srcCell) continue;

                    const prev = sheet.getCell(r, c);
                    const oldVal = prev ? prev.value : '';
                    const oldFormula = prev ? prev.formula : '';
                    const oldStyle = prev ? { ...prev.style } : {};

                    let newVal = srcCell.value;
                    let newFormula = srcCell.formula;
                    let newCell;

                    if (newFormula) {
                        newFormula = shiftFormula(newFormula, 0, dCol);
                        newCell = sheet.setCellValue(r, c, newFormula);
                    } else {
                        newCell = sheet.setCellValue(r, c, newVal);
                    }

                    if (srcCell.style) {
                        sheet.setCellStyle(r, c, { ...srcCell.style });
                    }

                    changes.push({
                        row: r,
                        col: c,
                        oldVal,
                        oldFormula,
                        oldStyle,
                        newVal: newCell.value,
                        newFormula: newCell.formula,
                        newStyle: { ...newCell.style }
                    });
                }
            }
            this.renderer.selection.endCol = targetCol;
        }

        if (changes.length > 0) {
            this.workbook.pushUndo(new CellChangeAction(sheet, changes));
        }
    }

    handleDoubleClick(e) {
        const rect = this.canvas.getBoundingClientRect();
        const mouseX = e.clientX - rect.left;
        const mouseY = e.clientY - rect.top;

        if (mouseY < this.renderer.headerRowHeight && mouseX >= this.renderer.headerColWidth) {
            const gridX = mouseX - this.renderer.headerColWidth + this.renderer.scrollX;
            const col = this.renderer.sheet.getColAtX(gridX);
            const colRightEdge = this.renderer.headerColWidth + this.renderer.sheet.getColOffset(col + 1) - this.renderer.scrollX;

            if (Math.abs(mouseX - colRightEdge) <= 4) {
                this.renderer.sheet.autoFitColumn(col, (text, style) => {
                    this.renderer.ctx.save();
                    this.renderer.ctx.font = `${style.bold ? 'bold ' : ''}${style.fontSize || 12}px ${style.fontFamily || 'Segoe UI'}`;
                    const w = this.renderer.ctx.measureText(text).width;
                    this.renderer.ctx.restore();
                    return w;
                });
                this.renderer.render();
                return;
            }
        }

        const hit = this.renderer.screenToCell(mouseX, mouseY);
        if (!hit.inHeader) {
            this.startInCellEdit(hit.row, hit.col);
        }
    }

    handleWheel(e) {
        e.preventDefault();
        this.renderer.scrollX = Math.max(0, this.renderer.scrollX + e.deltaX);
        this.renderer.scrollY = Math.max(0, this.renderer.scrollY + e.deltaY);

        this.renderer.render();
        if (this.isEditingInCell()) {
            this.positionInCellEditor();
        }
    }

    handleKeyDown(e) {
        if (this.isEditingInCell()) return;
        if (document.activeElement === this.formulaInput) return;

        const sel = this.renderer.selection;
        const isCtrl = e.ctrlKey || e.metaKey;
        const key = e.key.toLowerCase();

        // Undo: Ctrl+Z (without Shift)
        if (isCtrl && key === 'z' && !e.shiftKey) {
            e.preventDefault();
            this.undo();
            return;
        }

        // Redo: Ctrl+Y or Ctrl+Shift+Z
        if ((isCtrl && key === 'y') || (isCtrl && key === 'z' && e.shiftKey)) {
            e.preventDefault();
            this.redo();
            return;
        }

        // Copy: Ctrl+C
        if (isCtrl && key === 'c') {
            e.preventDefault();
            this.copySelectionToClipboard();
            return;
        }

        // Cut: Ctrl+X
        if (isCtrl && key === 'x') {
            e.preventDefault();
            this.cutSelectionToClipboard();
            return;
        }

        // Paste: Ctrl+V
        if (isCtrl && key === 'v') {
            e.preventDefault();
            this.pasteFromClipboard();
            return;
        }

        // Select All: Ctrl+A
        if (isCtrl && key === 'a') {
            e.preventDefault();
            this.selectAll();
            return;
        }

        // Bold: Ctrl+B
        if (isCtrl && key === 'b') {
            e.preventDefault();
            this.toggleStyle?.('bold');
            return;
        }

        // Italic: Ctrl+I
        if (isCtrl && key === 'i') {
            e.preventDefault();
            this.toggleStyle?.('italic');
            return;
        }

        // Underline: Ctrl+U
        if (isCtrl && key === 'u') {
            e.preventDefault();
            this.toggleStyle?.('underline');
            return;
        }

        // Save: Ctrl+S
        if (isCtrl && key === 's') {
            e.preventDefault();
            this.exportXLSX();
            return;
        }

        // Navigation: Ctrl+Home -> A1
        if (isCtrl && e.key === 'Home') {
            e.preventDefault();
            this.moveSelectionTo(0, 0);
            return;
        }

        // Navigation: Ctrl+End -> bottom-right of data
        if (isCtrl && e.key === 'End') {
            e.preventDefault();
            const bounds = this.renderer.sheet.data.getBounds();
            this.moveSelectionTo(Math.max(0, bounds.maxRow), Math.max(0, bounds.maxCol));
            return;
        }

        // Delete / Backspace
        if (e.key === 'Delete' || e.key === 'Backspace') {
            e.preventDefault();
            this.clearSelectionValues();
            return;
        }

        // Edit cell: F2
        if (e.key === 'F2') {
            e.preventDefault();
            this.startInCellEdit(sel.activeRow, sel.activeCol);
            return;
        }

        if (e.key === 'ArrowUp') {
            e.preventDefault();
            this.moveSelection(-1, 0, e.shiftKey);
        } else if (e.key === 'ArrowDown') {
            e.preventDefault();
            this.moveSelection(1, 0, e.shiftKey);
        } else if (e.key === 'ArrowLeft') {
            e.preventDefault();
            this.moveSelection(0, -1, e.shiftKey);
        } else if (e.key === 'ArrowRight') {
            e.preventDefault();
            this.moveSelection(0, 1, e.shiftKey);
        } else if (e.key === 'Tab') {
            e.preventDefault();
            this.moveSelection(0, e.shiftKey ? -1 : 1);
        } else if (e.key === 'Enter') {
            e.preventDefault();
            this.startInCellEdit(sel.activeRow, sel.activeCol);
        } else if (e.key.length === 1 && !e.ctrlKey && !e.altKey && !e.metaKey) {
            e.preventDefault();
            this.startInCellEdit(sel.activeRow, sel.activeCol, e.key);
        }
    }

    moveSelection(dRow, dCol, isRangeExtend = false) {
        const sel = this.renderer.selection;
        const newRow = Math.max(0, sel.activeRow + dRow);
        const newCol = Math.max(0, sel.activeCol + dCol);

        if (isRangeExtend) {
            sel.endRow = Math.max(0, sel.endRow + dRow);
            sel.endCol = Math.max(0, sel.endCol + dCol);
        } else {
            sel.activeRow = newRow;
            sel.activeCol = newCol;
            sel.startRow = newRow;
            sel.startCol = newCol;
            sel.endRow = newRow;
            sel.endCol = newCol;
        }

        this.ensureCellVisible(sel.activeRow, sel.activeCol);
        this.updateActiveCellDisplay();
        this.renderer.render();
    }

    moveSelectionTo(row, col) {
        const sel = this.renderer.selection;
        sel.activeRow = row;
        sel.activeCol = col;
        sel.startRow = row;
        sel.startCol = col;
        sel.endRow = row;
        sel.endCol = col;
        this.ensureCellVisible(row, col);
        this.updateActiveCellDisplay();
        this.renderer.render();
    }

    selectAll() {
        const bounds = this.renderer.sheet.data.getBounds();
        const sel = this.renderer.selection;
        sel.activeRow = 0;
        sel.activeCol = 0;
        sel.startRow = 0;
        sel.startCol = 0;
        sel.endRow = Math.max(bounds.maxRow, 30);
        sel.endCol = Math.max(bounds.maxCol, 15);
        this.updateActiveCellDisplay();
        this.renderer.render();
    }

    ensureCellVisible(row, col) {
        const bounds = this.renderer.getCellBounds(row, col);

        if (bounds.x < this.renderer.headerColWidth) {
            this.renderer.scrollX -= (this.renderer.headerColWidth - bounds.x);
        } else if (bounds.x + bounds.w > this.renderer.width) {
            this.renderer.scrollX += (bounds.x + bounds.w - this.renderer.width);
        }

        if (bounds.y < this.renderer.headerRowHeight) {
            this.renderer.scrollY -= (this.renderer.headerRowHeight - bounds.y);
        } else if (bounds.y + bounds.h > this.renderer.height) {
            this.renderer.scrollY += (bounds.y + bounds.h - this.renderer.height);
        }

        this.renderer.scrollX = Math.max(0, this.renderer.scrollX);
        this.renderer.scrollY = Math.max(0, this.renderer.scrollY);
    }

    updateActiveCellDisplay() {
        const r = this.renderer.selection.activeRow;
        const c = this.renderer.selection.activeCol;
        const coordStr = formatCoordinate(r, c);

        this.cellAddressBox.textContent = coordStr;

        const cell = this.renderer.sheet.getCell(r, c);
        const displayVal = cell ? (cell.formula || (cell.value !== undefined ? String(cell.value) : '')) : '';
        this.formulaInput.value = displayVal;

        this.updateToolbarState();
    }

    updateToolbarState() {
        const r = this.renderer.selection.activeRow;
        const c = this.renderer.selection.activeCol;
        const cell = this.renderer.sheet.getCell(r, c);
        const style = cell?.style || {};

        document.getElementById('btn-bold')?.classList.toggle('active', !!style.bold);
        document.getElementById('btn-italic')?.classList.toggle('active', !!style.italic);
        document.getElementById('btn-underline')?.classList.toggle('active', !!style.underline);

        const isMerged = !!this.renderer.sheet.getMerge(r, c);
        document.getElementById('btn-merge-center')?.classList.toggle('active', isMerged);
    }

    startInCellEdit(row, col, initialChar = null) {
        const cell = this.renderer.sheet.getCell(row, col);
        const currentFormula = cell?.formula || '';
        const currentValue = cell?.value !== undefined ? String(cell.value) : '';

        this.inCellEditor.value = (initialChar !== null) ? initialChar : (currentFormula || currentValue);
        this.inCellEditor.dataset.row = row;
        this.inCellEditor.dataset.col = col;

        this.positionInCellEditor();
        this.inCellEditor.style.display = 'block';
        this.inCellEditor.focus();
        if (initialChar === null) {
            this.inCellEditor.select();
        }
    }

    positionInCellEditor() {
        const r = parseInt(this.inCellEditor.dataset.row, 10);
        const c = parseInt(this.inCellEditor.dataset.col, 10);
        const bounds = this.renderer.getCellBounds(r, c);

        this.inCellEditor.style.left = `${bounds.x}px`;
        this.inCellEditor.style.top = `${bounds.y}px`;
        this.inCellEditor.style.width = `${Math.max(bounds.w, 120)}px`;
        this.inCellEditor.style.height = `${bounds.h}px`;
    }

    finishInCellEdit() {
        if (!this.isEditingInCell()) return;
        const r = parseInt(this.inCellEditor.dataset.row, 10);
        const c = parseInt(this.inCellEditor.dataset.col, 10);
        const val = this.inCellEditor.value;
        const sheet = this.workbook.getActiveSheet();

        const prevCell = sheet.getCell(r, c);
        const oldVal = prevCell ? prevCell.value : '';
        const oldFormula = prevCell ? prevCell.formula : '';
        const oldStyle = prevCell ? { ...prevCell.style } : {};

        this.inCellEditor.style.display = 'none';
        this.renderer.activeFormulaRefCell = null;

        if (val !== (oldFormula || (oldVal !== undefined ? String(oldVal) : ''))) {
            const newCell = sheet.setCellValue(r, c, val);
            this.workbook.pushUndo(new CellChangeAction(sheet, [{
                row: r,
                col: c,
                oldVal,
                oldFormula,
                oldStyle,
                newVal: newCell.value,
                newFormula: newCell.formula,
                newStyle: { ...newCell.style }
            }]));
        }

        this.updateActiveCellDisplay();
        this.renderer.render();
    }

    cancelInCellEdit() {
        this.inCellEditor.style.display = 'none';
        this.renderer.activeFormulaRefCell = null;
        this.updateActiveCellDisplay();
        this.renderer.render();
    }

    isEditingInCell() {
        return this.inCellEditor.style.display === 'block';
    }

    isEditingFormula() {
        if (this.isEditingInCell()) {
            return this.inCellEditor.value.trim().startsWith('=');
        }
        if (document.activeElement === this.formulaInput) {
            return this.formulaInput.value.trim().startsWith('=');
        }
        return false;
    }

    insertFormulaCellReference(row, col) {
        const refStr = formatCoordinate(row, col);
        const activeInput = this.isEditingInCell() ? this.inCellEditor : this.formulaInput;
        const val = activeInput.value;
        const start = activeInput.selectionStart ?? val.length;
        const end = activeInput.selectionEnd ?? val.length;

        // Check if there is an existing cell coordinate right before/at cursor
        const beforeCursor = val.slice(0, start);
        const afterCursor = val.slice(end);

        const match = beforeCursor.match(/(\$?[A-Za-z]+\$?[0-9]+)$/);
        let newVal;
        let newCursorPos;

        if (match) {
            // Replace previous reference
            const prefix = beforeCursor.slice(0, beforeCursor.length - match[1].length);
            newVal = prefix + refStr + afterCursor;
            newCursorPos = prefix.length + refStr.length;
        } else {
            // Append or insert at cursor
            newVal = beforeCursor + refStr + afterCursor;
            newCursorPos = beforeCursor.length + refStr.length;
        }

        activeInput.value = newVal;
        if (this.isEditingInCell()) {
            this.formulaInput.value = newVal;
        } else {
            this.inCellEditor.value = newVal;
        }

        activeInput.focus();
        activeInput.setSelectionRange(newCursorPos, newCursorPos);

        // Highlight referenced cell with dashed outline on grid
        this.renderer.activeFormulaRefCell = { row, col, color: '#2563eb' };
        this.renderer.render();
    }

    finishEditing() {
        if (this.isEditingInCell()) {
            this.finishInCellEdit();
        }
    }

    undo() {
        const action = this.workbook.undo();
        if (action) {
            this.updateActiveCellDisplay();
            this.renderer.render();
        }
    }

    redo() {
        const action = this.workbook.redo();
        if (action) {
            this.updateActiveCellDisplay();
            this.renderer.render();
        }
    }

    copySelectionToClipboard() {
        const sel = this.renderer.getSelectionRange();
        const sheet = this.renderer.sheet;
        const rows = [];

        for (let r = sel.minRow; r <= sel.maxRow; r++) {
            const cols = [];
            for (let c = sel.minCol; c <= sel.maxCol; c++) {
                const cell = sheet.getCell(r, c);
                if (!cell) {
                    cols.push('');
                } else if (cell.formula) {
                    cols.push(cell.formula);
                } else if (cell.value !== undefined && cell.value !== null) {
                    cols.push(String(cell.value));
                } else {
                    cols.push('');
                }
            }
            rows.push(cols.join('\t'));
        }

        const tsv = rows.join('\r\n');
        if (navigator.clipboard && navigator.clipboard.writeText) {
            navigator.clipboard.writeText(tsv).catch(() => {
                this.fallbackCopyToClipboard(tsv);
            });
        } else {
            this.fallbackCopyToClipboard(tsv);
        }
    }

    fallbackCopyToClipboard(text) {
        try {
            const textarea = document.createElement('textarea');
            textarea.value = text;
            textarea.style.position = 'fixed';
            textarea.style.left = '-9999px';
            textarea.style.top = '-9999px';
            textarea.setAttribute('readonly', '');
            document.body.appendChild(textarea);
            textarea.select();
            document.execCommand('copy');
            document.body.removeChild(textarea);
        } catch (err) {
            console.warn('Fallback copy failed:', err);
        }
    }

    cutSelectionToClipboard() {
        this.copySelectionToClipboard();
        this.clearSelectionValues();
    }

    async pasteFromClipboard() {
        try {
            if (navigator.clipboard && navigator.clipboard.readText) {
                const text = await navigator.clipboard.readText();
                if (text) {
                    this.pasteData(text);
                    return;
                }
            }
        } catch (err) {
            console.warn('Clipboard readText restricted, fallback to paste event handler:', err);
        }
    }

    pasteData(text) {
        if (!text || typeof text !== 'string') return;
        const sel = this.renderer.selection;
        const startRow = sel.activeRow;
        const startCol = sel.activeCol;
        const sheet = this.workbook.getActiveSheet();

        const normalized = text.replace(/\r\n/g, '\n').replace(/\r/g, '\n');
        let rows = normalized.split('\n');
        if (rows.length > 1 && rows[rows.length - 1] === '') {
            rows.pop();
        }

        const changes = [];
        let maxCols = 1;

        for (let i = 0; i < rows.length; i++) {
            const rowStr = rows[i];
            const cells = rowStr.includes('\t') ? rowStr.split('\t') : [rowStr];
            maxCols = Math.max(maxCols, cells.length);

            for (let j = 0; j < cells.length; j++) {
                const r = startRow + i;
                const c = startCol + j;
                const prevCell = sheet.getCell(r, c);
                const oldVal = prevCell ? prevCell.value : '';
                const oldFormula = prevCell ? prevCell.formula : '';
                const oldStyle = prevCell ? { ...prevCell.style } : {};

                const rawVal = cells[j].trim();
                const newCell = sheet.setCellValue(r, c, rawVal);

                changes.push({
                    row: r,
                    col: c,
                    oldVal,
                    oldFormula,
                    oldStyle,
                    newVal: newCell.value,
                    newFormula: newCell.formula,
                    newStyle: { ...newCell.style }
                });
            }
        }

        if (changes.length > 0) {
            this.workbook.pushUndo(new CellChangeAction(sheet, changes));
        }

        sel.startRow = startRow;
        sel.startCol = startCol;
        sel.endRow = startRow + rows.length - 1;
        sel.endCol = startCol + maxCols - 1;

        this.ensureCellVisible(sel.endRow, sel.endCol);
        this.updateActiveCellDisplay();
        this.renderer.render();
    }

    clearSelectionValues() {
        const sel = this.renderer.getSelectionRange();
        const sheet = this.renderer.sheet;
        const changes = [];

        for (let r = sel.minRow; r <= sel.maxRow; r++) {
            for (let c = sel.minCol; c <= sel.maxCol; c++) {
                const cell = sheet.getCell(r, c);
                if (cell && (cell.value !== '' || cell.formula !== '')) {
                    changes.push({
                        row: r,
                        col: c,
                        oldVal: cell.value,
                        oldFormula: cell.formula,
                        oldStyle: { ...cell.style },
                        newVal: '',
                        newFormula: '',
                        newStyle: { ...cell.style }
                    });
                    sheet.setCellValue(r, c, '');
                }
            }
        }

        if (changes.length > 0) {
            this.workbook.pushUndo(new CellChangeAction(sheet, changes));
        }
        this.updateActiveCellDisplay();
        this.renderer.render();
    }

    // AutoFilter interactive dropdown popup
    openFilterPopup(col, screenX, screenY) {
        const sheet = this.workbook.getActiveSheet();
        if (!sheet.autoFilter) return;

        const af = sheet.autoFilter;
        let hiddenVals = af.filters.get(col);
        if (!hiddenVals) {
            hiddenVals = new Set();
            af.filters.set(col, hiddenVals);
        }

        // Collect unique values in this column
        const uniqueValues = new Set();
        for (let r = af.r1 + 1; r <= af.r2; r++) {
            const cell = sheet.getCell(r, col);
            uniqueValues.add(cell ? String(cell.value) : '');
        }

        const sortedVals = Array.from(uniqueValues).sort();

        this.filterPopup.innerHTML = `
            <div style="font-weight: 600; font-size: 12px; margin-bottom: 8px; border-bottom: 1px solid #e2e8f0; padding-bottom: 4px;">
                Filter: ${colToName(col)}
            </div>
            <div style="max-height: 180px; overflow-y: auto;">
                ${sortedVals.map(val => `
                    <label style="display: flex; align-items: center; gap: 6px; font-size: 12px; padding: 2px 0; cursor: pointer;">
                        <input type="checkbox" value="${encodeURIComponent(val)}" ${!hiddenVals.has(val) ? 'checked' : ''}>
                        <span>${val || '(Blanks)'}</span>
                    </label>
                `).join('')}
            </div>
            <div style="display: flex; justify-content: flex-end; gap: 6px; margin-top: 8px; border-top: 1px solid #e2e8f0; padding-top: 6px;">
                <button id="btn-apply-filter" style="padding: 2px 8px; font-size: 11px; background: #107c41; color: white; border: none; border-radius: 3px; cursor: pointer;">Apply</button>
            </div>
        `;

        this.filterPopup.style.left = `${Math.min(window.innerWidth - 200, screenX)}px`;
        this.filterPopup.style.top = `${screenY + 10}px`;
        this.filterPopup.style.display = 'block';

        document.getElementById('btn-apply-filter')?.addEventListener('click', () => {
            const checkboxes = this.filterPopup.querySelectorAll('input[type="checkbox"]');
            hiddenVals.clear();
            checkboxes.forEach(cb => {
                const val = decodeURIComponent(cb.value);
                if (!cb.checked) {
                    hiddenVals.add(val);
                }
            });
            this.closeFilterPopup();
            this.renderer.render();
        });
    }

    closeFilterPopup() {
        if (this.filterPopup) {
            this.filterPopup.style.display = 'none';
        }
    }

    renderSheetTabs() {
        this.sheetTabsContainer.innerHTML = '';
        this.workbook.sheets.forEach((sheet, idx) => {
            const tab = document.createElement('div');
            tab.className = `sheet-tab ${idx === this.workbook.activeSheetIndex ? 'active' : ''}`;
            tab.textContent = sheet.name;
            tab.addEventListener('click', () => {
                this.workbook.setActiveSheet(idx);
                this.renderSheetTabs();
                this.updateActiveCellDisplay();
                this.renderer.render();
            });
            this.sheetTabsContainer.appendChild(tab);
        });
    }

    async exportXLSX() {
        if (!window.ExcelJS) {
            alert('Excel engine is loading, please try again in a moment.');
            return;
        }
        const buffer = await this.workbook.toXLSXBuffer(window.ExcelJS);
        const blob = new Blob([buffer], { type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' });
        const url = URL.createObjectURL(blob);
        const a = document.createElement('a');
        a.href = url;
        a.download = `${this.workbook.title}.xlsx`;
        a.click();
        URL.revokeObjectURL(url);
    }

    exportCSV() {
        const csv = this.workbook.toCSV();
        const blob = new Blob([csv], { type: 'text/csv;charset=utf-8;' });
        const url = URL.createObjectURL(blob);
        const a = document.createElement('a');
        a.href = url;
        a.download = `${this.workbook.title}.csv`;
        a.click();
        URL.revokeObjectURL(url);
    }

    exportJSON() {
        const json = JSON.stringify(this.workbook.toJSON(), null, 2);
        const blob = new Blob([json], { type: 'application/json' });
        const url = URL.createObjectURL(blob);
        const a = document.createElement('a');
        a.href = url;
        a.download = `${this.workbook.title}.json`;
        a.click();
        URL.revokeObjectURL(url);
    }

    async handleFileImport(e) {
        const file = e.target.files[0];
        if (!file) return;

        if (file.name.endsWith('.xlsx')) {
            if (!window.ExcelJS) {
                alert('Excel engine is still loading, please retry.');
                return;
            }
            const arrayBuffer = await file.arrayBuffer();
            await this.workbook.fromXLSXBuffer(arrayBuffer, window.ExcelJS);
            this.workbook.title = file.name.replace('.xlsx', '');
            this.renderSheetTabs();
            this.updateActiveCellDisplay();
            this.renderer.render();
        } else if (file.name.endsWith('.json')) {
            const text = await file.text();
            const data = JSON.parse(text);
            this.workbook.fromJSON(data);
            this.renderSheetTabs();
            this.updateActiveCellDisplay();
            this.renderer.render();
        } else if (file.name.endsWith('.csv')) {
            const text = await file.text();
            this.workbook.fromCSV(text, file.name.replace('.csv', ''));
            this.renderSheetTabs();
            this.updateActiveCellDisplay();
            this.renderer.render();
        }
    }
}

window.addEventListener('DOMContentLoaded', () => {
    window.app = new SpreadsheetApp();
});
