import { colToName, formatCoordinate } from '../core/cell.js';

export class CanvasGridRenderer {
    constructor(canvas, workbook) {
        this.canvas = canvas;
        this.ctx = canvas.getContext('2d');
        this.workbook = workbook;

        // Viewport scroll offsets
        this.scrollX = 0;
        this.scrollY = 0;

        // Header dimensions
        this.headerColWidth = 50;  // Row number column width
        this.headerRowHeight = 26; // Column letter header height

        // Selection state
        this.selection = {
            activeRow: 0,
            activeCol: 0,
            startRow: 0,
            startCol: 0,
            endRow: 0,
            endCol: 0,
            isSelecting: false
        };

        // Fill drag handle state
        this.fillState = {
            active: false,
            startRow: 0,
            startCol: 0,
            targetRow: 0,
            targetCol: 0
        };

        // Active formula reference cell highlight (row, col, color)
        this.activeFormulaRefCell = null;

        // Resize state
        this.resizeState = {
            type: null,
            index: -1,
            startPos: 0,
            originalSize: 0,
            active: false
        };

        // Filter button click boxes: Array<{ col, x, y, size }>
        this.filterButtons = [];

        // High DPI handling
        this.dpr = window.devicePixelRatio || 1;
        this.width = 0;
        this.height = 0;

        this.initCanvasSize();
    }

    initCanvasSize() {
        const rect = this.canvas.parentElement.getBoundingClientRect();
        this.width = rect.width;
        this.height = rect.height;

        this.canvas.width = Math.floor(this.width * this.dpr);
        this.canvas.height = Math.floor(this.height * this.dpr);

        this.canvas.style.width = `${this.width}px`;
        this.canvas.style.height = `${this.height}px`;

        this.ctx.scale(this.dpr, this.dpr);
    }

    onResize() {
        this.initCanvasSize();
        this.render();
    }

    get sheet() {
        return this.workbook.getActiveSheet();
    }

    screenToCell(screenX, screenY) {
        if (screenX < this.headerColWidth || screenY < this.headerRowHeight) {
            return {
                inHeader: true,
                isColHeader: screenY < this.headerRowHeight && screenX >= this.headerColWidth,
                isRowHeader: screenX < this.headerColWidth && screenY >= this.headerRowHeight,
                col: screenX >= this.headerColWidth ? this.sheet.getColAtX(screenX - this.headerColWidth + this.scrollX) : -1,
                row: screenY >= this.headerRowHeight ? this.sheet.getRowAtY(screenY - this.headerRowHeight + this.scrollY) : -1
            };
        }

        const gridX = screenX - this.headerColWidth + this.scrollX;
        const gridY = screenY - this.headerRowHeight + this.scrollY;

        return {
            inHeader: false,
            row: this.sheet.getRowAtY(gridY),
            col: this.sheet.getColAtX(gridX)
        };
    }

    getCellBounds(row, col) {
        const sheet = this.sheet;
        const merge = sheet.getMerge(row, col);

        if (merge) {
            const x = this.headerColWidth + sheet.getColOffset(merge.c1) - this.scrollX;
            const y = this.headerRowHeight + sheet.getRowOffset(merge.r1) - this.scrollY;
            let w = 0;
            for (let c = merge.c1; c <= merge.c2; c++) w += sheet.getColWidth(c);
            let h = 0;
            for (let r = merge.r1; r <= merge.r2; r++) h += sheet.getRowHeight(r);
            return { x, y, w, h, isMerged: true, merge };
        }

        const x = this.headerColWidth + sheet.getColOffset(col) - this.scrollX;
        const y = this.headerRowHeight + sheet.getRowOffset(row) - this.scrollY;
        const w = sheet.getColWidth(col);
        const h = sheet.getRowHeight(row);
        return { x, y, w, h, isMerged: false };
    }

    getSelectionRange() {
        const minRow = Math.min(this.selection.startRow, this.selection.endRow);
        const maxRow = Math.max(this.selection.startRow, this.selection.endRow);
        const minCol = Math.min(this.selection.startCol, this.selection.endCol);
        const maxCol = Math.max(this.selection.startCol, this.selection.endCol);
        return { minRow, maxRow, minCol, maxCol };
    }

    isOverFillHandle(screenX, screenY) {
        const sel = this.getSelectionRange();
        const sheet = this.sheet;
        const selX2 = this.headerColWidth + sheet.getColOffset(sel.maxCol + 1) - this.scrollX;
        const selY2 = this.headerRowHeight + sheet.getRowOffset(sel.maxRow + 1) - this.scrollY;

        const distance = Math.hypot(screenX - selX2, screenY - selY2);
        return distance <= 7;
    }

    getFilterButtonAt(screenX, screenY) {
        for (const btn of this.filterButtons) {
            if (screenX >= btn.x && screenX <= btn.x + btn.size &&
                screenY >= btn.y && screenY <= btn.y + btn.size) {
                return btn;
            }
        }
        return null;
    }

    render() {
        const ctx = this.ctx;
        const sheet = this.sheet;
        const width = this.width;
        const height = this.height;

        this.filterButtons = [];
        ctx.clearRect(0, 0, width, height);

        // Calculate frozen offsets if freeze panes are active
        const freezeRowH = sheet.freezeRow > 0 ? sheet.getRowOffset(sheet.freezeRow) : 0;
        const freezeColW = sheet.freezeCol > 0 ? sheet.getColOffset(sheet.freezeCol) : 0;

        // Virtualization window
        const startCol = sheet.getColAtX(this.scrollX);
        const endCol = sheet.getColAtX(this.scrollX + width - this.headerColWidth) + 1;

        const startRow = sheet.getRowAtY(this.scrollY);
        const endRow = sheet.getRowAtY(this.scrollY + height - this.headerRowHeight) + 1;

        const sel = this.getSelectionRange();

        // 1. Draw Grid Cells
        ctx.save();
        ctx.beginPath();
        ctx.rect(this.headerColWidth, this.headerRowHeight, width - this.headerColWidth, height - this.headerRowHeight);
        ctx.clip();

        // Render cell backgrounds
        for (let r = startRow; r <= endRow; r++) {
            if (sheet.isRowHidden(r)) continue;
            const rowY = this.headerRowHeight + sheet.getRowOffset(r) - this.scrollY;
            const rowH = sheet.getRowHeight(r);

            for (let c = startCol; c <= endCol; c++) {
                const colX = this.headerColWidth + sheet.getColOffset(c) - this.scrollX;
                const colW = sheet.getColWidth(c);
                const cell = sheet.getCell(r, c);

                if (cell && cell.style && cell.style.background && cell.style.background !== '#ffffff') {
                    ctx.fillStyle = cell.style.background;
                    ctx.fillRect(colX, rowY, colW, rowH);
                }
            }
        }

        // Render selection highlight background
        if (sel.minRow <= endRow && sel.maxRow >= startRow && sel.minCol <= endCol && sel.maxCol >= startCol) {
            const selX1 = this.headerColWidth + sheet.getColOffset(sel.minCol) - this.scrollX;
            const selY1 = this.headerRowHeight + sheet.getRowOffset(sel.minRow) - this.scrollY;
            const selX2 = this.headerColWidth + sheet.getColOffset(sel.maxCol + 1) - this.scrollX;
            const selY2 = this.headerRowHeight + sheet.getRowOffset(sel.maxRow + 1) - this.scrollY;

            ctx.fillStyle = 'rgba(16, 124, 65, 0.08)';
            ctx.fillRect(selX1, selY1, selX2 - selX1, selY2 - selY1);
        }

        // Render Fill Drag Preview Box
        if (this.fillState.active) {
            const minR = Math.min(sel.minRow, this.fillState.targetRow);
            const maxR = Math.max(sel.maxRow, this.fillState.targetRow);
            const minC = Math.min(sel.minCol, this.fillState.targetCol);
            const maxC = Math.max(sel.maxCol, this.fillState.targetCol);

            const fx1 = this.headerColWidth + sheet.getColOffset(minC) - this.scrollX;
            const fy1 = this.headerRowHeight + sheet.getRowOffset(minR) - this.scrollY;
            const fx2 = this.headerColWidth + sheet.getColOffset(maxC + 1) - this.scrollX;
            const fy2 = this.headerRowHeight + sheet.getRowOffset(maxR + 1) - this.scrollY;

            ctx.save();
            ctx.strokeStyle = '#107c41';
            ctx.setLineDash([4, 4]);
            ctx.lineWidth = 1.5;
            ctx.strokeRect(fx1 + 0.5, fy1 + 0.5, fx2 - fx1, fy2 - fy1);
            ctx.restore();
        }

        // Render Grid Lines
        ctx.strokeStyle = '#e2e8f0';
        ctx.lineWidth = 1;
        ctx.beginPath();

        for (let r = startRow; r <= endRow + 1; r++) {
            if (sheet.isRowHidden(r)) continue;
            const y = Math.floor(this.headerRowHeight + sheet.getRowOffset(r) - this.scrollY) + 0.5;
            ctx.moveTo(this.headerColWidth, y);
            ctx.lineTo(width, y);
        }

        for (let c = startCol; c <= endCol + 1; c++) {
            const x = Math.floor(this.headerColWidth + sheet.getColOffset(c) - this.scrollX) + 0.5;
            ctx.moveTo(x, this.headerRowHeight);
            ctx.lineTo(x, height);
        }
        ctx.stroke();

        // Render Merged Cells Overlays
        for (const m of sheet.merges) {
            if (m.r2 < startRow || m.r1 > endRow || m.c2 < startCol || m.c1 > endCol) continue;

            const mx = this.headerColWidth + sheet.getColOffset(m.c1) - this.scrollX;
            const my = this.headerRowHeight + sheet.getRowOffset(m.r1) - this.scrollY;
            let mw = 0;
            for (let c = m.c1; c <= m.c2; c++) mw += sheet.getColWidth(c);
            let mh = 0;
            for (let r = m.r1; r <= m.r2; r++) mh += sheet.getRowHeight(r);

            const masterCell = sheet.getCell(m.r1, m.c1);
            ctx.fillStyle = masterCell?.style?.background || '#ffffff';
            ctx.fillRect(mx + 0.5, my + 0.5, mw - 1, mh - 1);

            if (masterCell && masterCell.value !== undefined && masterCell.value !== null && masterCell.value !== '') {
                ctx.save();
                ctx.beginPath();
                ctx.rect(mx + 2, my + 1, mw - 4, mh - 2);
                ctx.clip();

                let fontStyle = '';
                if (masterCell.style.bold) fontStyle += 'bold ';
                if (masterCell.style.italic) fontStyle += 'italic ';
                ctx.font = `${fontStyle}${masterCell.style.fontSize}px ${masterCell.style.fontFamily}`;
                ctx.fillStyle = masterCell.style.color || '#1e293b';
                ctx.textAlign = 'center';
                ctx.textBaseline = 'middle';
                ctx.fillText(String(masterCell.value), mx + mw / 2, my + mh / 2);
                ctx.restore();
            }
        }

        // Render Standard Cell Content (Text & Numbers)
        ctx.textBaseline = 'middle';
        for (let r = startRow; r <= endRow; r++) {
            if (sheet.isRowHidden(r)) continue;
            const rowY = this.headerRowHeight + sheet.getRowOffset(r) - this.scrollY;
            const rowH = sheet.getRowHeight(r);

            for (let c = startCol; c <= endCol; c++) {
                const merge = sheet.getMerge(r, c);
                if (merge) continue;

                const cell = sheet.getCell(r, c);
                if (!cell || cell.value === undefined || cell.value === null || cell.value === '') continue;

                const colX = this.headerColWidth + sheet.getColOffset(c) - this.scrollX;
                const colW = sheet.getColWidth(c);

                ctx.save();
                ctx.beginPath();
                ctx.rect(colX + 2, rowY + 1, colW - 4, rowH - 2);
                ctx.clip();

                let fontStyle = '';
                if (cell.style.bold) fontStyle += 'bold ';
                if (cell.style.italic) fontStyle += 'italic ';
                ctx.font = `${fontStyle}${cell.style.fontSize}px ${cell.style.fontFamily}`;
                ctx.fillStyle = cell.style.color || '#1e293b';

                let displayText = String(cell.value);

                if (cell.style.numFormat === 'currency' && typeof cell.value === 'number') {
                    displayText = '$' + cell.value.toLocaleString(undefined, { minimumFractionDigits: 2, maximumFractionDigits: 2 });
                } else if (cell.style.numFormat === 'percent' && typeof cell.value === 'number') {
                    displayText = (cell.value * 100).toFixed(1) + '%';
                }

                let textX = colX + 6;
                const textY = rowY + rowH / 2;
                const isNumber = typeof cell.value === 'number';
                const align = cell.style.align || (isNumber ? 'right' : 'left');

                if (align === 'right') {
                    textX = colX + colW - 6;
                    ctx.textAlign = 'right';
                } else if (align === 'center') {
                    textX = colX + colW / 2;
                    ctx.textAlign = 'center';
                } else {
                    ctx.textAlign = 'left';
                }

                ctx.fillText(displayText, textX, textY);
                ctx.restore();
            }
        }

        // Render AutoFilter Dropdown Buttons on Header Row
        if (sheet.autoFilter) {
            const filterRow = sheet.autoFilter.r1;
            const rowY = this.headerRowHeight + sheet.getRowOffset(filterRow) - this.scrollY;
            const rowH = sheet.getRowHeight(filterRow);

            for (let c = sheet.autoFilter.c1; c <= sheet.autoFilter.c2; c++) {
                const colX = this.headerColWidth + sheet.getColOffset(c) - this.scrollX;
                const colW = sheet.getColWidth(c);
                const btnSize = 14;
                const btnX = colX + colW - btnSize - 4;
                const btnY = rowY + (rowH - btnSize) / 2;

                // Draw button
                ctx.fillStyle = '#e2e8f0';
                ctx.fillRect(btnX, btnY, btnSize, btnSize);
                ctx.strokeStyle = '#94a3b8';
                ctx.strokeRect(btnX, btnY, btnSize, btnSize);

                // Draw dropdown arrow
                ctx.fillStyle = '#475569';
                ctx.beginPath();
                ctx.moveTo(btnX + 3, btnY + 5);
                ctx.lineTo(btnX + btnSize - 3, btnY + 5);
                ctx.lineTo(btnX + btnSize / 2, btnY + btnSize - 4);
                ctx.fill();

                this.filterButtons.push({ col: c, x: btnX, y: btnY, size: btnSize });
            }
        }

        // Render Freeze Panes Line Indicator
        if (sheet.freezeRow > 0) {
            const fy = this.headerRowHeight + sheet.getRowOffset(sheet.freezeRow) - this.scrollY;
            ctx.save();
            ctx.strokeStyle = '#0f766e';
            ctx.lineWidth = 2;
            ctx.beginPath();
            ctx.moveTo(this.headerColWidth, fy);
            ctx.lineTo(width, fy);
            ctx.stroke();
            ctx.restore();
        }

        // Render Selection Border & Active Cell Box
        if (sel.minRow <= endRow && sel.maxRow >= startRow && sel.minCol <= endCol && sel.maxCol >= startCol) {
            const selX1 = Math.floor(this.headerColWidth + sheet.getColOffset(sel.minCol) - this.scrollX) + 0.5;
            const selY1 = Math.floor(this.headerRowHeight + sheet.getRowOffset(sel.minRow) - this.scrollY) + 0.5;
            const selX2 = Math.floor(this.headerColWidth + sheet.getColOffset(sel.maxCol + 1) - this.scrollX) + 0.5;
            const selY2 = Math.floor(this.headerRowHeight + sheet.getRowOffset(sel.maxRow + 1) - this.scrollY) + 0.5;

            ctx.strokeStyle = '#107c41';
            ctx.lineWidth = 2;
            ctx.strokeRect(selX1, selY1, selX2 - selX1, selY2 - selY1);

            if (sel.minRow !== sel.maxRow || sel.minCol !== sel.maxCol) {
                const actX = Math.floor(this.headerColWidth + sheet.getColOffset(this.selection.activeCol) - this.scrollX) + 0.5;
                const actY = Math.floor(this.headerRowHeight + sheet.getRowOffset(this.selection.activeRow) - this.scrollY) + 0.5;
                const actW = sheet.getColWidth(this.selection.activeCol);
                const actH = sheet.getRowHeight(this.selection.activeRow);

                ctx.fillStyle = '#ffffff';
                ctx.fillRect(actX, actY, actW, actH);
                ctx.strokeStyle = '#107c41';
                ctx.lineWidth = 1;
                ctx.strokeRect(actX, actY, actW, actH);

                const actCell = sheet.getCell(this.selection.activeRow, this.selection.activeCol);
                if (actCell && actCell.value !== undefined && actCell.value !== null && actCell.value !== '') {
                    ctx.save();
                    ctx.beginPath();
                    ctx.rect(actX + 2, actY + 1, actW - 4, actH - 2);
                    ctx.clip();
                    ctx.fillStyle = actCell.style.color || '#1e293b';
                    ctx.font = `${actCell.style.bold ? 'bold ' : ''}${actCell.style.fontSize}px ${actCell.style.fontFamily}`;
                    ctx.textAlign = (typeof actCell.value === 'number') ? 'right' : 'left';
                    const tx = (typeof actCell.value === 'number') ? actX + actW - 6 : actX + 6;
                    ctx.fillText(String(actCell.value), tx, actY + actH / 2);
                    ctx.restore();
                }
            }

            ctx.fillStyle = '#107c41';
            ctx.fillRect(selX2 - 4, selY2 - 4, 7, 7);
        }

        // Render formula reference cell highlights (Excel-like colored dashed box)
        if (this.activeFormulaRefCell) {
            const { row: rfR, col: rfC, color = '#2563eb' } = this.activeFormulaRefCell;
            const rx = Math.floor(this.headerColWidth + sheet.getColOffset(rfC) - this.scrollX) + 0.5;
            const ry = Math.floor(this.headerRowHeight + sheet.getRowOffset(rfR) - this.scrollY) + 0.5;
            const rw = sheet.getColWidth(rfC);
            const rh = sheet.getRowHeight(rfR);

            ctx.save();
            ctx.strokeStyle = color;
            ctx.lineWidth = 2;
            ctx.setLineDash([4, 4]);
            ctx.fillStyle = 'rgba(37, 99, 235, 0.12)';
            ctx.fillRect(rx, ry, rw, rh);
            ctx.strokeRect(rx, ry, rw, rh);
            ctx.restore();
        }

        ctx.restore();

        // 2. Render Column Headers (A, B, C...)
        ctx.save();
        ctx.fillStyle = '#f8fafc';
        ctx.fillRect(this.headerColWidth, 0, width - this.headerColWidth, this.headerRowHeight);
        ctx.strokeStyle = '#cbd5e1';
        ctx.lineWidth = 1;

        ctx.beginPath();
        ctx.moveTo(this.headerColWidth, this.headerRowHeight - 0.5);
        ctx.lineTo(width, this.headerRowHeight - 0.5);
        ctx.stroke();

        ctx.font = '600 11px Segoe UI, sans-serif';
        ctx.textAlign = 'center';
        ctx.textBaseline = 'middle';

        for (let c = startCol; c <= endCol; c++) {
            const colX = this.headerColWidth + sheet.getColOffset(c) - this.scrollX;
            const colW = sheet.getColWidth(c);

            const isColSelected = c >= sel.minCol && c <= sel.maxCol;
            if (isColSelected) {
                ctx.fillStyle = '#e2e8f0';
                ctx.fillRect(colX, 0, colW, this.headerRowHeight - 1);
                ctx.fillStyle = '#107c41';
                ctx.fillRect(colX, this.headerRowHeight - 3, colW, 2);
            }

            ctx.strokeStyle = '#cbd5e1';
            ctx.beginPath();
            ctx.moveTo(colX + colW - 0.5, 0);
            ctx.lineTo(colX + colW - 0.5, this.headerRowHeight);
            ctx.stroke();

            ctx.fillStyle = isColSelected ? '#107c41' : '#64748b';
            ctx.fillText(colToName(c), colX + colW / 2, this.headerRowHeight / 2);
        }
        ctx.restore();

        // 3. Render Row Headers (1, 2, 3...)
        ctx.save();
        ctx.fillStyle = '#f8fafc';
        ctx.fillRect(0, this.headerRowHeight, this.headerColWidth, height - this.headerRowHeight);

        ctx.strokeStyle = '#cbd5e1';
        ctx.lineWidth = 1;
        ctx.beginPath();
        ctx.moveTo(this.headerColWidth - 0.5, this.headerRowHeight);
        ctx.lineTo(this.headerColWidth - 0.5, height);
        ctx.stroke();

        ctx.font = '600 11px Segoe UI, sans-serif';
        ctx.textAlign = 'center';
        ctx.textBaseline = 'middle';

        for (let r = startRow; r <= endRow; r++) {
            if (sheet.isRowHidden(r)) continue;
            const rowY = this.headerRowHeight + sheet.getRowOffset(r) - this.scrollY;
            const rowH = sheet.getRowHeight(r);

            const isRowSelected = r >= sel.minRow && r <= sel.maxRow;
            if (isRowSelected) {
                ctx.fillStyle = '#e2e8f0';
                ctx.fillRect(0, rowY, this.headerColWidth - 1, rowH);
                ctx.fillStyle = '#107c41';
                ctx.fillRect(this.headerColWidth - 3, rowY, 2, rowH);
            }

            ctx.strokeStyle = '#cbd5e1';
            ctx.beginPath();
            ctx.moveTo(0, rowY + rowH - 0.5);
            ctx.lineTo(this.headerColWidth, rowY + rowH - 0.5);
            ctx.stroke();

            ctx.fillStyle = isRowSelected ? '#107c41' : '#64748b';
            ctx.fillText(String(r + 1), this.headerColWidth / 2, rowY + rowH / 2);
        }
        ctx.restore();

        // 4. Top-Left Corner
        ctx.fillStyle = '#f1f5f9';
        ctx.fillRect(0, 0, this.headerColWidth, this.headerRowHeight);
        ctx.strokeStyle = '#cbd5e1';
        ctx.strokeRect(0.5, 0.5, this.headerColWidth - 1, this.headerRowHeight - 1);

        ctx.fillStyle = '#94a3b8';
        ctx.beginPath();
        ctx.moveTo(this.headerColWidth - 8, this.headerRowHeight - 3);
        ctx.lineTo(this.headerColWidth - 3, this.headerRowHeight - 8);
        ctx.lineTo(this.headerColWidth - 3, this.headerRowHeight - 3);
        ctx.fill();
    }
}
