import { SparseMatrix } from './sparseMatrix.js';
import { FormulaEngine } from './formula.js';
import { colToName, parseCoordinate } from './cell.js';

export class Worksheet {
    constructor(name, workbook) {
        this.name = name;
        this.workbook = workbook;
        this.data = new SparseMatrix();
        this.formulaEngine = new FormulaEngine(() => this);

        // Column and Row dimensions
        this.defaultColWidth = 100;
        this.defaultRowHeight = 24;
        this.colWidths = new Map();  // col -> width
        this.rowHeights = new Map(); // row -> height

        // Merged ranges: Array<{ r1, c1, r2, c2 }>
        this.merges = [];

        // Freeze panes
        this.freezeRow = 0;
        this.freezeCol = 0;

        // Auto-filter: { r1, c1, r2, c2, filters: Map<col, Set<hiddenValues>> }
        this.autoFilter = null;
    }

    getCell(row, col) {
        return this.data.getCell(row, col);
    }

    getOrCreateCell(row, col) {
        return this.data.getOrCreateCell(row, col);
    }

    setCellValue(row, col, rawInput) {
        const cell = this.getOrCreateCell(row, col);
        const str = String(rawInput !== undefined && rawInput !== null ? rawInput : '');

        if (str.startsWith('=')) {
            cell.formula = str;
            this.formulaEngine.evaluateCell(cell);
        } else {
            cell.formula = '';
            const num = Number(str);
            if (str !== '' && !isNaN(num)) {
                cell.value = num;
            } else if (str.toUpperCase() === 'TRUE') {
                cell.value = true;
            } else if (str.toUpperCase() === 'FALSE') {
                cell.value = false;
            } else {
                cell.value = str;
            }
        }

        this.formulaEngine.recalculateFrom(row, col);
        return cell;
    }

    setCellStyle(row, col, styleProps) {
        const cell = this.getOrCreateCell(row, col);
        cell.style = { ...cell.style, ...styleProps };
        return cell;
    }

    // Merge & Center support
    mergeCells(r1, c1, r2, c2) {
        const minRow = Math.min(r1, r2);
        const maxRow = Math.max(r1, r2);
        const minCol = Math.min(c1, c2);
        const maxCol = Math.max(c1, c2);

        if (minRow === maxRow && minCol === maxCol) return;

        this.unmergeCells(minRow, minCol);
        this.merges.push({ r1: minRow, c1: minCol, r2: maxRow, c2: maxCol });
        this.setCellStyle(minRow, minCol, { align: 'center' });
    }

    unmergeCells(row, col) {
        this.merges = this.merges.filter(m => {
            return !(row >= m.r1 && row <= m.r2 && col >= m.c1 && col <= m.c2);
        });
    }

    getMerge(row, col) {
        for (const m of this.merges) {
            if (row >= m.r1 && row <= m.r2 && col >= m.c1 && col <= m.c2) {
                return m;
            }
        }
        return null;
    }

    isMergeTopLeft(row, col) {
        const m = this.getMerge(row, col);
        return m ? (m.r1 === row && m.c1 === col) : true;
    }

    getColWidth(col) {
        return this.colWidths.get(col) || this.defaultColWidth;
    }

    setColWidth(col, width) {
        this.colWidths.set(col, Math.max(20, width));
    }

    getRowHeight(row) {
        // If row is hidden by filter, height is 0
        if (this.isRowHidden(row)) return 0;
        return this.rowHeights.get(row) || this.defaultRowHeight;
    }

    setRowHeight(row, height) {
        this.rowHeights.set(row, Math.max(16, height));
    }

    isRowHidden(row) {
        if (!this.autoFilter) return false;
        if (row <= this.autoFilter.r1) return false; // Header is never hidden

        for (const [col, hiddenVals] of this.autoFilter.filters.entries()) {
            const cell = this.getCell(row, col);
            const val = cell ? String(cell.value) : '';
            if (hiddenVals.has(val)) return true;
        }
        return false;
    }

    getColOffset(col) {
        let x = 0;
        for (let c = 0; c < col; c++) {
            x += this.getColWidth(c);
        }
        return x;
    }

    getRowOffset(row) {
        let y = 0;
        for (let r = 0; r < row; r++) {
            y += this.getRowHeight(r);
        }
        return y;
    }

    getColAtX(x) {
        let cur = 0;
        let c = 0;
        while (true) {
            const w = this.getColWidth(c);
            if (x < cur + w) return c;
            cur += w;
            c++;
            if (c > 16384) return 16384;
        }
    }

    getRowAtY(y) {
        let cur = 0;
        let r = 0;
        while (true) {
            const h = this.getRowHeight(r);
            if (y < cur + h) return r;
            cur += h;
            r++;
            if (r > 1048576) return 1048576;
        }
    }

    autoFitColumn(col, measureTextFn) {
        let maxTextWidth = 40;
        this.data.forEach((cell, r, c) => {
            if (c === col && cell && cell.value !== undefined && cell.value !== null && cell.value !== '') {
                const text = String(cell.value);
                const measured = measureTextFn(text, cell.style);
                if (measured > maxTextWidth) {
                    maxTextWidth = measured;
                }
            }
        });
        this.setColWidth(col, Math.min(500, Math.ceil(maxTextWidth + 24)));
    }

    // Sort range by given column
    sortRange(minRow, maxRow, minCol, maxCol, sortCol, ascending = true) {
        const rowsData = [];

        for (let r = minRow; r <= maxRow; r++) {
            const rowCells = [];
            for (let c = minCol; c <= maxCol; c++) {
                const cell = this.getCell(r, c);
                rowCells.push(cell ? cell.clone() : null);
            }
            rowsData.push({ originalRow: r, cells: rowCells });
        }

        const colOffset = sortCol - minCol;

        rowsData.sort((a, b) => {
            const valA = a.cells[colOffset]?.value;
            const valB = b.cells[colOffset]?.value;

            if (valA === valB) return 0;
            if (valA === undefined || valA === null || valA === '') return 1;
            if (valB === undefined || valB === null || valB === '') return -1;

            if (typeof valA === 'number' && typeof valB === 'number') {
                return ascending ? valA - valB : valB - valA;
            }

            const strA = String(valA).toLowerCase();
            const strB = String(valB).toLowerCase();
            return ascending ? strA.localeCompare(strB) : strB.localeCompare(strA);
        });

        // Write sorted rows back
        rowsData.forEach((rowItem, idx) => {
            const targetRow = minRow + idx;
            rowItem.cells.forEach((cell, colIdx) => {
                const targetCol = minCol + colIdx;
                if (cell) {
                    this.setCellValue(targetRow, targetCol, cell.formula || cell.value);
                    if (cell.style) this.setCellStyle(targetRow, targetCol, cell.style);
                } else {
                    this.data.deleteCell(targetRow, targetCol);
                }
            });
        });
    }

    // AutoFilter controls
    toggleAutoFilter(r1, c1, r2, c2) {
        if (this.autoFilter) {
            this.autoFilter = null;
        } else {
            this.autoFilter = {
                r1: Math.min(r1, r2),
                c1: Math.min(c1, c2),
                r2: Math.max(r1, r2),
                c2: Math.max(c1, c2),
                filters: new Map() // col -> Set<hiddenValues>
            };
        }
    }

    toJSON() {
        const cells = [];
        this.data.forEach((cell, r, c) => {
            cells.push({
                row: r,
                col: c,
                value: cell.value,
                formula: cell.formula,
                style: cell.style
            });
        });

        return {
            name: this.name,
            cells,
            merges: this.merges,
            colWidths: Array.from(this.colWidths.entries()),
            rowHeights: Array.from(this.rowHeights.entries()),
            freezeRow: this.freezeRow,
            freezeCol: this.freezeCol
        };
    }

    fromJSON(obj) {
        this.name = obj.name;
        this.data.clear();
        this.merges = obj.merges || [];
        this.colWidths = new Map(obj.colWidths || []);
        this.rowHeights = new Map(obj.rowHeights || []);
        this.freezeRow = obj.freezeRow || 0;
        this.freezeCol = obj.freezeCol || 0;

        if (Array.isArray(obj.cells)) {
            for (const item of obj.cells) {
                const cell = this.getOrCreateCell(item.row, item.col);
                cell.value = item.value;
                cell.formula = item.formula || '';
                if (item.style) cell.style = { ...cell.style, ...item.style };
            }
        }

        this.data.forEach((cell) => {
            if (cell.formula) {
                this.formulaEngine.evaluateCell(cell);
            }
        });
    }
}
