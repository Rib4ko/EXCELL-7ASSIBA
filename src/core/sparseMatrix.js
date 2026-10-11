import { Cell } from './cell.js';

export class SparseMatrix {
    constructor() {
        // Map<row, Map<col, Cell>>
        this.rows = new Map();
    }

    getCell(row, col) {
        const rowMap = this.rows.get(row);
        if (!rowMap) return null;
        return rowMap.get(col) || null;
    }

    getOrCreateCell(row, col) {
        let rowMap = this.rows.get(row);
        if (!rowMap) {
            rowMap = new Map();
            this.rows.set(row, rowMap);
        }
        let cell = rowMap.get(col);
        if (!cell) {
            cell = new Cell(row, col);
            rowMap.set(col, cell);
        }
        return cell;
    }

    setCell(row, col, cell) {
        let rowMap = this.rows.get(row);
        if (!rowMap) {
            rowMap = new Map();
            this.rows.set(row, rowMap);
        }
        rowMap.set(col, cell);
    }

    deleteCell(row, col) {
        const rowMap = this.rows.get(row);
        if (!rowMap) return;
        rowMap.delete(col);
        if (rowMap.size === 0) {
            this.rows.delete(row);
        }
    }

    forEach(callback) {
        for (const [r, rowMap] of this.rows.entries()) {
            for (const [c, cell] of rowMap.entries()) {
                callback(cell, r, c);
            }
        }
    }

    getBounds() {
        let maxRow = 0;
        let maxCol = 0;
        for (const [r, rowMap] of this.rows.entries()) {
            if (r > maxRow) maxRow = r;
            for (const c of rowMap.keys()) {
                if (c > maxCol) maxCol = c;
            }
        }
        return { maxRow, maxCol };
    }

    clear() {
        this.rows.clear();
    }
}
