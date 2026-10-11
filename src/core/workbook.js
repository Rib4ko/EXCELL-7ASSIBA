import { Worksheet } from './worksheet.js';

export class CellChangeAction {
    constructor(sheet, changes) {
        this.sheet = sheet;
        this.changes = changes;
    }

    undo() {
        for (const ch of this.changes) {
            const cell = this.sheet.getOrCreateCell(ch.row, ch.col);
            cell.formula = ch.oldFormula !== undefined ? ch.oldFormula : '';
            cell.value = ch.oldVal !== undefined ? ch.oldVal : '';
            if (ch.oldStyle !== undefined) {
                cell.style = { ...ch.oldStyle };
            }
            if (cell.formula) {
                this.sheet.formulaEngine.evaluateCell(cell);
            }
            this.sheet.formulaEngine.recalculateFrom(ch.row, ch.col);
        }
    }

    redo() {
        for (const ch of this.changes) {
            const cell = this.sheet.getOrCreateCell(ch.row, ch.col);
            cell.formula = ch.newFormula !== undefined ? ch.newFormula : '';
            cell.value = ch.newVal !== undefined ? ch.newVal : '';
            if (ch.newStyle !== undefined) {
                cell.style = { ...ch.newStyle };
            }
            if (cell.formula) {
                this.sheet.formulaEngine.evaluateCell(cell);
            }
            this.sheet.formulaEngine.recalculateFrom(ch.row, ch.col);
        }
    }
}

export class Workbook {
    constructor(title = 'Book1') {
        this.title = title;
        this.sheets = [];
        this.activeSheetIndex = 0;

        // Undo / Redo history stack
        this.undoStack = [];
        this.redoStack = [];

        // Initial default sheet
        this.addSheet('Sheet1');
    }

    getActiveSheet() {
        return this.sheets[this.activeSheetIndex] || this.sheets[0];
    }

    addSheet(name) {
        let sheetName = name;
        if (!sheetName) {
            sheetName = `Sheet${this.sheets.length + 1}`;
        }
        const sheet = new Worksheet(sheetName, this);
        this.sheets.push(sheet);
        return sheet;
    }

    removeSheet(index) {
        if (this.sheets.length <= 1) return false;
        this.sheets.splice(index, 1);
        if (this.activeSheetIndex >= this.sheets.length) {
            this.activeSheetIndex = this.sheets.length - 1;
        }
        return true;
    }

    setActiveSheet(index) {
        if (index >= 0 && index < this.sheets.length) {
            this.activeSheetIndex = index;
            return true;
        }
        return false;
    }

    pushUndo(action) {
        this.undoStack.push(action);
        if (this.undoStack.length > 100) {
            this.undoStack.shift();
        }
        this.redoStack = [];
    }

    undo() {
        if (this.undoStack.length === 0) return null;
        const action = this.undoStack.pop();
        if (typeof action.undo === 'function') {
            action.undo();
        } else if (typeof action.executeUndo === 'function') {
            const inverse = action.executeUndo();
            if (inverse) {
                this.redoStack.push(inverse);
                return action;
            }
        }
        this.redoStack.push(action);
        return action;
    }

    redo() {
        if (this.redoStack.length === 0) return null;
        const action = this.redoStack.pop();
        if (typeof action.redo === 'function') {
            action.redo();
        } else if (typeof action.executeRedo === 'function') {
            const inverse = action.executeRedo();
            if (inverse) {
                this.undoStack.push(inverse);
                return action;
            }
        }
        this.undoStack.push(action);
        return action;
    }

    toJSON() {
        return {
            title: this.title,
            activeSheetIndex: this.activeSheetIndex,
            sheets: this.sheets.map(s => s.toJSON())
        };
    }

    fromJSON(obj) {
        this.title = obj.title || 'Book1';
        this.sheets = [];
        if (Array.isArray(obj.sheets)) {
            for (const sObj of obj.sheets) {
                const sheet = new Worksheet(sObj.name, this);
                sheet.fromJSON(sObj);
                this.sheets.push(sheet);
            }
        }
        this.activeSheetIndex = obj.activeSheetIndex || 0;
    }

    toCSV(sheetIndex = this.activeSheetIndex) {
        const sheet = this.sheets[sheetIndex];
        if (!sheet) return '';
        const bounds = sheet.data.getBounds();
        const rows = [];

        for (let r = 0; r <= bounds.maxRow; r++) {
            const cols = [];
            for (let c = 0; c <= bounds.maxCol; c++) {
                const cell = sheet.getCell(r, c);
                let val = cell ? (cell.value !== undefined ? String(cell.value) : '') : '';
                if (val.includes(',') || val.includes('"') || val.includes('\n')) {
                    val = `"${val.replace(/"/g, '""')}"`;
                }
                cols.push(val);
            }
            rows.push(cols.join(','));
        }
        return rows.join('\r\n');
    }

    fromCSV(csvText, sheetName = 'Imported') {
        const sheet = new Worksheet(sheetName, this);
        const lines = csvText.split(/\r?\n/);

        lines.forEach((line, r) => {
            if (!line.trim()) return;
            const rowValues = [];
            let inQuotes = false;
            let current = '';

            for (let i = 0; i < line.length; i++) {
                const char = line[i];
                if (char === '"') {
                    if (inQuotes && line[i + 1] === '"') {
                        current += '"';
                        i++;
                    } else {
                        inQuotes = !inQuotes;
                    }
                } else if (char === ',' && !inQuotes) {
                    rowValues.push(current);
                    current = '';
                } else {
                    current += char;
                }
            }
            rowValues.push(current);

            rowValues.forEach((val, c) => {
                sheet.setCellValue(r, c, val);
            });
        });

        this.sheets.push(sheet);
        this.activeSheetIndex = this.sheets.length - 1;
        return sheet;
    }

    // Export entire workbook to native Microsoft Excel .XLSX via ExcelJS
    async toXLSXBuffer(ExcelJS) {
        const wb = new ExcelJS.Workbook();
        wb.creator = 'ApexSpreadsheet';
        wb.lastModifiedBy = 'ApexSpreadsheet';
        wb.created = new Date();

        for (const sheet of this.sheets) {
            const ws = wb.addWorksheet(sheet.name);

            // Set column widths
            for (const [col, width] of sheet.colWidths.entries()) {
                const colLetter = ws.getColumn(col + 1);
                colLetter.width = Math.round(width / 7.5);
            }

            // Export cell values, formulas, and styles
            sheet.data.forEach((cell, r, c) => {
                const excelCell = ws.getCell(r + 1, c + 1);

                if (cell.formula && cell.formula.startsWith('=')) {
                    excelCell.value = {
                        formula: cell.formula.substring(1),
                        result: cell.value
                    };
                } else {
                    excelCell.value = cell.value;
                }

                // Apply style
                if (cell.style) {
                    const font = {};
                    if (cell.style.bold) font.bold = true;
                    if (cell.style.italic) font.italic = true;
                    if (cell.style.underline) font.underline = true;
                    if (cell.style.fontSize) font.size = cell.style.fontSize;
                    if (cell.style.color) font.color = { argb: 'FF' + cell.style.color.replace('#', '') };
                    excelCell.font = font;

                    if (cell.style.align) {
                        excelCell.alignment = { horizontal: cell.style.align };
                    }

                    if (cell.style.background && cell.style.background !== '#ffffff') {
                        excelCell.fill = {
                            type: 'pattern',
                            pattern: 'solid',
                            fgColor: { argb: 'FF' + cell.style.background.replace('#', '') }
                        };
                    }

                    if (cell.style.numFormat === 'currency') {
                        excelCell.numFmt = '$#,##0.00';
                    } else if (cell.style.numFormat === 'percent') {
                        excelCell.numFmt = '0.0%';
                    }
                }
            });

            // Export Merges
            for (const m of sheet.merges) {
                ws.mergeCells(m.r1 + 1, m.c1 + 1, m.r2 + 1, m.c2 + 1);
            }
        }

        return await wb.xlsx.writeBuffer();
    }

    // Import workbook from .XLSX buffer
    async fromXLSXBuffer(buffer, ExcelJS) {
        const wb = new ExcelJS.Workbook();
        await wb.xlsx.load(buffer);

        this.sheets = [];
        wb.eachSheet((ws, sheetId) => {
            const sheet = new Worksheet(ws.name, this);

            // Import Column widths
            ws.columns.forEach((col, idx) => {
                if (col.width) {
                    sheet.setColWidth(idx, Math.round(col.width * 7.5));
                }
            });

            // Import Merges
            if (ws.model && ws.model.merges) {
                for (const mergeStr of ws.model.merges) {
                    const parts = mergeStr.split(':');
                    if (parts.length === 2) {
                        const mStart = ws.getCell(parts[0]);
                        const mEnd = ws.getCell(parts[1]);
                        sheet.mergeCells(mStart.row - 1, mStart.col - 1, mEnd.row - 1, mEnd.col - 1);
                    }
                }
            }

            // Import Cells
            ws.eachRow({ includeEmpty: false }, (row, rowNumber) => {
                row.eachCell({ includeEmpty: false }, (cell, colNumber) => {
                    const r = rowNumber - 1;
                    const c = colNumber - 1;

                    let rawVal = cell.value;
                    if (rawVal && typeof rawVal === 'object' && rawVal.formula) {
                        rawVal = '=' + rawVal.formula;
                    }

                    const appCell = sheet.setCellValue(r, c, rawVal);

                    // Import style
                    if (cell.font) {
                        if (cell.font.bold) appCell.style.bold = true;
                        if (cell.font.italic) appCell.style.italic = true;
                        if (cell.font.underline) appCell.style.underline = true;
                        if (cell.font.size) appCell.style.fontSize = cell.font.size;
                    }
                    if (cell.alignment && cell.alignment.horizontal) {
                        appCell.style.align = cell.alignment.horizontal;
                    }
                    if (cell.fill && cell.fill.fgColor && cell.fill.fgColor.argb) {
                        appCell.style.background = '#' + cell.fill.fgColor.argb.substring(2);
                    }
                });
            });

            this.sheets.push(sheet);
        });

        this.activeSheetIndex = 0;
    }
}
