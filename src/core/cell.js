// Cell data model & coordinate utilities
export function colToName(col) {
    let name = '';
    let temp = col;
    while (temp >= 0) {
        name = String.fromCharCode((temp % 26) + 65) + name;
        temp = Math.floor(temp / 26) - 1;
    }
    return name;
}

export function nameToCol(name) {
    name = name.toUpperCase();
    let col = 0;
    for (let i = 0; i < name.length; i++) {
        col = col * 26 + (name.charCodeAt(i) - 64);
    }
    return col - 1;
}

export function parseCoordinate(coordStr) {
    const match = coordStr.trim().toUpperCase().match(/^(\$?)([A-Z]+)(\$?)([0-9]+)$/);
    if (!match) return null;
    const colStr = match[2];
    const rowStr = match[4];
    return {
        col: nameToCol(colStr),
        row: parseInt(rowStr, 10) - 1,
        absCol: match[1] === '$',
        absRow: match[3] === '$'
    };
}

export function formatCoordinate(row, col) {
    return `${colToName(col)}${row + 1}`;
}

// Shifts relative cell coordinates within a formula when drag-filling
export function shiftFormula(formula, dRow, dCol) {
    if (!formula || !formula.startsWith('=')) return formula;

    // Matches cell coordinates like $A$1, A$1, $A1, A1, and ranges A1:B10
    const cellRegex = /(\$?)([A-Za-z]+)(\$?)([0-9]+)/g;

    return formula.replace(cellRegex, (match, colDollar, colStr, rowDollar, rowStr) => {
        const isAbsCol = colDollar === '$';
        const isAbsRow = rowDollar === '$';

        let col = nameToCol(colStr);
        let row = parseInt(rowStr, 10) - 1;

        if (!isAbsCol) {
            col = Math.max(0, col + dCol);
        }
        if (!isAbsRow) {
            row = Math.max(0, row + dRow);
        }

        return `${isAbsCol ? '$' : ''}${colToName(col)}${isAbsRow ? '$' : ''}${row + 1}`;
    });
}

export class Cell {
    constructor(row, col, value = '', formula = '') {
        this.row = row;
        this.col = col;
        this.value = value;         // Evaluated display/scalar value
        this.formula = formula;     // Raw formula string e.g. "=SUM(A1:A5)"
        this.style = {
            bold: false,
            italic: false,
            underline: false,
            fontSize: 12,
            fontFamily: 'Segoe UI, Tahoma, sans-serif',
            color: '#1e293b',
            background: '#ffffff',
            align: 'left',          // 'left', 'center', 'right'
            valign: 'middle',
            numFormat: 'general'    // 'general', 'currency', 'percent', 'number', 'text'
        };
    }

    clone() {
        const c = new Cell(this.row, this.col, this.value, this.formula);
        c.style = { ...this.style };
        return c;
    }
}
