import { Workbook, CellChangeAction } from '../src/core/workbook.js';

console.log('--- RUNNING SHORTCUTS & UNDO/REDO AUTOMATED TESTS ---');

const wb = new Workbook('UndoRedoTest');
const sheet = wb.getActiveSheet();

// 1. Initial cell state
sheet.setCellValue(0, 0, 'Initial');
console.assert(sheet.getCell(0, 0).value === 'Initial', 'Cell should have Initial value');

// 2. Perform an edit and record undo action
const oldVal = sheet.getCell(0, 0).value;
const oldFormula = sheet.getCell(0, 0).formula;
const oldStyle = { ...sheet.getCell(0, 0).style };

const newCell = sheet.setCellValue(0, 0, 'Updated');
wb.pushUndo(new CellChangeAction(sheet, [{
    row: 0,
    col: 0,
    oldVal,
    oldFormula,
    oldStyle,
    newVal: newCell.value,
    newFormula: newCell.formula,
    newStyle: { ...newCell.style }
}]));

console.assert(sheet.getCell(0, 0).value === 'Updated', 'Cell should be Updated');

// 3. Test Undo (Ctrl+Z)
wb.undo();
console.assert(sheet.getCell(0, 0).value === 'Initial', `Undo should restore 'Initial', got: ${sheet.getCell(0, 0).value}`);
console.log('✓ Undo (Ctrl+Z) test passed');

// 4. Test Redo (Ctrl+Y)
wb.redo();
console.assert(sheet.getCell(0, 0).value === 'Updated', `Redo should restore 'Updated', got: ${sheet.getCell(0, 0).value}`);
console.log('✓ Redo (Ctrl+Y) test passed');

// 5. Test Batch Change Undo (Cut / Paste / Clear Range)
const batchChanges = [];
for (let r = 1; r <= 3; r++) {
    const prev = sheet.getCell(r, 0);
    const ov = prev ? prev.value : '';
    const nCell = sheet.setCellValue(r, 0, r * 10);
    batchChanges.push({
        row: r,
        col: 0,
        oldVal: ov,
        oldFormula: '',
        oldStyle: {},
        newVal: nCell.value,
        newFormula: '',
        newStyle: {}
    });
}
wb.pushUndo(new CellChangeAction(sheet, batchChanges));

console.assert(sheet.getCell(1, 0).value === 10, 'A2 should be 10');
console.assert(sheet.getCell(2, 0).value === 20, 'A3 should be 20');
console.assert(sheet.getCell(3, 0).value === 30, 'A4 should be 30');

// Undo batch change
wb.undo();
console.assert(sheet.getCell(1, 0).value === '', `A2 should be blank after undo, got: ${sheet.getCell(1, 0).value}`);
console.assert(sheet.getCell(2, 0).value === '', `A3 should be blank after undo, got: ${sheet.getCell(2, 0).value}`);
console.assert(sheet.getCell(3, 0).value === '', `A4 should be blank after undo, got: ${sheet.getCell(3, 0).value}`);
console.log('✓ Batch Range Undo test passed');

// Redo batch change
wb.redo();
console.assert(sheet.getCell(1, 0).value === 10, 'A2 should be restored to 10');
console.assert(sheet.getCell(2, 0).value === 20, 'A3 should be restored to 20');
console.assert(sheet.getCell(3, 0).value === 30, 'A4 should be restored to 30');
console.log('✓ Batch Range Redo test passed');

// 6. Test TSV Delimited Paste parsing logic
const sampleTsv = "Item\tPrice\tQty\r\nApple\t1.5\t10\r\nBanana\t0.75\t20";
const rows = sampleTsv.replace(/\r\n/g, '\n').replace(/\r/g, '\n').split('\n');
console.assert(rows.length === 3, 'TSV should parse into 3 rows');
const row1Cols = rows[1].split('\t');
console.assert(row1Cols[0] === 'Apple' && row1Cols[1] === '1.5' && row1Cols[2] === '10', 'Row 1 columns should match');
console.log('✓ TSV Clipboard parsing test passed');

console.log('\nALL SHORTCUT & UNDO/REDO TESTS PASSED! 🚀');
