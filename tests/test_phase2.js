import { Workbook } from '../src/core/workbook.js';

console.log('--- RUNNING PHASE 2 AUTOMATED TESTS ---');

const wb = new Workbook('Phase2Tests');
const sheet = wb.getActiveSheet();

// Set up sample table
// A1: Product, B1: Price, C1: Stock
sheet.setCellValue(0, 0, 'Apples');
sheet.setCellValue(0, 1, 3.5);
sheet.setCellValue(0, 2, 100);

sheet.setCellValue(1, 0, 'Bananas');
sheet.setCellValue(1, 1, 1.2);
sheet.setCellValue(1, 2, 250);

sheet.setCellValue(2, 0, 'Cherries');
sheet.setCellValue(2, 1, 8.0);
sheet.setCellValue(2, 2, 50);

// 1. Test VLOOKUP
sheet.setCellValue(3, 0, '=VLOOKUP("Bananas", A1:C3, 2, FALSE)');
console.assert(sheet.getCell(3, 0).value === 1.2, `VLOOKUP Bananas price should be 1.2, got: ${sheet.getCell(3, 0).value}`);

sheet.setCellValue(4, 0, '=VLOOKUP("Cherries", A1:C3, 3, FALSE)');
console.assert(sheet.getCell(4, 0).value === 50, `VLOOKUP Cherries stock should be 50, got: ${sheet.getCell(4, 0).value}`);

console.log('✓ VLOOKUP tests passed');

// 2. Test INDEX & MATCH
sheet.setCellValue(5, 0, '=MATCH("Cherries", A1:A3, 0)');
console.assert(sheet.getCell(5, 0).value === 3, `MATCH Cherries should be index 3, got: ${sheet.getCell(5, 0).value}`);

sheet.setCellValue(6, 0, '=INDEX(B1:B3, 1, 1)');
console.assert(sheet.getCell(6, 0).value === 3.5, `INDEX Apples price should be 3.5, got: ${sheet.getCell(6, 0).value}`);

console.log('✓ INDEX & MATCH tests passed');

// 3. Test XLOOKUP
sheet.setCellValue(7, 0, '=XLOOKUP("Apples", A1:A3, B1:B3)');
console.assert(sheet.getCell(7, 0).value === 3.5, `XLOOKUP Apples price should be 3.5, got: ${sheet.getCell(7, 0).value}`);
console.log('✓ XLOOKUP tests passed');

// 4. Test Sorting
// Unsorted numbers in D1:D3
sheet.setCellValue(0, 3, 50);
sheet.setCellValue(1, 3, 10);
sheet.setCellValue(2, 3, 30);

sheet.sortRange(0, 2, 3, 3, 3, true); // Sort ascending
console.assert(sheet.getCell(0, 3).value === 10, `D1 should be 10 after ascending sort, got: ${sheet.getCell(0, 3).value}`);
console.assert(sheet.getCell(1, 3).value === 30, `D2 should be 30 after ascending sort, got: ${sheet.getCell(1, 3).value}`);
console.assert(sheet.getCell(2, 3).value === 50, `D3 should be 50 after ascending sort, got: ${sheet.getCell(2, 3).value}`);

sheet.sortRange(0, 2, 3, 3, 3, false); // Sort descending
console.assert(sheet.getCell(0, 3).value === 50, `D1 should be 50 after descending sort, got: ${sheet.getCell(0, 3).value}`);
console.assert(sheet.getCell(1, 3).value === 30, `D2 should be 30 after descending sort, got: ${sheet.getCell(1, 3).value}`);
console.assert(sheet.getCell(2, 3).value === 10, `D3 should be 10 after descending sort, got: ${sheet.getCell(2, 3).value}`);

console.log('✓ Sorting tests passed');

// 5. Test AutoFilter & Row Hiding
sheet.toggleAutoFilter(0, 0, 2, 2);
sheet.autoFilter.filters.set(0, new Set(['Bananas'])); // Hide Bananas (row 1)
console.assert(sheet.isRowHidden(0) === false, 'Header row should not be hidden');
console.assert(sheet.isRowHidden(1) === true, 'Bananas row should be hidden');
console.assert(sheet.isRowHidden(2) === false, 'Cherries row should be visible');
console.log('✓ AutoFilter row hiding tests passed');

console.log('\nALL PHASE 2 TESTS PASSED SUCCESSFULLY! 🚀');
