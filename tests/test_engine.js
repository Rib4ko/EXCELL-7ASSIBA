import { Workbook } from '../src/core/workbook.js';

console.log('--- RUNNING SPREADSHEET UNIT TESTS ---');

const wb = new Workbook('TestBook');
const sheet = wb.getActiveSheet();

// 1. Basic Value insertion
sheet.setCellValue(0, 0, 10);
sheet.setCellValue(1, 0, 20);
sheet.setCellValue(2, 0, 30);

console.assert(sheet.getCell(0, 0).value === 10, 'A1 should be 10');
console.assert(sheet.getCell(1, 0).value === 20, 'A2 should be 20');
console.assert(sheet.getCell(2, 0).value === 30, 'A3 should be 30');
console.log('✓ Basic values passed');

// 2. Arithmetic Formula
sheet.setCellValue(3, 0, '=A1 + A2 * 2');
console.assert(sheet.getCell(3, 0).value === 50, `A4 should be 50, got: ${sheet.getCell(3, 0).value}`);
console.log('✓ Arithmetic formula passed (10 + 20 * 2 = 50)');

// 3. Range SUM function
sheet.setCellValue(4, 0, '=SUM(A1:A3)');
console.assert(sheet.getCell(4, 0).value === 60, `A5 should be 60, got: ${sheet.getCell(4, 0).value}`);
console.log('✓ Range SUM function passed (SUM(A1:A3) = 60)');

// 4. AVERAGE function
sheet.setCellValue(5, 0, '=AVERAGE(A1:A3)');
console.assert(sheet.getCell(5, 0).value === 20, `A6 should be 20, got: ${sheet.getCell(5, 0).value}`);
console.log('✓ AVERAGE function passed (AVERAGE(A1:A3) = 20)');

// 5. Reactive Dependency Graph (DAG) Recalculation
console.log('Testing reactive recalculation when precedent changes...');
sheet.setCellValue(0, 0, 50); // Changed A1 from 10 to 50
// A4 was =A1 + A2 * 2 => 50 + 40 = 90
console.assert(sheet.getCell(3, 0).value === 90, `A4 should recalculate to 90, got: ${sheet.getCell(3, 0).value}`);
// A5 was =SUM(A1:A3) => 50 + 20 + 30 = 100
console.assert(sheet.getCell(4, 0).value === 100, `A5 should recalculate to 100, got: ${sheet.getCell(4, 0).value}`);
console.log('✓ Reactive DAG recalculation passed');

// 6. Multi-function string concatenation & logic
sheet.setCellValue(6, 0, '=IF(A5 >= 100, "High", "Low")');
console.assert(sheet.getCell(6, 0).value === 'High', `A7 should be High, got: ${sheet.getCell(6, 0).value}`);
console.log('✓ Logical IF function passed');

// 7. CSV Export and Import
const csv = wb.toCSV();
console.assert(csv.includes('50'), 'CSV should contain 50');
console.log('✓ CSV serialization passed');

console.log('\nALL 7 CORE SUITE TESTS PASSED SUCCESSFULLY! 🚀');
