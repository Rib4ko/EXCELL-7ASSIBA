import { shiftFormula } from '../src/core/cell.js';
import { Workbook } from '../src/core/workbook.js';

console.log('--- RUNNING PHASE 1 AUTOMATED TESTS ---');

// 1. Test Formula Relative Reference Shifting
const f1 = '=A1 + B1';
const shiftedDown = shiftFormula(f1, 1, 0);
console.assert(shiftedDown === '=A2 + B2', `Expected =A2 + B2, got: ${shiftedDown}`);

const shiftedRight = shiftFormula(f1, 0, 1);
console.assert(shiftedRight === '=B1 + C1', `Expected =B1 + C1, got: ${shiftedRight}`);

// Absolute references should stay locked
const f2 = '=$A$1 + B1';
const shiftedAbs = shiftFormula(f2, 5, 2);
console.assert(shiftedAbs === '=$A$1 + D6', `Expected =$A$1 + D6, got: ${shiftedAbs}`);

// Range shifting
const f3 = '=SUM(A1:A5)';
const shiftedRange = shiftFormula(f3, 2, 0);
console.assert(shiftedRange === '=SUM(A3:A7)', `Expected =SUM(A3:A7), got: ${shiftedRange}`);

console.log('✓ Formula reference shifting tests passed');

// 2. Test Merge & Center Logic
const wb = new Workbook('MergeTest');
const sheet = wb.getActiveSheet();

sheet.setCellValue(0, 0, 'Merged Header');
sheet.mergeCells(0, 0, 0, 3); // A1:D1

console.assert(sheet.merges.length === 1, 'Should have 1 merged range');
const merge = sheet.getMerge(0, 2); // C1
console.assert(merge !== null && merge.r1 === 0 && merge.c2 === 3, 'C1 should be inside merge A1:D1');
console.assert(sheet.isMergeTopLeft(0, 0) === true, 'A1 is top-left master');
console.assert(sheet.isMergeTopLeft(0, 1) === false, 'B1 is not top-left master');

sheet.unmergeCells(0, 1);
console.assert(sheet.merges.length === 0, 'Unmerge should remove the merged range');
console.log('✓ Merge & Center logic tests passed');

console.log('\nALL PHASE 1 CORE TESTS PASSED! 🚀');
