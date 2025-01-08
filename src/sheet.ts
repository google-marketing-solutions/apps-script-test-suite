import { byA1Notation } from './common';
import { FakeRange } from './range';

export class FakeSheet {
  readonly cells: string[][] = Array.from({ length: 100 }).map(() =>
    Array.from({ length: 30 }),
  );
  readonly checkboxes: Record<number, Record<number, boolean>> = {};
  private readonly bandings: FakeBandings[] = [];

  getRange(a1Notation: string): FakeRange;
  getRange(row: number, column: number): FakeRange;
  getRange(row: number, column: number, numRows: number): FakeRange;
  getRange<T extends number | string>(
    row: number,
    column: number,
    numRows: number,
    numColumns: number,
  ): FakeRange;
  getRange(
    arg1: number | string,
    column?: number,
    numRows?: number,
    numColumns?: number,
  ) {
    if (typeof arg1 === 'string') {
      return byA1Notation(this, arg1);
    } else if (typeof arg1 === 'number' && !column) {
      throw new Error('Required to include a column');
    }
    return new FakeRange(this, arg1, column!, numRows, numColumns);
  }

  getDataRange(): FakeRange {
    return new FakeRange(this, 1, 1, this.getLastRow(), this.getLastColumn());
  }

  getLastColumn() {
    const lastIndexes = this.cells.map((row) =>
      row.findLastIndex((cell) => cell),
    );
    return Math.max(...lastIndexes) + 1;
  }

  clear(): FakeSheet {
    const emptyCells = this.cells.map((row) => row.map(() => ''));
    this.cells.splice(0, this.cells.length, ...emptyCells);
    return this;
  }

  getMaxRows() {
    return 10000;
  }

  getMaxColumns() {
    return 10000;
  }

  getLastRow() {
    return (
      this.cells.findLastIndex((row) => row.filter((cell) => cell).length > 0) +
      1
    );
  }

  deleteRows(start: number, end: number) {
    this.cells.splice(start, end);
  }

  deleteColumns(start: number, end: number) {
    for (const row of this.cells) {
      row.splice(start, end);
    }
  }

  getBandings(): FakeBandings[] {
    return this.bandings;
  }
}

export class FakeSpreadsheet {
  private static lastNum = 1;
  private readonly namedRange: Record<string, FakeRange> = {};
  private readonly sheets: Record<string, FakeSheet> = {
    Sheet1: new FakeSheet(),
  };
  private lastActive = 'Sheet1';

  insertSheet(sheetName: string) {
    const computedSheetName = sheetName || `Sheet${++FakeSpreadsheet.lastNum}`;
    this.sheets[computedSheetName] = new FakeSheet();
    return this.sheets[computedSheetName];
  }

  getRangeByName(rangeName: string) {
    return this.namedRange[rangeName];
  }

  setNamedRange(rangeName: string, range: FakeRange) {
    this.namedRange[rangeName] = range;
  }

  getSheetByName(sheetName: keyof typeof this.sheets) {
    return this.sheets[sheetName];
  }

  getActiveSheet() {
    return this.sheets[this.lastActive];
  }
}

class FakeBandings {}
