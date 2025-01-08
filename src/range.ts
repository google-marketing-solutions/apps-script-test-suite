/**
 * @license
 * Copyright 2024 Google LLC.
 *
 * Licensed under the Apache License, Version 2.0 (the "License");
 * you may not use this file except in compliance with the License.
 * You may obtain a copy of the License at
 *
 *     https://www.apache.org/licenses/LICENSE-2.0
 *
 * Unless required by applicable law or agreed to in writing, software
 * distributed under the License is distributed on an "AS IS" BASIS,
 * WITHOUT WARRANTIES OR CONDITIONS OF ANY KIND, either express or implied.
 * See the License for the specific language governing permissions and
 * limitations under the License.
 */

import { FakeNewDataValidation, FakeFilter } from './common';
import { FakeSheet } from './sheet';

export class FakeRange {
  private readonly arrayRange: string[][];
  validation?: FakeNewDataValidation;

  constructor(
    private readonly sheet: FakeSheet,
    private readonly row: number,
    private readonly column: number,
    private readonly numRows = 1,
    private readonly numColumns = 1,
  ) {
    this.arrayRange = this.initializeSheet();
  }

  private getRangeComponent() {
    return this.sheet.cells.slice(this.row - 1, this.row - 1 + this.numRows).map(
      (row) => row.slice(this.column - 1, this.column - 1 + this.numColumns),
    );
  }

  initializeSheet() {
    return this.getRangeComponent().map((columns) =>
      columns
        .slice(this.column - 1, this.column - 1 + this.numColumns)
        .map((cell) => cell ?? ''),
    );
  }

  getValues() {
    return this.arrayRange;
  }

  getValue() {
    return this.arrayRange[0][0];
  }

  setValues(range: string[][]) {
    if (this.numRows !== range.length) {
      throw new Error(
        `Invalid row length: ${this.arrayRange.length} vs ${range.length}`,
      );
    }
    for (const [i, row] of range.entries()) {
      if (row.length === this.numColumns) {
        this.arrayRange[i] = row;
      } else {
        throw new Error('Invalid column length');
      }
    }
    this.sheet.cells.splice(
      this.row - 1,
      this.arrayRange.length,
      ...this.arrayRange.map((row) => {
        const newArr = [
          ...row,
          ...Array.from<string>({
            length: this.sheet.cells[0].length - row.length,
          }).fill(''),
        ];
        return newArr;
      }),
    );
    return this;
  }

  setValue(value: string) {
    this.arrayRange[0][0] = value;
    this.sheet.cells[this.row][this.column] = value;
    return this;
  }

  clearDataValidations(): FakeRange {
    return this;
  }

  setDataValidation(validation: FakeNewDataValidation) {
    this.validation = validation;
    return this;
  }

  breakApart() {
    return this;
  }

  merge() {
    const value = this.getValue();
    this.setValues(this.arrayRange.map(row => row.map(cell => '')));
    this.setValue(value);
    // TODO add more here to get merging to work or scrap all of this?
    return this;
  }

  applyRowBanding() {
    return this;
  }

  insertCells() {
    return this;
  }

  setRichTextValue() {
    return this;
  }

  getFilter() {
    return new FakeFilter();
  }

  insertCheckboxes() {
    for (let r = this.row; r < this.row + this.numRows; r++) {
      for (let c = this.column; c < this.column + this.numColumns; c++) {
        if (!this.sheet.checkboxes[r]) {
          this.sheet.checkboxes[r] = {};
        }
        this.sheet.checkboxes[r][c] = true;
      }
    }
  }
}
