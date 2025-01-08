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

import { expect } from 'chai';
import { mockAppsScript } from '../index';

describe('test ranges', function () {
  before(function () {
    mockAppsScript();
    const a = 'A'.charCodeAt(0);
    this.sheet = SpreadsheetApp.getActive().insertSheet(
      'new',
    ) as GoogleAppsScript.Spreadsheet.Sheet;
    const values = Array.from({ length: 4 })
      .fill('')
      .map((_, x) =>
        Array.from({ length: 26 }).map(
          (_, y) => `${String.fromCharCode(a + y)}${x + 1}`,
        ),
      );
    this.sheet.getRange('A1:Z4').setValues(values);
  });

  context('A1 notation', function () {
    it('grabs =1 cell', function () {
      const range: GoogleAppsScript.Spreadsheet.Range =
        this.sheet.getRange('A1');

      expect(range.getValues()).to.deep.equal([['A1']]);
    });

    it('grabs >1 cell', function () {
      const range: GoogleAppsScript.Spreadsheet.Range =
        this.sheet.getRange('A1:B2');

      expect(range.getValues()).to.deep.equal([
        ['A1', 'B1'],
        ['A2', 'B2'],
      ]);
    });

    it('grabs the right value from getValue', function () {
      const range: GoogleAppsScript.Spreadsheet.Range =
        this.sheet.getRange('Y1:Z2');
      expect(range.getValue()).to.deep.equal('Y1');
    });
  });

  context('non-A1 notation', function() {
    it('grabs =1 cell', function () {
      const range: GoogleAppsScript.Spreadsheet.Range =
        this.sheet.getRange(1, 1);

      expect(range.getValues()).to.deep.equal([['A1']]);
    });

    it('grabs >1 cell', function () {
      const range: GoogleAppsScript.Spreadsheet.Range =
        this.sheet.getRange(1, 1, 2, 2);

      expect(range.getValues()).to.deep.equal([
        ['A1', 'B1'],
        ['A2', 'B2'],
      ]);
    });

    it('grabs the right value from getValue', function () {
      const yZeroBase = 'Y'.charCodeAt(0) - 'A'.charCodeAt(0);
      const y = yZeroBase + 1;
      const range: GoogleAppsScript.Spreadsheet.Range = this.sheet.getRange(
        1,
        y,
        2,
        2,
      );
      expect(range.getValue()).to.deep.equal('Y1');
    });
  });

  context('merge', function() {
    it('clears merged ranges', function() {
      const range = this.sheet.getRange(1, 1, 2, 2);
      range.setValues([['A1', 'B1'], ['A2', 'B2']]);
      range.merge();
      expect(range.getValues()).to.deep.equal([['A1', ''], ['', '']])
    });
  });
});
