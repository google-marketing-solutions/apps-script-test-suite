/**
 * @license
 * Copyright 2025 Google LLC.
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

import { FakeRange } from './range';
import { FakeSheet } from './sheet';

export class FakeNewDataValidation {
  requireFormulaSatisfied() {
    return this;
  }

  build() {}
}

export class FakeFilter {}

export function byA1Notation(sheet: FakeSheet, a1Notation: string) {
  const parts = a1Notation.split(':');
  const { row: row1, column: column1 } = a1NotationToRowColumn(parts[0], true);
  const { row: row2, column: column2 } = a1NotationToRowColumn(
    parts[1] || parts[0],
    false,
  );
  return new FakeRange(
    sheet,
    row1,
    column1,
    row2 - row1 + 1,
    column2 - column1 + 1,
  );
}

export function a1NotationToRowColumn(a1Notation: string, start = true) {
  const a = 'A'.charCodeAt(0);
  let column = 0;
  let i: number;
  const parts = a1Notation.toUpperCase().match(/([A-Z]+)([1-9]\d*)?/);
  if (!parts) {
    throw new Error('Invalid A1 notation');
  }
  const letters = parts[1];
  const row: number = parts[2]
    ? Number.parseInt(parts[2], 10)
    : start
      ? 1
      : 100_000;
  for (i = 0; i < letters.length; i++) {
    column += letters.charCodeAt(i) - a + 1;
  }

  return { row, column };
}
