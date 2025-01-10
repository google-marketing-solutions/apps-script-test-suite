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

import { FakeSpreadsheet } from './sheet';
import { FakeNewDataValidation } from './common';

export class FakeSpreadsheetApp {
  private readonly fakeSpreadsheet = new FakeSpreadsheet();
  readonly BandingTheme = {
    BLUE: 1,
  };
  readonly Dimension = {
    ROWS: 1,
    COLUMNS: 2,
  };

  getActive() {
    return this.fakeSpreadsheet;
  }

  flush() {
    // do nothing
  }

  newDataValidation() {
    return new FakeNewDataValidation();
  }

  newTextStyle() {
    return new FakeTextStyle();
  }

  newRichTextValue() {
    return new FakeRichTextValue();
  }
}

class FakeTextStyle {
  setBold() {
    return this;
  }

  setItalic() {
    return this;
  }

  setFontSize() {
    return this;
  }

  setForegroundColor() {
    return this;
  }

  setFontFamily() {
    return this;
  }

  build() {
    return this;
  }
}

class FakeRichTextValue {
  text: string = '';

  setText(text: string) {
    this.text = text;
    return this;
  }

  setTextStyle() {
    return this;
  }

  build() {
    return this;
  }

  getText() {
    return this.text;
  }
}
