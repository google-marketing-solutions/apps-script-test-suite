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
