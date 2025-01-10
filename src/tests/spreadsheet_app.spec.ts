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

import { expect } from 'chai';
import { mockAppsScript } from '../index';
import { FakeSpreadsheetApp } from '../spreadsheet_app';
import { FakeSpreadsheet } from '../sheet';

describe('FakeSpreadsheetApp', function() {
  it('should get the active spreadsheet', function() {
    mockAppsScript();
    const app = new FakeSpreadsheetApp();
    expect(app.getActive()).to.be.instanceOf(FakeSpreadsheet);
  });

  it('should flush', function() {
    mockAppsScript();
    const app = new FakeSpreadsheetApp();
    expect(() => app.flush()).to.not.throw();
  });

  it('creates a new data validation', function() {
    mockAppsScript();
    const app = new FakeSpreadsheetApp();
    const result = app.newDataValidation();
    expect(result).to.exist;
    expect(result.build).to.be.a('function');  // Check if build is a function
  });

  it('creates a new text style', function() {
    mockAppsScript();
    const app = new FakeSpreadsheetApp();
    const result = app.newTextStyle();
    expect(result).to.exist;
    expect(result.setBold).to.be.a('function');
    expect(result.setItalic).to.be.a('function');
    expect(result.setFontSize).to.be.a('function');
    expect(result.setForegroundColor).to.be.a('function');
    expect(result.setFontFamily).to.be.a('function');
    expect(result.build).to.be.a('function');
  });

  it('creates a new rich text value', function() {
    mockAppsScript();
    const app = new FakeSpreadsheetApp();
    const result = app.newRichTextValue();
    expect(result).to.exist;
    expect(result.setText).to.be.a('function');
    expect(result.setTextStyle).to.be.a('function');
    expect(result.build).to.be.a('function');
    expect(result.getText).to.be.a('function');
  });


  it('should set and get text with rich text value', function() {
    mockAppsScript();
    const app = new FakeSpreadsheetApp();
    const richText = app.newRichTextValue();

    const testString = 'test string';
    richText.setText(testString);
    expect(richText.getText()).to.equal(testString);

    const builtRichText = richText.build();
    // The implementation of build just returns itself, so check for referential equality
    expect(builtRichText).to.equal(richText); 
  });

  it('banding theme constants should exist', function() {
    mockAppsScript();
    const app = new FakeSpreadsheetApp(); // Instantiate the class
    expect(app.BandingTheme.BLUE).to.equal(1);
  });

  it('Dimension constants should exist', function() {
    mockAppsScript();
    const app = new FakeSpreadsheetApp(); // Instantiate the class
    expect(app.Dimension.ROWS).to.equal(1);
    expect(app.Dimension.COLUMNS).to.equal(2);
  });
});
