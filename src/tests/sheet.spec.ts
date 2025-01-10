import { expect } from 'chai';
import { mockAppsScript } from '../index';
import { FakeSpreadsheetApp } from '../spreadsheet_app';
import { FakeSheet, FakeSpreadsheet } from '../sheet';
import { FakeRange } from '../range';

const EMPTY_RANGE = [['']];

describe('FakeSpreadsheet', function() {
  let spreadsheet: FakeSpreadsheet;

  beforeEach(function() {
    mockAppsScript();
    spreadsheet = new FakeSpreadsheet();
  });

  it('should insert a new sheet with a given name', function() {
    const sheet = spreadsheet.insertSheet('New Sheet');
    expect(sheet).to.be.instanceOf(FakeSheet);
    expect(spreadsheet.getSheetByName('New Sheet')).to.equal(sheet);
  });

  it('should insert a new sheet with a default name if none is provided', function() {

    const sheet1 = spreadsheet.insertSheet();
    expect(sheet1.getRange('A1').getValue()).to.equal('');

    const sheet2 = spreadsheet.insertSheet('');
    expect(sheet2.getRange('A1').getValue()).to.equal('');

    expect(spreadsheet.getSheetByName('Sheet2')).to.equal(sheet1);
    expect(spreadsheet.getSheetByName('Sheet3')).to.equal(sheet2);
  });

  it('should retrieve a sheet by name', function() {
    const sheet = spreadsheet.insertSheet('My Sheet');
    const retrievedSheet = spreadsheet.getSheetByName('My Sheet');
    expect(retrievedSheet).to.equal(sheet);
  });

  it('should return undefined when retrieving a non-existent sheet', function() {
    const sheet = spreadsheet.getSheetByName('NonExistent Sheet');
    expect(sheet).to.be.undefined;
  });

  it('should get the active sheet', function() {
    const sheet1 = spreadsheet.getSheetByName('Sheet1');
    expect(spreadsheet.getActiveSheet()).to.equal(sheet1);

    spreadsheet.insertSheet('new sheet');
    expect(spreadsheet.getActiveSheet()).to.equal(sheet1);
  });


  it('should set and retrieve named ranges', function() {
    const sheet = spreadsheet.getActiveSheet();
    const range = sheet.getRange('A1:B2');
    spreadsheet.setNamedRange('MyRange', range);
    const namedRange = spreadsheet.getRangeByName('MyRange');

    expect(namedRange).to.equal(range);
    expect(namedRange?.getValue()).to.equal('');
  });
});


describe('FakeSheet', function() {
  let sheet: FakeSheet;

  beforeEach(function() {
    mockAppsScript();
    sheet = new FakeSheet();
  });

  it('should get a range by A1 notation', function() {
    const range = sheet.getRange('A1:B2');
    expect(range).to.be.instanceOf(FakeRange);
    expect(range.getValues()).to.deep.equal([['', ''], ['', '']]);
  });

  it('should get a range by row and column', function() {
    const range = sheet.getRange(1, 1, 2, 2);
    expect(range).to.be.instanceOf(FakeRange);
    expect(range.getValues()).to.deep.equal([['', ''], ['', '']]);
  });

  it('should throw error without column', function() {
    expect(() => sheet.getRange(1)).to.throw(Error, 'Required to include a column');
  });

  it('should get data range equal to [[\'\']] when empty', function() {
    const range = sheet.getDataRange();
    expect(range.getValues()).to.deep.equal([['']]);
  });

  it('should get the last column', function() {
    sheet.getRange('C1').setValue('Non-empty');
    expect(sheet.getLastColumn()).to.equal(3);
    sheet.getRange('D1').setValue('Non-empty');
    expect(sheet.getLastColumn()).to.equal(4);
  });

  it('should get the last row', function() {
    sheet.getRange('A3').setValue('Non-empty');
    expect(sheet.getLastRow()).to.equal(3);
    sheet.getRange('A4').setValue('Non-empty');
    expect(sheet.getLastRow()).to.equal(4);
  });

  it('should clear the sheet', function() {
    sheet.getRange('A1').setValue('value');
    sheet.clear();
    expect(sheet.getDataRange().getValues()).to.deep.equal(EMPTY_RANGE);
  });

  it('should get max rows and cols', function() {
    expect(sheet.getMaxRows()).to.equal(10000);
    expect(sheet.getMaxColumns()).to.equal(10000);
  });

  it('should delete rows', function() {
    sheet.getRange('A1').setValue('A1');
    sheet.getRange('A2').setValue('A2');
    sheet.deleteRows(1, 1);
    expect(sheet.getLastRow()).to.equal(1);
    expect(sheet.getRange('A1').getValue()).to.equal('A2');
  });

  it('should delete cols', function() {
    sheet.getRange('A1').setValue('A1');
    sheet.getRange('B1').setValue('B1');
    sheet.deleteColumns(1, 1); // delete col A
    expect(sheet.getLastColumn()).to.equal(1);
    expect(sheet.getRange('A1').getValue()).to.equal('B1');
  });

  it('should get bandings', function() {
    const bandings = sheet.getBandings();
    expect(bandings).to.deep.equal([]);
  });

  it('inserts checkboxes', function() {
    sheet.getRange('A1:B2').insertCheckboxes();
    expect(sheet.checkboxes).to.deep.equal({0: {0: true, 1: true}, 1: {0: true, 1: true}});
  });
});

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
    expect(result.build).to.exist;
  });


  it('creates a new text style', function() {
    mockAppsScript();
    const app = new FakeSpreadsheetApp();
    const result = app.newTextStyle();
    expect(result).to.exist;
    expect(result.setBold).to.exist;
    expect(result.setItalic).to.exist;
    expect(result.setFontSize).to.exist;
    expect(result.setForegroundColor).to.exist;
    expect(result.setFontFamily).to.exist;
    expect(result.build).to.exist;
  });


  it('creates a new rich text value', function() {
    mockAppsScript();
    const app = new FakeSpreadsheetApp();
    const result = app.newRichTextValue();
    expect(result).to.exist;
    expect(result.setText).to.exist;
    expect(result.setTextStyle).to.exist;
    expect(result.build).to.exist;
    expect(result.getText).to.exist;
  });

  it('banding theme constants should exist', function() {
    mockAppsScript();
    expect(SpreadsheetApp.BandingTheme.BLUE).to.equal(1);
  })

  it('Dimension constants should exist', function() {
    mockAppsScript();
    expect(SpreadsheetApp.Dimension.ROWS).to.equal(1);
    expect(SpreadsheetApp.Dimension.COLUMNS).to.equal(2);
  })
});
