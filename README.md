## Apps Script Test Suite

This repository contains test helpers for Apps Script features in order to
effectively write tests for Apps Script libraries.

This is not feature complete, so we welcome Pull Requests that enhance
this test suite for new use cases.


### How to use

#### Chai example

```
import { mockAppsScript } from 'apps-script-test-suite/';
import { expect } from 'chai';

/**
 * A very simple example of a function to unit test.
 */
class DessertCounter {
    private readonly sheet: GoogleAppsScript.Spreadsheet.Sheet;
    private readonly rangeValue = 'A1:B3';

    constructor(public erase: boolean = false) {
       this.sheet = SpreadsheetApp.getActive().getSheetByName('Test Sheet'); 
    }

    init() {
        this.sheet.getRange(this.rangeValue).setValues([['name', 'cookies'], ['Mo', 1], ['Mary', 0]]);
    }

    increment(name: string) {
        const range = this.sheet.getRange(this.rangeValue);
        const values = range.getValues();
        const index = values.findIndex(row => row[0] === name);
        if (index < 0) {
            throw new Error(`No one named ${name} to give a cookie to!`);
        }
        ++values[index][1];
        range.setValues(values);
    }

    numCookies(name: string) {
        return this.sheet.getRange(1, 1, 2, 3).getValues().find(row => row[0] === name)[1];
    }
}


describe('Dessert Counter', function() {
    beforeEach(() => {
        // this creates the function stubs in the global space for testing Apps Script.
        mockAppsScript();

        // set up a sheet
        const sheet = SpreadsheetApp.getActive().insertSheet('Test Sheet');
        const range = sheet.getRange('A1:B3');
    });

    it('increments', function () {
        const dessertCounter = new DessetCounter();

        dessertCounter.increment('Mary');

        expect(
            SpreadsheetApp.getActive().getRangeByName('Test Sheet').getDataRange()
        ).toEqual(
            [['name', 'cookies'], ['Mo', 1], ['Mary', 1]]
        ); // passes
    });

    it('fails spectacularly when a name is not in the roster', function () {
        const dessertCounter = new DessetCounter();

        expect(() => dessertCounter.increment('Nobody')).to.throw(Error, 'No one named Nobody to give a cookie to!');
    });
});
```
