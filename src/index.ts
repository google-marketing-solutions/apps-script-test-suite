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

/**
 * @fileoverview Mocked classes for Apps Script to help with unit tests.
 */

import { FakeSpreadsheetApp } from './spreadsheet_app';
import BigQuery = GoogleAppsScript.BigQuery;
import Properties = GoogleAppsScript.Properties;
import Cache = GoogleAppsScript.Cache;
import Mail = GoogleAppsScript.Mail;
import { FakeSpreadsheet } from './sheet';

class FakeHtmlService {
  createTemplateFromFile() {
    throw new Error('Not implemented. Stub me.');
  }
}

/**
 * Used to generate mocks for Apps Script libraries that are used in this
 * client.
 */
export function mockAppsScript() {
  (globalThis.MailApp as unknown as FakeMailApp) = new FakeMailApp();
  (globalThis.PropertiesService as unknown as FakePropertiesService) =
    new FakePropertiesService();
  (globalThis.CacheService as unknown as FakeCacheService) =
    new FakeCacheService();
  (globalThis.Utilities as unknown as FakeUtilitiesService) =
    new FakeUtilitiesService();
  (globalThis.SpreadsheetApp as unknown as FakeSpreadsheetApp) =
    new FakeSpreadsheetApp();
  (globalThis.ScriptApp as unknown as FakeScriptApp) = new FakeScriptApp();
  (globalThis.HtmlService as unknown as FakeHtmlService) =
    new FakeHtmlService();
  (globalThis.UrlFetchApp as unknown as FakeUrlFetchApp) =
    new FakeUrlFetchApp();
  (globalThis.Drive as unknown as FakeDrive) = new FakeDrive();
  (globalThis.BigQuery as unknown as FakeBigQuery) = new FakeBigQuery();
}

class FakeUrlFetchApp {
  fetch() {
    throw new Error('Not implemented. Mock me.');
  }
}

/**
 * A return value for a FakeUrlFetchApp
 */
export function generateFakeHttpResponse(args: { contentText: string }) {
  return {
    getContentText() {
      return args.contentText;
    },
  } as unknown as GoogleAppsScript.URL_Fetch.HTTPResponse;
}

class FakeScriptApp {
  getOAuthToken() {
    return 'token';
  }
}

/**
 * Retrieves emails
 */
export function getEmails() {
  return (MailApp as unknown as FakeMailApp).getEmails();
}

class FakeMailApp {
  private readonly emails: Mail.MailAdvancedParameters[] = [];

  sendEmail(message: Mail.MailAdvancedParameters) {
    this.emails.push(message);
  }

  getEmails() {
    return this.emails;
  }
}

class PropertyStub implements Properties.Properties {
  private storage: { [key: string]: string } = {};

  getProperties() {
    return this.storage;
  }

  setProperties(properties: { [key: string]: string }) {
    this.storage = properties;
    return this;
  }

  getProperty(key: string) {
    return this.storage[key];
  }

  setProperty(key: string, value: string) {
    this.storage[key] = value;
    return this;
  }

  deleteProperty(key: string): Properties.Properties {
    delete this.storage[key];
    return this;
  }

  deleteAllProperties(): Properties.Properties {
    this.storage = {};
    return this;
  }

  getKeys(): string[] {
    return Object.keys(this.storage);
  }
}

class CacheStub implements Cache.Cache {
  private cache: Record<string, string> = {};
  expirationInSeconds: number | undefined;

  getAll(): Record<string, string> {
    throw new Error('Method not implemented.');
  }
  putAll(values: Record<string, string>): void;
  putAll(values: Record<string, string>, expirationInSeconds: number): void;
  putAll(values: Record<string, string>, expirationInSeconds?: number): void {
    this.cache = values;
    this.expirationInSeconds = expirationInSeconds;
  }
  remove(key: string): void {
    delete this.cache[key];
  }
  removeAll(keys: string[]): void {
    for (const key of keys) {
      delete this.cache[key];
    }
  }
  put(key: string, value: string, expirationInSeconds?: number): void {
    this.cache[key] = value;
    this.expirationInSeconds = expirationInSeconds;
  }
  get(key: string) {
    return this.cache[key];
  }
}
class FakePropertiesService {
  constructor(private readonly propertyStub = new PropertyStub()) {}

  getScriptProperties() {
    return this.propertyStub;
  }
}

class FakeCacheService {
  constructor(private readonly cacheStub = new CacheStub()) {}

  getScriptCache() {
    return this.cacheStub;
  }
}

class FakeBlob {
  constructor(readonly content: string) {}

  getDataAsString() {
    return this.content;
  }

  getBytes() {
    return `bytes:${this.content}`;
  }
}

class FakeGzip extends FakeBlob {
  constructor(gzipped: FakeBlob) {
    super('gzipped:' + gzipped.content);
  }
}

/**
 * Stubs utilities for testing gzip, blobs, etc.
 */
export class FakeUtilitiesService {
  newBlob(content: string): FakeBlob {
    return new FakeBlob(content.replace(/^bytes:/, ''));
  }

  gzip(content: FakeBlob): FakeGzip {
    return new FakeGzip(content);
  }

  ungzip(blob: FakeGzip | FakeBlob): FakeBlob {
    if (!blob.content.startsWith('gzipped')) {
      throw new Error('Not gzipped');
    }
    const content = blob.content.replace(/^gzipped:/, '');
    return new FakeBlob(content);
  }

  parseCsv(text: string) {
    // We don't need a special package because this test CSV is very
    // basic. No escaping, etc.
    const lines = text.split('\n').map((line: string) => line.split(','));
    return lines;
  }

  base64Encode(text: string) {
    if (!text.startsWith('bytes:')) {
      throw new Error('Not bytes');
    }
    return `encoded:${text.replace(/^bytes:/, '')}`;
  }

  base64Decode(text: string) {
    if (!text.startsWith('encoded')) {
      throw new Error('Not encoded');
    }

    return text.replace(/^encoded:/, 'bytes:');
  }

  sleep(msecs: number) {
    console.info(`skip sleep for ${msecs}`);
  }
}

/**
 * Stub for HTML output
 */
export class FakeHtmlOutput {}

/**
 * Stub for Drive testing
 */
export class FakeDrive {}

class FakeBigQuery {
  Jobs = new FakeBigQueryJobs();
}

class FakeBigQueryJobs {
  query: () => BigQuery.Schema.QueryResponse = () => {
    throw new Error('Not implemented');
  };
}

// findLastIndex isn't properly supported in TypeScript definitions at the moemnt.
declare global {
  interface Array<T> {
    /**
     * Returns the value of the last element in the array where predicate is true, and undefined
     * otherwise.
     * @param predicate findLast calls predicate once for each element of the array, in descending
     * order, until it finds one where predicate returns true. If such an element is found, findLast
     * immediately returns that element value. Otherwise, findLast returns undefined.
     * @param thisArg If provided, it will be used as the this value for each invocation of
     * predicate. If it is not provided, undefined is used instead.
     */
    findLast<S extends T>(
      predicate: (value: T, index: number, array: T[]) => value is S,
      // eslint-disable-next-line @typescript-eslint/no-explicit-any
      thisArg?: any,
    ): S | undefined;
    findLast(
      predicate: (value: T, index: number, array: T[]) => unknown,
      // eslint-disable-next-line @typescript-eslint/no-explicit-any
      thisArg?: any,
    ): T | undefined;

    /**
     * Returns the index of the last element in the array where predicate is true, and -1
     * otherwise.
     * @param predicate findLastIndex calls predicate once for each element of the array, in descending
     * order, until it finds one where predicate returns true. If such an element is found,
     * findLastIndex immediately returns that element index. Otherwise, findLastIndex returns -1.
     * @param thisArg If provided, it will be used as the this value for each invocation of
     * predicate. If it is not provided, undefined is used instead.
     */
    findLastIndex(
      predicate: (value: T, index: number, array: T[]) => unknown,
      // eslint-disable-next-line @typescript-eslint/no-explicit-any
      thisArg?: any,
    ): number;

    /**
     * Returns a copy of an array with its elements reversed.
     */
    toReversed(): T[];

    /**
     * Returns a copy of an array with its elements sorted.
     * @param compareFn Function used to determine the order of the elements. It is expected to return
     * a negative value if the first argument is less than the second argument, zero if they're equal, and a positive
     * value otherwise. If omitted, the elements are sorted in ascending, ASCII character order.
     * ```ts
     * [11, 2, 22, 1].toSorted((a, b) => a - b) // [1, 2, 11, 22]
     * ```
     */
    toSorted(compareFn?: (a: T, b: T) => number): T[];

    /**
     * Copies an array and removes elements and, if necessary, inserts new elements in their place. Returns the copied array.
     * @param start The zero-based location in the array from which to start removing elements.
     * @param deleteCount The number of elements to remove.
     * @param items Elements to insert into the copied array in place of the deleted elements.
     * @returns The copied array.
     */
    toSpliced(start: number, deleteCount: number, ...items: T[]): T[];

    /**
     * Copies an array and removes elements while returning the remaining elements.
     * @param start The zero-based location in the array from which to start removing elements.
     * @param deleteCount The number of elements to remove.
     * @returns A copy of the original array with the remaining elements.
     */
    toSpliced(start: number, deleteCount?: number): T[];

    /**
     * Copies an array, then overwrites the value at the provided index with the
     * given value. If the index is negative, then it replaces from the end
     * of the array.
     * @param index The index of the value to overwrite. If the index is
     * negative, then it replaces from the end of the array.
     * @param value The value to write into the copied array.
     * @returns The copied array with the updated value.
     */
    with(index: number, value: T): T[];
  }
}
