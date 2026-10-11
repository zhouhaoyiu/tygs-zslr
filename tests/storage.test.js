import { test } from "node:test";
import assert from "node:assert/strict";
import { indexedDB, IDBObjectStore } from "fake-indexeddb";
import { BookStorage, openDatabase } from "../src/storage.js";
globalThis.indexedDB = indexedDB;
const make = () => ({
  id: crypto.randomUUID(),
  updatedAt: new Date().toISOString(),
  entries: {},
  drafts: {},
});
test("atomic initial save retains source and draft across reopening", async () => {
  const store = new BookStorage(await openDatabase());
  const book = make();
  const bytes = new Uint8Array([1, 2, 3]);
  const revision = await store.save(book, 0, bytes);
  book.drafts.row = { value: "12" };
  await store.save(book, revision);
  store.db.close();
  const reopened = new BookStorage(await openDatabase());
  assert.deepEqual(await reopened.source(book.id), bytes);
  assert.equal(
    (await reopened.list()).find((item) => item.id === book.id).book.drafts.row
      .value,
    "12",
  );
  reopened.db.close();
});
test("stale tab cannot silently overwrite a newer revision", async () => {
  const a = new BookStorage(await openDatabase()),
    b = new BookStorage(await openDatabase());
  const book = make();
  await a.save(book, 0, new Uint8Array([1]));
  await a.save({ ...book, entries: { first: 110 } }, 1);
  await assert.rejects(
    b.save({ ...book, entries: { second: 220 } }, 1),
    /另一页面/,
  );
  assert.deepEqual(
    (await a.list()).find((item) => item.id === book.id).book.entries,
    { first: 110 },
  );
  a.db.close();
  b.db.close();
});
test("a state-only save cannot create a book without its original source", async () => {
  const store = new BookStorage(await openDatabase());
  const book = make();
  await assert.rejects(store.save(book, 0), /原始工作簿/);
  assert.equal(
    (await store.list()).some((item) => item.id === book.id),
    false,
  );
  store.db.close();
});

test("a failed write aborts atomically and preserves the previous revision", async () => {
  const store = new BookStorage(await openDatabase());
  const book = make();
  await store.save(book, 0, new Uint8Array([1]));
  const originalPut = IDBObjectStore.prototype.put;
  IDBObjectStore.prototype.put = function () {
    throw new DOMException("存储空间不足", "QuotaExceededError");
  };
  try {
    await assert.rejects(
      store.save({ ...book, drafts: { row: "123" } }, 1),
      /空间不足/,
    );
  } finally {
    IDBObjectStore.prototype.put = originalPut;
  }
  const saved = (await store.list()).find((item) => item.id === book.id);
  assert.equal(saved.revision, 1);
  assert.deepEqual(saved.book.drafts, {});
  store.db.close();
});
