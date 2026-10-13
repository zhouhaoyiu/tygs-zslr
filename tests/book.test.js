import { test } from "node:test";
import assert from "node:assert/strict";
import { readFileSync } from "node:fs";
import * as XLSX from "xlsx";
import { DOMParser, XMLSerializer } from "@xmldom/xmldom";
import { unzipSync, zipSync, strToU8, strFromU8 } from "fflate";
import {
  assess,
  backupBook,
  confirmReading,
  createBook,
  deferReading,
  exceptionRows,
  handoffBook,
  exportWorkbook,
  inspectWorkbook,
  parseSheet,
  progress,
  reading,
  restoreBackup,
} from "../src/book.js";

globalThis.DOMParser = DOMParser;
globalThis.XMLSerializer = XMLSerializer;

test("structured deferrals survive backup and reject unknown categories", async () => {
  const { book, bytes } = await setup();
  deferReading(book, book.rows[0].id, "表盘无法辨认", "unreadable");
  assert.equal(exceptionRows(book).length, 1);
  const recovered = await restoreBackup(backupBook(book, bytes));
  assert.equal(recovered.book.entries[book.rows[0].id].reason, "unreadable");
  assert.throws(() => deferReading(book, book.rows[1].id, "x", "invalid"));
});

test("incomplete handoff separates blanks and drafts, restores directly and never mutates source", async () => {
  const { book, bytes } = await setup([
    first,
    second,
    { ...second, 抄表号: "003" },
  ]);
  commit(book, book.rows[0], 110);
  deferReading(book, book.rows[1].id, "无法入户", "inaccessible");
  book.drafts[book.rows[2].id] = {
    value: "211",
    mode: "normal",
    note: "=1+1",
    acknowledged: false,
  };
  const before = JSON.stringify(book);
  const original = new Uint8Array(bytes);
  const out = handoffBook(book, bytes);
  assert.deepEqual(out, handoffBook(book, bytes));
  assert.equal(JSON.stringify(book), before);
  assert.deepEqual(bytes, original);
  const files = unzipSync(out);
  assert.equal(files["已完成结果.xlsx"], undefined);
  assert.deepEqual(files["source.xlsx"], bytes);
  const wb = XLSX.read(files["交接清单-未完成-非计费表.xlsx"], {
    type: "array",
  });
  const rows = XLSX.utils.sheet_to_json(wb.Sheets.交接明细, { defval: "" });
  assert.equal(rows[0].已确认指数, 110);
  assert.equal(rows[1].已确认指数, "");
  assert.equal(rows[1].暂缓类型, "无法入户");
  assert.equal(rows[2].已确认指数, "");
  assert.equal(rows[2].未确认草稿, "211");
  assert.equal(wb.Sheets.交接明细.P4.t, "s");
  assert.equal(wb.Sheets.交接明细.P4.f, undefined);
  assert.equal(XLSX.utils.sheet_to_json(wb.Sheets.未完成清单).length, 2);
  const recovered = await restoreBackup(out);
  assert.deepEqual(recovered.book.drafts, book.drafts);
  assert.equal(recovered.book.entries[book.rows[1].id].status, "deferred");
});

test("complete handoff embeds the same formal workbook and includes anomaly review", async () => {
  const { book, bytes } = await setup();
  commit(book, book.rows[0], 130);
  commit(book, book.rows[1], 210);
  const files = unzipSync(handoffBook(book, bytes));
  assert.deepEqual(files["已完成结果.xlsx"], exportWorkbook(book, bytes));
  assert.equal(exceptionRows(book).length, 1);
});
const first = {
  抄表号: "001",
  计费号: "B01",
  用户名称: "测试甲",
  水表地址: "测试101",
  指数: 100,
  录入月份: "2026-09",
  "2026-08": 90,
  备注信息: "旧备注",
};
const second = {
  抄表号: "002",
  指数: 200,
  录入月份: "2026-09",
  "2026-08": 190,
};
function workbook(rows = [first, second], extra = false) {
  const wb = XLSX.utils.book_new();
  const sheet = XLSX.utils.json_to_sheet(rows);
  sheet.A2.z = "@";
  XLSX.utils.book_append_sheet(wb, sheet, "抄表册");
  if (extra) {
    const other = XLSX.utils.aoa_to_sheet([["统计"], [123]]);
    other.A3 = { t: "n", f: "A2*2", v: 246 };
    other["!ref"] = "A1:A3";
    XLSX.utils.book_append_sheet(wb, other, "其他表");
  }
  return new Uint8Array(XLSX.write(wb, { bookType: "xlsx", type: "array" }));
}
async function setup(rows, extra) {
  const bytes = workbook(rows, extra);
  const candidate = inspectWorkbook(bytes, "input.xlsx");
  return { bytes, book: await createBook(candidate, "抄表册", "2026-10") };
}
function commit(book, row, value, overrides = {}) {
  return confirmReading(book, row.id, {
    value: String(value),
    mode: "normal",
    note: "",
    acknowledged: true,
    ...overrides,
  });
}

test("blank, negative, nonfinite and malformed readings never become zero", () => {
  for (const value of [
    "",
    " ",
    null,
    undefined,
    "-1",
    NaN,
    Infinity,
    "1e3",
    "1,000",
    "1.1234",
    "123abc",
  ])
    assert.throws(() => reading(value));
  assert.equal(reading("0"), 0);
  assert.equal(reading("100.125"), 100.125);
});
test("duplicate meter codes are distinct physical worksheet rows", async () => {
  const { book } = await setup([first, { ...second, 抄表号: first.抄表号 }]);
  commit(book, book.rows[0], 111);
  commit(book, book.rows[1], 222);
  assert.equal(book.entries[book.rows[0].id].value, 111);
  assert.equal(book.entries[book.rows[1].id].value, 222);
  assert.equal(book.duplicates, 2);
});
test("missing identity or baseline rejects import without mutating prior state", async () => {
  const { book } = await setup();
  const before = JSON.stringify(book);
  assert.throws(
    () => inspectWorkbook(workbook([first, { 指数: 200 }]), "bad.xlsx"),
    /缺少抄表号/,
  );
  assert.throws(
    () => inspectWorkbook(workbook([first, { 抄表号: "B" }]), "bad.xlsx"),
    /非负读数/,
  );
  assert.equal(JSON.stringify(book), before);
});
test("optional fields are never inherited from the preceding user", async () => {
  const { book } = await setup();
  assert.equal(book.rows[1].name, "");
  assert.equal(book.rows[1].address, "");
  assert.equal(book.rows[1].fields.备注信息, "");
});
test("invalid file, xls and duplicate headers are rejected", () => {
  assert.throws(() => inspectWorkbook(strToU8("invalid"), "invalid.xlsx"));
  assert.throws(() => inspectWorkbook(workbook(), "old.xls"), /另存为/);
  const sheet = XLSX.utils.aoa_to_sheet([
    ["抄表号", "指数", "指数"],
    ["A", 1, 2],
  ]);
  assert.throws(() => parseSheet(sheet, "s"), /重复表头/);
});
test("history uses month names rather than column positions", async () => {
  const { book } = await setup([
    { ...first, "2026-07": 30, 额外备注: "不参与计算" },
  ]);
  assert.equal(book.rows[0].previousUsage, 10);
  assert.equal(book.rows[0].historyMonth, "2026-08");
});
test("zero or absent usage history is handled explicitly", async () => {
  const { book } = await setup([
    { ...first, "2026-08": 100 },
    { 抄表号: "B", 指数: 0 },
  ]);
  assert.match(
    assess(book.rows[0], { value: "110", mode: "normal" }).warning,
    /上期用量为 0/,
  );
  assert.equal(
    assess(book.rows[1], { value: "0", mode: "normal" }).warning,
    "",
  );
});
test("anomaly requires explicit acknowledgment before it changes progress", async () => {
  const { book } = await setup();
  assert.throws(
    () => commit(book, book.rows[0], 150, { acknowledged: false }),
    /勾选确认/,
  );
  assert.equal(progress(book).recorded, 0);
});
test("lower readings require replacement and a reason", async () => {
  const { book } = await setup();
  assert.throws(() => commit(book, book.rows[0], 0), /低于上期/);
  assert.throws(
    () => commit(book, book.rows[0], 0, { mode: "replaced" }),
    /填写备注/,
  );
  commit(book, book.rows[0], 0, { mode: "replaced", note: "新表起始 0" });
  assert.equal(book.entries[book.rows[0].id].value, 0);
  assert.equal(book.entries[book.rows[0].id].usage, null);
});
test("pending, deferred and unconfirmed corrections all prevent final export", async () => {
  const { book, bytes } = await setup();
  assert.throws(() => exportWorkbook(book, bytes), /未录/);
  commit(book, book.rows[0], 110);
  deferReading(book, book.rows[1].id, "无法入户");
  assert.equal(progress(book).deferred, 1);
  assert.throws(() => exportWorkbook(book, bytes), /暂缓/);
  commit(book, book.rows[1], 210);
  book.drafts[book.rows[0].id] = { value: "112", mode: "normal", note: "" };
  assert.throws(() => exportWorkbook(book, bytes), /未确认/);
});
test("export is repeatable and never mutates the book or original bytes", async () => {
  const { book, bytes } = await setup(undefined, true);
  book.rows.forEach((row) => commit(book, row, row.baseline + 10));
  const beforeBook = JSON.stringify(book);
  const beforeBytes = bytes.slice();
  const a = exportWorkbook(book, bytes);
  const b = exportWorkbook(book, bytes);
  assert.deepEqual(a, b);
  assert.deepEqual(bytes, beforeBytes);
  assert.equal(JSON.stringify(book), beforeBook);
  const output = XLSX.read(a, { type: "array" });
  const rows = XLSX.utils.sheet_to_json(output.Sheets.抄表册);
  assert.equal(rows[0].指数, 110);
  assert.equal(rows[0]["2026-09"], 100);
  assert.equal(rows[0].录入月份, "2026-10");
  assert.equal(rows[0]["2026-10"], undefined);
  assert.equal(rows[1].指数, 210);
  const sourceFiles = unzipSync(bytes),
    outputFiles = unzipSync(a);
  for (const name of Object.keys(sourceFiles)) {
    if (name !== "xl/worksheets/sheet1.xml" && name !== "xl/workbook.xml")
      assert.deepEqual(outputFiles[name], sourceFiles[name], name);
  }
  assert.equal(output.Sheets.其他表.A3.f, "A2*2");
});
test("zero replacement reading exports as zero and retains original remarks", async () => {
  const { book, bytes } = await setup([first]);
  commit(book, book.rows[0], 0, { mode: "replaced", note: "新表编号 123" });
  const rows = XLSX.utils.sheet_to_json(
    XLSX.read(exportWorkbook(book, bytes)).Sheets.抄表册,
  );
  assert.equal(rows[0].指数, 0);
  assert.match(rows[0].备注信息, /旧备注.*换表.*新表编号/);
});
test("corrections retain source baseline and an edit history", async () => {
  const { book } = await setup([first]);
  commit(book, book.rows[0], 110);
  commit(book, book.rows[0], 112);
  assert.equal(book.rows[0].baseline, 100);
  assert.equal(book.history.length, 2);
  assert.equal(book.history[1].before.value, 110);
  assert.equal(book.history[1].after.value, 112);
});
test("incomplete backup round-trips original workbook, drafts and progress", async () => {
  const { book, bytes } = await setup();
  commit(book, book.rows[0], 110);
  book.activeRowId = book.rows[1].id;
  book.drafts[book.rows[1].id] = {
    value: "2",
    mode: "normal",
    note: "输入中",
    acknowledged: false,
  };
  const restored = await restoreBackup(backupBook(book, bytes));
  assert.deepEqual(restored.bytes, bytes);
  assert.equal(restored.book.entries[book.rows[0].id].value, 110);
  assert.deepEqual(restored.book.drafts, book.drafts);
  assert.equal(restored.book.activeRowId, book.activeRowId);
  assert.notEqual(restored.book.id, book.id);
  assert.deepEqual(restored.book.history, book.history);
});
test("tampered backups and unsupported versions are rejected", async () => {
  const { book, bytes } = await setup();
  const invalid = { ...book, schema: 99 };
  await assert.rejects(
    restoreBackup(
      zipSync({
        "source.xlsx": bytes,
        "progress.json": strToU8(JSON.stringify(invalid)),
      }),
    ),
    /版本/,
  );
  invalid.schema = 1;
  invalid.sourceHash = "tampered";
  await assert.rejects(
    restoreBackup(
      zipSync({
        "source.xlsx": bytes,
        "progress.json": strToU8(JSON.stringify(invalid)),
      }),
    ),
    /校验失败/,
  );
});
test("invalid or non-forward period is rejected", async () => {
  const candidate = inspectWorkbook(workbook(), "input.xlsx");
  await assert.rejects(createBook(candidate, "抄表册", "2026-09"), /晚于/);
  await assert.rejects(createBook(candidate, "抄表册", "2026-13"), /有效/);
});
test("repository samples import with integer readings and text meter IDs", () => {
  for (const filename of [
    "2026-08 A小区1单元.xlsx",
    "2026-10 A小区1单元示例.xlsx",
    "A小区1单元.xlsx",
    "TEST.xlsx",
    "水表数据.xlsx",
  ]) {
    const result = inspectWorkbook(
      readFileSync(new URL(`../${filename}`, import.meta.url)),
      filename,
    );
    assert.equal(result.sheets.length, filename === "TEST.xlsx" ? 9 : 1);
    for (const sheet of result.sheets) {
      assert.equal(sheet.error, "", `${filename}: ${sheet.name}`);
      assert.equal(sheet.count, sheet.name === "Sheet1" ? 30 : 6);
      const parsed = parseSheet(result.workbook.Sheets[sheet.name], sheet.name);
      assert.equal(parsed.rows[0].meter, "000001");
      assert.equal(parsed.rows[0].baseline, 0);
      assert.ok(parsed.rows.every((row) => Number.isInteger(row.baseline)));
      assert.ok(parsed.rows.every((row) => row.name.startsWith("示例用户")));
    }
  }
});

test("notes starting with a formula remain literal text in exported Excel", async () => {
  const { book, bytes } = await setup([{ ...first, 备注信息: "" }]);
  const note = '=HYPERLINK("https://example.invalid") < & >';
  commit(book, book.rows[0], 110, { note });
  const output = XLSX.read(exportWorkbook(book, bytes));
  const cell =
    output.Sheets.抄表册[
      XLSX.utils.encode_cell({ r: 1, c: book.columns.备注信息 })
    ];
  assert.equal(cell.v, note);
  assert.equal(cell.f, undefined);
  assert.equal(cell.t, "s");
});

test("backup can recover confirmed values pending recheck after threshold changes", async () => {
  const { book, bytes } = await setup([first]);
  commit(book, book.rows[0], 112, { acknowledged: false });
  book.ratio = 10;
  book.drafts[book.rows[0].id] = {
    value: "113",
    mode: "normal",
    note: "",
    acknowledged: false,
  };
  const restored = await restoreBackup(backupBook(book, bytes));
  assert.equal(
    restored.book.entries[book.rows[0].id].savedAt,
    book.entries[book.rows[0].id].savedAt,
  );
  assert.equal(restored.book.drafts[book.rows[0].id].value, "113");
  assert.equal(progress(restored.book).complete, false);
});

test("replacing cached formula readings removes stale calculation chain metadata", async () => {
  const files = unzipSync(workbook([first]));
  files["xl/calcChain.xml"] = strToU8(
    '<calcChain xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><c r="E2" i="1"/></calcChain>',
  );
  files["xl/_rels/workbook.xml.rels"] = strToU8(
    strFromU8(files["xl/_rels/workbook.xml.rels"]).replace(
      "</Relationships>",
      '<Relationship Id="rIdCalc" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/calcChain" Target="calcChain.xml"/></Relationships>',
    ),
  );
  files["[Content_Types].xml"] = strToU8(
    strFromU8(files["[Content_Types].xml"]).replace(
      "</Types>",
      '<Override PartName="/xl/calcChain.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.calcChain+xml"/></Types>',
    ),
  );
  const bytes = zipSync(files);
  const book = await createBook(
    inspectWorkbook(bytes, "input.xlsx"),
    "抄表册",
    "2026-10",
  );
  commit(book, book.rows[0], 110);
  const output = unzipSync(exportWorkbook(book, bytes));
  assert.equal(output["xl/calcChain.xml"], undefined);
  assert.doesNotMatch(
    strFromU8(output["xl/_rels/workbook.xml.rels"]),
    /calcChain/,
  );
  assert.doesNotMatch(strFromU8(output["[Content_Types].xml"]), /calcChain/);
  assert.match(strFromU8(output["xl/workbook.xml"]), /fullCalcOnLoad="1"/);
});
