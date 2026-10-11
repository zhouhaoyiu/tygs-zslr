import * as XLSX from "xlsx";
import { unzipSync, zipSync, strFromU8, strToU8 } from "fflate";

export const SCHEMA = 1;
export const MAX_FILE_SIZE = 20 * 1024 * 1024;
const NS = "http://schemas.openxmlformats.org/spreadsheetml/2006/main";
const MONTH = /^\d{4}-(0[1-9]|1[0-2])$/;
const modes = ["normal", "replaced", "removed"];
export const DEFER_REASONS = {
  inaccessible: "无法入户",
  unreadable: "表盘不清",
  broken: "水表故障",
  missing: "未找到水表",
  other: "其他原因",
};
const copy = (value) => JSON.parse(JSON.stringify(value));
const fail = (message) => {
  throw new Error(message);
};
export const localMonth = (date = new Date()) =>
  `${date.getFullYear()}-${String(date.getMonth() + 1).padStart(2, "0")}`;

export function reading(value) {
  const text = String(value ?? "").trim();
  if (!/^\d+(\.\d{1,3})?$/.test(text))
    fail("请输入非负读数，最多保留 3 位小数。");
  const number = Number(text);
  if (!Number.isFinite(number) || number > 1e12) fail("读数超出有效范围。");
  return number;
}

function xml(bytes) {
  const text = strFromU8(bytes);
  if (/<!DOCTYPE|<!ENTITY/i.test(text)) fail("不支持含外部实体的工作簿。");
  const doc = new DOMParser().parseFromString(text, "application/xml");
  if (doc.getElementsByTagName("parsererror").length)
    fail("工作簿 XML 已损坏。");
  return doc;
}
const elements = (node, name) =>
  Array.from(node.getElementsByTagNameNS("*", name));
const encodeXml = (doc) => strToU8(new XMLSerializer().serializeToString(doc));

function unpack(bytes) {
  let total = 0;
  return unzipSync(bytes, {
    filter(file) {
      total += file.originalSize;
      if (total > 100 * 1024 * 1024) fail("解压后的文件过大，请拆分工作簿。");
      return true;
    },
  });
}

export function inspectWorkbook(input, filename) {
  if (!/\.xlsx$/i.test(filename))
    fail("请选择 .xlsx 工作簿。旧版 .xls 请在 Excel 中另存为 .xlsx。");
  const bytes = new Uint8Array(input);
  if (!bytes.length || bytes.length > MAX_FILE_SIZE)
    fail("文件为空或超过 20 MB。");
  let files, workbook;
  try {
    files = unpack(bytes);
    if (!files["xl/workbook.xml"] || !files["xl/_rels/workbook.xml.rels"])
      fail("无效的 .xlsx 文件。");
    workbook = XLSX.read(bytes, {
      type: "array",
      cellStyles: true,
      cellFormula: true,
    });
  } catch (error) {
    fail(`无法读取工作簿：${error.message}`);
  }
  const sheets = workbook.SheetNames.map((name) => {
    try {
      const parsed = parseSheet(workbook.Sheets[name], name);
      return {
        name,
        count: parsed.rows.length,
        error: "",
        duplicates: parsed.duplicates,
      };
    } catch (error) {
      return { name, count: 0, error: error.message, duplicates: 0 };
    }
  });
  if (!sheets.some((sheet) => !sheet.error))
    fail(
      `没有可录入的工作表。${sheets.map((sheet) => `${sheet.name}：${sheet.error}`).join(" ")}`,
    );
  return { filename, bytes, workbook, sheets };
}

function valueAt(sheet, row, col) {
  return sheet[XLSX.utils.encode_cell({ r: row, c: col })];
}
function textAt(sheet, row, col) {
  const cell = valueAt(sheet, row, col);
  return cell ? String(cell.w ?? cell.v ?? "").trim() : "";
}

export function parseSheet(sheet, sheetName) {
  if (!sheet["!ref"]) fail("空工作表。");
  const range = XLSX.utils.decode_range(sheet["!ref"]);
  if (range.e.r > 50000 || range.e.c > 1000)
    fail("表格范围过大，请清理空白格式后重试。");
  let headerRow = -1,
    headers = [];
  for (let r = range.s.r; r <= Math.min(range.e.r, range.s.r + 19); r++) {
    const values = Array.from({ length: range.e.c + 1 }, (_, c) =>
      textAt(sheet, r, c),
    );
    if (values.includes("抄表号") && values.includes("指数")) {
      headerRow = r;
      headers = values;
      break;
    }
  }
  if (headerRow < 0) fail("缺少「抄表号」或「指数」表头。");
  const named = headers.filter(Boolean);
  if (new Set(named).size !== named.length)
    fail("存在重复表头，请先为每一列设置唯一名称。");
  const columns = Object.fromEntries(
    headers.map((name, index) => [name, index]).filter(([name]) => name),
  );
  const historical = headers
    .filter((header) => MONTH.test(header))
    .sort()
    .reverse();
  const rows = [];
  for (let r = headerRow + 1; r <= range.e.r; r++) {
    if (!headers.some((name, c) => name && textAt(sheet, r, c) !== ""))
      continue;
    const label = `第 ${r + 1} 行`;
    const meter = textAt(sheet, r, columns.抄表号);
    if (!meter) fail(`${label}缺少抄表号，不能确定用户。`);
    let baseline;
    try {
      const cell = valueAt(sheet, r, columns.指数);
      baseline = reading(cell?.v);
    } catch (error) {
      fail(`${label}：${error.message}`);
    }
    const month =
      columns.录入月份 === undefined ? "" : textAt(sheet, r, columns.录入月份);
    if (month && !MONTH.test(month)) fail(`${label}录入月份应为 YYYY-MM。`);
    const historyMonth = historical.find(
      (key) => (!month || key < month) && textAt(sheet, r, columns[key]) !== "",
    );
    let previous = null;
    if (historyMonth) {
      try {
        previous = reading(valueAt(sheet, r, columns[historyMonth])?.v);
      } catch {
        /* Invalid history never becomes a zero baseline. */
      }
    }
    const fields = Object.fromEntries(
      headers
        .map((name, c) => [name, name ? textAt(sheet, r, c) : ""])
        .filter(([name]) => name),
    );
    rows.push({
      id: `${sheetName}:${r + 1}`,
      excelRow: r + 1,
      meter,
      baseline,
      month,
      baselineFormula: Boolean(valueAt(sheet, r, columns.指数)?.f),
      name: fields.用户名称 || "",
      account: fields.计费号 || "",
      address: fields.水表地址 || "",
      previousUsage:
        previous !== null && baseline >= previous
          ? Math.round((baseline - previous) * 1000) / 1000
          : null,
      historyMonth: historyMonth || "",
      fields,
    });
  }
  if (!rows.length) fail("没有可录入的数据行。");
  const counts = new Map();
  rows.forEach((row) =>
    counts.set(row.meter, (counts.get(row.meter) || 0) + 1),
  );
  rows.forEach((row) => {
    row.duplicate = counts.get(row.meter) > 1;
  });
  return {
    headers,
    columns,
    headerRow,
    rows,
    duplicates: rows.filter((row) => row.duplicate).length,
  };
}

export async function createBook(candidate, sheetName, period) {
  if (!MONTH.test(period)) fail("请选择有效的抄表月份。");
  if (!candidate.workbook.SheetNames.includes(sheetName))
    fail("工作表不存在。");
  const parsed = parseSheet(candidate.workbook.Sheets[sheetName], sheetName);
  if (parsed.rows.some((row) => row.month && row.month >= period))
    fail("本期月份应晚于表内的录入月份。修改已有记录请打开已保存的抄表本。");
  const sourceHash = Array.from(
    new Uint8Array(await crypto.subtle.digest("SHA-256", candidate.bytes)),
    (b) => b.toString(16).padStart(2, "0"),
  ).join("");
  return {
    schema: SCHEMA,
    id: crypto.randomUUID(),
    filename: candidate.filename,
    sheetName,
    period,
    sourceHash,
    ...parsed,
    entries: {},
    drafts: {},
    history: [],
    activeRowId: parsed.rows[0].id,
    ratio: 30,
    createdAt: new Date().toISOString(),
    updatedAt: new Date().toISOString(),
  };
}

export function assess(row, draft, ratio = 30) {
  if (!modes.includes(draft.mode)) fail("请选择有效的抄表情况。");
  const value = reading(draft.value);
  const note = String(draft.note || "").trim();
  if (note.length > 500) fail("备注不能超过 500 字。");
  if (draft.mode !== "normal" && !note) fail("换表或拆表需要填写备注。");
  if (value < row.baseline && draft.mode !== "replaced")
    fail("本期读数低于上期。发生换表时，请选择「换表」并填写备注。");
  const usage =
    draft.mode === "replaced"
      ? null
      : Math.round((value - row.baseline) * 1000) / 1000;
  let warning = "";
  if (usage !== null && row.previousUsage !== null) {
    if (row.previousUsage === 0 && usage > 0)
      warning = "上期用量为 0，本期出现用量，请核对。";
    else if (row.previousUsage > 0) {
      const change = ((usage - row.previousUsage) / row.previousUsage) * 100;
      if (Math.abs(change) > ratio)
        warning = `用量较上期${change > 0 ? "增加" : "减少"} ${Math.abs(change).toFixed(1)}%，请核对。`;
    }
  }
  return { value, mode: draft.mode, note, usage, warning };
}

export function confirmReading(book, rowId, draft) {
  const row = book.rows.find((item) => item.id === rowId);
  if (!row) fail("用户行不存在，请重新打开抄表本。");
  const result = assess(row, draft, book.ratio);
  if (result.warning && !draft.acknowledged) fail("请核对异常用量并勾选确认。");
  const previous = book.entries[rowId] || null;
  const entry = {
    ...result,
    status: "recorded",
    acknowledged: Boolean(draft.acknowledged),
    savedAt: new Date().toISOString(),
  };
  if (
    !previous ||
    ["value", "mode", "note"].some((key) => previous[key] !== entry[key])
  ) {
    book.history.push({
      rowId,
      at: entry.savedAt,
      before: previous,
      after: entry,
    });
  }
  book.entries[rowId] = entry;
  delete book.drafts[rowId];
  book.updatedAt = entry.savedAt;
  return entry;
}

export function deferReading(book, rowId, note, reason = "other") {
  if (!Object.hasOwn(DEFER_REASONS, reason)) fail("暂缓类型无效。");
  if (!String(note || "").trim()) fail("请填写暂缓原因。");
  if (String(note).length > 500) fail("备注不能超过 500 字。");
  if (!book.rows.some((row) => row.id === rowId)) fail("用户行不存在。");
  const previous = book.entries[rowId] || null;
  const entry = {
    status: "deferred",
    reason,
    note: String(note).trim(),
    savedAt: new Date().toISOString(),
  };
  book.entries[rowId] = entry;
  book.history.push({
    rowId,
    at: entry.savedAt,
    before: previous,
    after: entry,
  });
  book.updatedAt = entry.savedAt;
}

export function progress(book) {
  const recorded = book.rows.filter(
    (row) => book.entries[row.id]?.status === "recorded",
  ).length;
  const deferred = book.rows.filter(
    (row) => book.entries[row.id]?.status === "deferred",
  ).length;
  const drafts = Object.keys(book.drafts).length;
  return {
    total: book.rows.length,
    recorded,
    deferred,
    pending: book.rows.length - recorded - deferred,
    drafts,
    complete: recorded === book.rows.length && drafts === 0,
  };
}

function worksheetPath(files, sheetName) {
  const book = xml(files["xl/workbook.xml"]);
  const sheet = elements(book, "sheet").find(
    (node) => node.getAttribute("name") === sheetName,
  );
  const rel = elements(
    xml(files["xl/_rels/workbook.xml.rels"]),
    "Relationship",
  ).find(
    (node) =>
      node.getAttribute("Id") ===
      (sheet?.getAttribute("r:id") ||
        sheet?.getAttributeNS(
          "http://schemas.openxmlformats.org/officeDocument/2006/relationships",
          "id",
        )),
  );
  if (!rel) fail("无法定位工作表。");
  const target = rel.getAttribute("Target");
  const path = target.startsWith("/") ? target.slice(1) : `xl/${target}`;
  if (path.includes("..") || !files[path]) fail("工作表路径无效。");
  return path;
}

function setCell(doc, rowNumber, column, value) {
  const sheetData = elements(doc, "sheetData")[0];
  const rows = Array.from(sheetData.childNodes).filter(
    (node) => node.localName === "row",
  );
  let row = rows.find((node) => Number(node.getAttribute("r")) === rowNumber);
  if (!row) {
    row = doc.createElementNS(NS, "row");
    row.setAttribute("r", String(rowNumber));
    sheetData.insertBefore(
      row,
      rows.find((node) => Number(node.getAttribute("r")) > rowNumber) || null,
    );
  }
  const ref = XLSX.utils.encode_cell({ r: rowNumber - 1, c: column });
  const cells = Array.from(row.childNodes).filter(
    (node) => node.localName === "c",
  );
  let cell = cells.find((node) => node.getAttribute("r") === ref);
  if (!cell) {
    cell = doc.createElementNS(NS, "c");
    cell.setAttribute("r", ref);
    row.insertBefore(
      cell,
      cells.find(
        (node) => XLSX.utils.decode_cell(node.getAttribute("r")).c > column,
      ) || null,
    );
  }
  while (cell.firstChild) cell.removeChild(cell.firstChild);
  if (typeof value === "number") {
    cell.setAttribute("t", "n");
    const v = doc.createElementNS(NS, "v");
    v.textContent = String(value);
    cell.appendChild(v);
  } else {
    cell.setAttribute("t", "inlineStr");
    const inline = doc.createElementNS(NS, "is");
    const t = doc.createElementNS(NS, "t");
    t.setAttribute("xml:space", "preserve");
    t.textContent = String(value);
    inline.appendChild(t);
    cell.appendChild(inline);
  }
}

export function exportWorkbook(book, original) {
  if (!progress(book).complete)
    fail(
      "仍有未录、暂缓或未确认修改，请处理完后导出。进度可通过「备份」保存。",
    );
  const files = unpack(new Uint8Array(original));
  const path = worksheetPath(files, book.sheetName);
  const doc = xml(files[path]);
  const columns = { ...book.columns };
  let next = book.headers.length;
  const ensure = (name) => {
    if (columns[name] === undefined) {
      columns[name] = next++;
      setCell(doc, book.headerRow + 1, columns[name], name);
    }
    return columns[name];
  };
  ensure("录入月份");
  ensure("备注信息");
  for (const row of book.rows) {
    const entry = book.entries[row.id];
    const result = assess(row, entry, book.ratio);
    if (result.warning && !entry.acknowledged)
      fail(`第 ${row.excelRow} 行异常用量尚未确认。`);
    if (row.month) setCell(doc, row.excelRow, ensure(row.month), row.baseline);
    setCell(doc, row.excelRow, columns.指数, result.value);
    setCell(doc, row.excelRow, columns.录入月份, book.period);
    const note = [
      row.fields.备注信息,
      result.mode === "normal"
        ? ""
        : result.mode === "replaced"
          ? `${book.period} 换表`
          : `${book.period} 拆表`,
      result.note,
    ]
      .filter(Boolean)
      .join("；");
    if (note) setCell(doc, row.excelRow, columns.备注信息, note);
  }
  const dimension = elements(doc, "dimension")[0];
  if (dimension) {
    const range = XLSX.utils.decode_range(dimension.getAttribute("ref"));
    range.e.c = Math.max(range.e.c, next - 1);
    dimension.setAttribute("ref", XLSX.utils.encode_range(range));
  }
  files[path] = encodeXml(doc);
  // Keep worksheet content intact outside the selected sheet and invalidate calculation metadata.
  const wb = xml(files["xl/workbook.xml"]);
  let calc = elements(wb, "calcPr")[0];
  if (!calc) {
    calc = wb.createElementNS(NS, "calcPr");
    const following = new Set([
      "oleSize",
      "customWorkbookViews",
      "pivotCaches",
      "smartTagPr",
      "smartTagTypes",
      "webPublishing",
      "fileRecoveryPr",
      "webPublishObjects",
      "extLst",
    ]);
    const before = Array.from(wb.documentElement.childNodes).find((node) =>
      following.has(node.localName),
    );
    wb.documentElement.insertBefore(calc, before || null);
  }
  calc.setAttribute("fullCalcOnLoad", "1");
  calc.setAttribute("forceFullCalc", "1");
  calc.setAttribute("calcMode", "auto");
  files["xl/workbook.xml"] = encodeXml(wb);
  if (files["xl/calcChain.xml"]) {
    delete files["xl/calcChain.xml"];
    const relationships = xml(files["xl/_rels/workbook.xml.rels"]);
    for (const rel of elements(relationships, "Relationship")) {
      if (rel.getAttribute("Type").endsWith("/calcChain"))
        rel.parentNode.removeChild(rel);
    }
    files["xl/_rels/workbook.xml.rels"] = encodeXml(relationships);
    const types = xml(files["[Content_Types].xml"]);
    for (const type of elements(types, "Override")) {
      if (type.getAttribute("PartName") === "/xl/calcChain.xml")
        type.parentNode.removeChild(type);
    }
    files["[Content_Types].xml"] = encodeXml(types);
  }
  return zipSync(
    Object.fromEntries(
      Object.entries(files).map(([name, bytes]) => [
        name,
        [bytes, { mtime: new Date(2000, 0, 1) }],
      ]),
    ),
    { level: 6 },
  );
}

export function backupBook(book, original) {
  return zipSync({
    "source.xlsx": new Uint8Array(original),
    "progress.json": strToU8(JSON.stringify(copy(book))),
  });
}

export function exceptionRows(book) {
  return book.rows.filter((row) => {
    const entry = book.entries[row.id];
    return (
      entry?.status === "deferred" ||
      (entry?.status === "recorded" &&
        (entry.mode !== "normal" ||
          Boolean(assess(row, entry, book.ratio).warning)))
    );
  });
}

export function handoffBook(book, original) {
  const stats = progress(book);
  const exceptions = new Set(exceptionRows(book).map((row) => row.id));
  const headers = [
    "工作表",
    "原表行号",
    "抄表号",
    "计费号",
    "用户名称",
    "水表地址",
    "本期月份",
    "状态",
    "上期指数",
    "已确认指数",
    "已确认用量",
    "未确认草稿",
    "情况",
    "暂缓类型",
    "已确认备注",
    "草稿备注",
    "确认时间",
  ];
  const data = book.rows.map((row) => {
    const entry = book.entries[row.id];
    const draft = book.drafts[row.id];
    const recorded = entry?.status === "recorded";
    return [
      book.sheetName,
      row.excelRow,
      row.meter,
      row.account,
      row.name,
      row.address,
      book.period,
      draft ? "未确认修改" : recorded ? "已录" : entry ? "暂缓" : "未录",
      row.baseline,
      recorded ? entry.value : "",
      recorded ? (assess(row, entry, book.ratio).usage ?? "") : "",
      draft?.value ?? "",
      recorded
        ? { normal: "正常", replaced: "换表待核算", removed: "拆表" }[
            entry.mode
          ]
        : "",
      entry?.status === "deferred"
        ? DEFER_REASONS[entry.reason || "other"]
        : "",
      entry?.note || "",
      draft?.note || "",
      entry?.savedAt || "",
    ];
  });
  const wb = XLSX.utils.book_new();
  for (const [name, rows] of [
    ["交接明细", data],
    [
      "未完成清单",
      data.filter(
        (_, i) =>
          book.entries[book.rows[i].id]?.status !== "recorded" ||
          book.drafts[book.rows[i].id],
      ),
    ],
    ["异常复核清单", data.filter((_, i) => exceptions.has(book.rows[i].id))],
  ]) {
    const sheet = XLSX.utils.aoa_to_sheet([headers, ...rows]);
    sheet["!cols"] = headers.map(() => ({ wch: 20 }));
    XLSX.utils.book_append_sheet(wb, sheet, name);
  }
  const state = stats.complete ? "已完成" : "未完成";
  const files = {
    "source.xlsx": new Uint8Array(original),
    "progress.json": strToU8(JSON.stringify(copy(book))),
    [`交接清单-${state}-非计费表.xlsx`]: new Uint8Array(
      XLSX.write(wb, { type: "array", bookType: "xlsx" }),
    ),
    "交接说明.txt": strToU8(
      `状态：${state}\n月份：${book.period}\n工作表：${book.sheetName}\n总户数：${stats.total}\n已录：${stats.recorded}\n未录：${stats.pending}\n暂缓：${stats.deferred}\n未确认修改：${stats.drafts}\n异常复核：${exceptions.size}\n原表 SHA-256：${book.sourceHash}\n\n交接清单不是计费表。未录与暂缓的本期指数留空，草稿单独列出。换表用量待业务核算。\n本 ZIP 可直接在电子抄表本中恢复。source.xlsx 是未改动的原始表格。\n${stats.complete ? "本包另含已完成结果.xlsx。" : "本包不包含正式结果表，继续录入后再正式导出。"}\n`,
    ),
  };
  if (stats.complete) files["已完成结果.xlsx"] = exportWorkbook(book, original);
  return zipSync(
    Object.fromEntries(
      Object.entries(files).map(([name, bytes]) => [
        name,
        [bytes, { mtime: new Date(2000, 0, 1) }],
      ]),
    ),
    { level: 6 },
  );
}

export async function restoreBackup(bytes) {
  if (bytes.byteLength > MAX_FILE_SIZE * 2) fail("备份文件过大。");
  const files = unpack(new Uint8Array(bytes));
  if (!files["source.xlsx"] || !files["progress.json"])
    fail("不是有效的抄表本备份。");
  const saved = JSON.parse(strFromU8(files["progress.json"]));
  if (saved.schema !== SCHEMA) fail("备份版本不兼容。");
  const candidate = inspectWorkbook(files["source.xlsx"], saved.filename);
  const book = await createBook(candidate, saved.sheetName, saved.period);
  if (book.sourceHash !== saved.sourceHash) fail("原始工作簿校验失败。");
  if (!Number.isFinite(saved.ratio) || saved.ratio < 1 || saved.ratio > 1000)
    fail("异常阈值无效。");
  book.ratio = saved.ratio;
  const ids = new Set(book.rows.map((row) => row.id));
  for (const [id, entry] of Object.entries(saved.entries || {})) {
    if (!ids.has(id)) fail("备份包含未知用户行。");
    if (!entry || typeof entry.savedAt !== "string") fail("备份记录时间无效。");
    if (entry.status === "recorded") {
      const result = assess(
        book.rows.find((row) => row.id === id),
        entry,
        book.ratio,
      );
      if (result.warning && !entry.acknowledged && !saved.drafts?.[id])
        fail("备份中存在未经确认的异常读数。");
      book.entries[id] = {
        ...result,
        status: "recorded",
        acknowledged: Boolean(entry.acknowledged),
        savedAt: entry.savedAt,
      };
    } else if (entry.status === "deferred") {
      deferReading(book, id, entry.note, entry.reason || "other");
      book.entries[id].savedAt = entry.savedAt;
    } else fail("备份包含无效录入状态。");
  }
  for (const [id, draft] of Object.entries(saved.drafts || {})) {
    if (
      !ids.has(id) ||
      typeof draft !== "object" ||
      !draft ||
      !modes.includes(draft.mode) ||
      String(draft.value ?? "").length > 40 ||
      String(draft.note ?? "").length > 500
    )
      fail("草稿数据无效。");
    book.drafts[id] = {
      value: String(draft.value ?? ""),
      mode: draft.mode,
      note: String(draft.note || ""),
      acknowledged: Boolean(draft.acknowledged),
      ...(draft.deferReason
        ? {
            deferReason: Object.hasOwn(DEFER_REASONS, draft.deferReason)
              ? draft.deferReason
              : "other",
          }
        : {}),
    };
  }
  if (ids.has(saved.activeRowId)) book.activeRowId = saved.activeRowId;
  if (
    !Array.isArray(saved.history) ||
    saved.history.some(
      (item) => !ids.has(item.rowId) || typeof item.at !== "string",
    )
  )
    fail("修改记录无效。");
  book.history = copy(saved.history);
  for (const key of ["lastBackupAt", "lastHandoffAt"]) {
    if (
      typeof saved[key] === "string" &&
      Number.isFinite(Date.parse(saved[key]))
    )
      book[key] = saved[key];
  }
  book.restoredAt = new Date().toISOString();
  return { book, bytes: files["source.xlsx"] };
}

export function makeExample() {
  const wb = XLSX.utils.book_new();
  const month = localMonth();
  const date = new Date(`${month}-15T12:00:00`);
  date.setMonth(date.getMonth() - 1);
  const last = localMonth(date);
  date.setMonth(date.getMonth() - 1);
  const previous = localMonth(date);
  const rows = Array.from({ length: 12 }, (_, i) => ({
    抄表号: `DEMO-${String(i + 1).padStart(3, "0")}`,
    计费号: `TEST-${i + 1}`,
    用户名称: `示例用户 ${i + 1}`,
    水表地址: `示例小区 1 号楼 ${Math.floor(i / 2) + 1}0${(i % 2) + 1}`,
    指数: 120 + i * 37,
    录入月份: last,
    [previous]: 108 + i * 37,
    备注信息: "",
  }));
  XLSX.utils.book_append_sheet(
    wb,
    XLSX.utils.json_to_sheet(rows),
    "示例抄表册",
  );
  return {
    bytes: new Uint8Array(XLSX.write(wb, { type: "array", bookType: "xlsx" })),
    filename: "示例抄表本.xlsx",
  };
}
