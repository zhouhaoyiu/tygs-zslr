<script setup>
import { computed, nextTick, onMounted, ref } from "vue";
import {
  AlertCircle,
  ArrowDownToLine,
  ArrowLeft,
  ArrowRight,
  BookOpen,
  Check,
  CheckCheck,
  ChevronRight,
  Clock3,
  CloudOff,
  Download,
  ClipboardCheck,
  FileSpreadsheet,
  FolderOpen,
  Gauge,
  History,
  LoaderCircle,
  Menu,
  Plus,
  RotateCcw,
  Search,
  ShieldCheck,
  Upload,
  X,
} from "@lucide/vue";
import {
  assess,
  backupBook,
  confirmReading,
  createBook,
  deferReading,
  DEFER_REASONS,
  exceptionRows,
  handoffBook,
  exportWorkbook,
  inspectWorkbook,
  localMonth,
  makeExample,
  progress,
  restoreBackup,
} from "./book.js";
import { BookStorage, openDatabase } from "./storage.js";
import { vModal } from "./modal.js";
import {
  offlineReady,
  offline,
  pwaError,
  updateAvailable,
  installAvailable,
  installApp,
} from "./pwa.js";

const book = ref(null),
  books = ref([]),
  ready = ref(false),
  busy = ref(false);
const storageError = ref(""),
  saveError = ref(""),
  error = ref(""),
  notice = ref("");
const pendingSaves = ref(0),
  savedAt = ref(""),
  sidebar = ref(false),
  mobileDetail = ref(false);
const search = ref(""),
  filter = ref("all"),
  showImport = ref(false),
  candidate = ref(null);
const selectedSheet = ref(""),
  period = ref(localMonth()),
  importError = ref(""),
  showExport = ref(false);
const showHistory = ref(false),
  showHandoff = ref(false),
  fieldError = ref(""),
  draft = ref({ value: "", mode: "normal", note: "", acknowledged: false });
const readingInput = ref(null),
  uploadInput = ref(null),
  backupInput = ref(null);
const smallScreen = ref(matchMedia("(max-width: 760px)").matches);
const modalOpen = computed(
  () =>
    showImport.value ||
    showExport.value ||
    showHistory.value ||
    showHandoff.value,
);
let storage,
  original,
  saveChain = Promise.resolve(),
  revisions = new Map();
const stats = computed(() =>
  book.value
    ? progress(book.value)
    : {
        total: 0,
        recorded: 0,
        pending: 0,
        deferred: 0,
        drafts: 0,
        complete: false,
      },
);
const active = computed(() =>
  book.value?.rows.find((row) => row.id === book.value.activeRowId),
);
const currentIndex = computed(
  () =>
    book.value?.rows.findIndex((row) => row.id === book.value.activeRowId) ??
    -1,
);
const activeEntry = computed(() => book.value?.entries[active.value?.id]);
const exceptions = computed(() =>
  book.value ? exceptionRows(book.value) : [],
);
const exceptionIds = computed(
  () => new Set(exceptions.value.map((row) => row.id)),
);
const dirty = computed(() => Boolean(book.value?.drafts[active.value?.id]));
const assessment = computed(() => {
  if (!active.value || draft.value.value === "") return null;
  try {
    return assess(active.value, draft.value, book.value.ratio);
  } catch {
    return null;
  }
});
const rows = computed(() =>
  (book.value?.rows || []).filter((row) => {
    const state = book.value.entries[row.id]?.status || "pending";
    if (
      filter.value === "exception"
        ? !exceptionIds.value.has(row.id)
        : filter.value !== "all" && state !== filter.value
    )
      return false;
    const text = search.value.trim().toLowerCase();
    return (
      !text ||
      [row.meter, row.name, row.account, row.address, row.excelRow].some(
        (value) => String(value).toLowerCase().includes(text),
      )
    );
  }),
);
const audit = computed(() =>
  (book.value?.history || [])
    .filter((item) => item.rowId === active.value?.id)
    .slice()
    .reverse(),
);
const formatTime = (value) =>
  value
    ? new Date(value).toLocaleString("zh-CN", {
        month: "2-digit",
        day: "2-digit",
        hour: "2-digit",
        minute: "2-digit",
        second: "2-digit",
      })
    : "";
const modeLabel = (mode) =>
  ({ normal: "正常", replaced: "换表", removed: "拆表" })[mode] || "暂缓";
const bookTitle = (value) =>
  value.filename.replace(/\.xlsx$/i, "").replace(/^\d{4}-\d{2}\s*/, "");
const rowState = (row) =>
  book.value.drafts[row.id]
    ? "有修改"
    : { recorded: "已录", deferred: "暂缓" }[
        book.value.entries[row.id]?.status
      ] || "未录";
const resetMessages = () => {
  error.value = "";
  fieldError.value = "";
  notice.value = "";
};

function loadDraft() {
  const entry = activeEntry.value;
  draft.value = {
    ...(book.value.drafts[active.value.id] || {
      value: entry?.status === "recorded" ? String(entry.value) : "",
      mode: entry?.mode || "normal",
      note: entry?.note || "",
      acknowledged: entry?.acknowledged || false,
      deferReason: entry?.reason || "inaccessible",
    }),
  };
  fieldError.value = "";
}
function updateList(snapshot, revision) {
  const item = { id: snapshot.id, revision, book: snapshot };
  const index = books.value.findIndex((value) => value.id === snapshot.id);
  if (index < 0) books.value.unshift(item);
  else books.value[index] = item;
}
function persist() {
  if (!book.value || !storage || saveError.value) return;
  const snapshot = JSON.parse(JSON.stringify(book.value));
  snapshot.updatedAt = new Date().toISOString();
  book.value.updatedAt = snapshot.updatedAt;
  pendingSaves.value++;
  saveChain = saveChain
    .then(async () => {
      if (saveError.value) return;
      const revision = await storage.save(
        snapshot,
        revisions.get(snapshot.id) || 0,
      );
      revisions.set(snapshot.id, revision);
      updateList(snapshot, revision);
      savedAt.value = snapshot.updatedAt;
    })
    .catch((problem) => {
      saveError.value = problem.message || "自动保存失败，请立即下载备份。";
    })
    .finally(() => {
      pendingSaves.value--;
    });
}
async function retrySave() {
  saveError.value = "";
  persist();
  await saveChain;
}
function editDraft(field, value) {
  draft.value[field] = value;
  if (field !== "acknowledged") draft.value.acknowledged = false;
  book.value.drafts[active.value.id] = { ...draft.value };
  fieldError.value = "";
  notice.value = "";
  persist();
}
async function selectRow(row) {
  book.value.activeRowId = row.id;
  loadDraft();
  mobileDetail.value = true;
  persist();
  await nextTick();
  document.querySelector(".detail-body")?.scrollTo({ top: 0 });
  readingInput.value?.focus({ preventScroll: true });
}
function adjacent(direction) {
  const next = book.value.rows[currentIndex.value + direction];
  if (next) selectRow(next);
}
async function saveReading(advance) {
  resetMessages();
  try {
    confirmReading(book.value, active.value.id, draft.value);
    if (advance) {
      const following = book.value.rows
        .slice(currentIndex.value + 1)
        .concat(book.value.rows.slice(0, currentIndex.value));
      const next = following.find(
        (row) =>
          book.value.entries[row.id]?.status !== "recorded" ||
          book.value.drafts[row.id],
      );
      if (next) {
        book.value.activeRowId = next.id;
        loadDraft();
      } else {
        loadDraft();
        notice.value = "全部用户已录入，可以导出本期结果。";
      }
    } else {
      loadDraft();
      notice.value = "修改已确认。";
    }
    persist();
    await saveChain;
    await nextTick();
    if (advance) document.querySelector(".detail-body")?.scrollTo({ top: 0 });
    readingInput.value?.focus({ preventScroll: true });
  } catch (problem) {
    fieldError.value = problem.message;
    await revealError();
  }
}
async function revealError() {
  await nextTick();
  const element = document.querySelector(".field-error[role='alert']");
  element?.scrollIntoView({ block: "center" });
  element?.focus({ preventScroll: true });
}
async function defer() {
  try {
    const reason = draft.value.deferReason || "inaccessible";
    const note =
      draft.value.note.trim() ||
      (reason === "other" ? "" : DEFER_REASONS[reason]);
    deferReading(book.value, active.value.id, note, reason);
    if (!draft.value.value.trim()) delete book.value.drafts[active.value.id];
    loadDraft();
    persist();
    notice.value = "已标记暂缓，正式导出前仍需完成录入。";
    fieldError.value = "";
  } catch (problem) {
    fieldError.value = problem.message;
    await revealError();
  }
}
function discardDraft() {
  delete book.value.drafts[active.value.id];
  loadDraft();
  persist();
}
async function openBook(item) {
  await saveChain;
  if (saveError.value) {
    error.value = "当前进度尚未保存，请先处理保存失败。";
    return;
  }
  try {
    original = await storage.source(item.id);
    const latest = books.value.find((value) => value.id === item.id) || item;
    book.value = JSON.parse(JSON.stringify(latest.book));
    revisions.set(item.id, latest.revision);
    loadDraft();
    resetMessages();
    search.value = "";
    filter.value = "all";
    sidebar.value = false;
    mobileDetail.value = false;
    savedAt.value = book.value.updatedAt;
  } catch (problem) {
    error.value = problem.message;
  }
}
async function beginImport() {
  await saveChain;
  if (saveError.value) {
    error.value = "当前进度尚未保存，请先处理保存失败。";
    return;
  }
  candidate.value = null;
  importError.value = "";
  showImport.value = true;
}
async function chooseFile(event) {
  const file = event.target.files?.[0];
  event.target.value = "";
  if (!file) return;
  busy.value = true;
  importError.value = "";
  try {
    const next = inspectWorkbook(await file.arrayBuffer(), file.name);
    candidate.value = next;
    selectedSheet.value = next.sheets.find((sheet) => !sheet.error).name;
    period.value = localMonth();
  } catch (problem) {
    candidate.value = null;
    importError.value = problem.message;
  } finally {
    busy.value = false;
  }
}
async function useExample() {
  await beginImport();
  if (saveError.value) return;
  const example = makeExample();
  candidate.value = inspectWorkbook(example.bytes, example.filename);
  selectedSheet.value = candidate.value.sheets[0].name;
  period.value = localMonth();
  importError.value = "";
}
async function finishImport() {
  busy.value = true;
  importError.value = "";
  try {
    const next = await createBook(
      candidate.value,
      selectedSheet.value,
      period.value,
    );
    const existing = books.value.find(
      (item) =>
        item.book.sourceHash === next.sourceHash &&
        item.book.sheetName === next.sheetName &&
        item.book.period === next.period,
    );
    if (existing) {
      await openBook(existing);
      showImport.value = false;
      notice.value = "已打开这个工作簿的现有进度。";
      return;
    }
    const revision = await storage.save(next, 0, candidate.value.bytes);
    revisions.set(next.id, revision);
    updateList(next, revision);
    original = new Uint8Array(candidate.value.bytes);
    book.value = next;
    savedAt.value = next.updatedAt;
    loadDraft();
    resetMessages();
    filter.value = "all";
    search.value = "";
    sidebar.value = false;
    showImport.value = false;
    mobileDetail.value = false;
    navigator.storage?.persist?.().catch(() => {});
  } catch (problem) {
    importError.value = problem.message;
  } finally {
    busy.value = false;
  }
}
function download(bytes, filename, type) {
  const url = URL.createObjectURL(new Blob([bytes], { type }));
  const link = document.createElement("a");
  link.href = url;
  link.download = filename;
  link.click();
  setTimeout(() => URL.revokeObjectURL(url), 1000);
}
function backup() {
  try {
    const at = new Date().toISOString();
    download(
      backupBook({ ...book.value, lastBackupAt: at }, original),
      `${book.value.period} ${book.value.sheetName}.zslr.zip`,
      "application/zip",
    );
    book.value.lastBackupAt = at;
    persist();
    notice.value = "备份已生成，包含原始表格、已录记录和未提交草稿。";
  } catch (problem) {
    error.value = problem.message;
  }
}
async function handoff() {
  await saveChain;
  if (saveError.value) {
    error.value = "请先完成本机保存；保存失败时可下载进度备份。";
    return;
  }
  try {
    const at = new Date().toISOString();
    const snapshot = { ...book.value, lastHandoffAt: at };
    download(
      handoffBook(snapshot, original),
      `${book.value.period} ${book.value.sheetName} 交接-${stats.value.complete ? "已完成" : "未完成"}.zslr.zip`,
      "application/zip",
    );
    book.value.lastHandoffAt = at;
    persist();
    showHandoff.value = false;
    notice.value = "交接包已生成，请确认文件已保存在设备中。";
  } catch (problem) {
    error.value = problem.message;
  }
}
async function importBackup(event) {
  const file = event.target.files?.[0];
  event.target.value = "";
  if (!file) return;
  await saveChain;
  if (saveError.value) {
    error.value = "请先处理当前抄表本的保存失败。";
    return;
  }
  busy.value = true;
  try {
    const recovered = await restoreBackup(await file.arrayBuffer());
    const revision = await storage.save(recovered.book, 0, recovered.bytes);
    revisions.set(recovered.book.id, revision);
    updateList(recovered.book, revision);
    original = recovered.bytes;
    book.value = recovered.book;
    loadDraft();
    resetMessages();
    savedAt.value = book.value.updatedAt;
    filter.value = "all";
    search.value = "";
    sidebar.value = false;
    notice.value = "备份已恢复为独立抄表本。";
  } catch (problem) {
    error.value = `恢复失败：${problem.message}`;
  } finally {
    busy.value = false;
  }
}
async function doExport() {
  await saveChain;
  if (saveError.value) {
    error.value = "请先完成本机保存，或下载备份。";
    return;
  }
  try {
    const bytes = exportWorkbook(book.value, original);
    const base = book.value.filename
      .replace(/\.xlsx$/i, "")
      .replace(/\d{4}-\d{2}\s*/, "");
    download(
      bytes,
      `${book.value.period} ${base} 已录.xlsx`,
      "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    );
    showExport.value = false;
    notice.value = "本期结果已生成。";
  } catch (problem) {
    error.value = problem.message;
  }
}
function changeRatio(event) {
  const number = Number(event.target.value);
  if (!Number.isFinite(number) || number < 1 || number > 1000) {
    event.target.value = book.value.ratio;
    return;
  }
  book.value.ratio = number;
  for (const row of book.value.rows) {
    const entry = book.value.entries[row.id];
    if (
      entry?.status === "recorded" &&
      assess(row, entry, number).warning &&
      !entry.acknowledged &&
      !book.value.drafts[row.id]
    )
      book.value.drafts[row.id] = {
        value: String(entry.value),
        mode: entry.mode,
        note: entry.note,
        acknowledged: false,
      };
  }
  loadDraft();
  persist();
}
onMounted(async () => {
  matchMedia("(max-width: 760px)").addEventListener("change", (event) => {
    smallScreen.value = event.matches;
    sidebar.value = false;
  });
  window.addEventListener("keydown", (event) => {
    if (event.key === "Escape") sidebar.value = false;
  });
  window.addEventListener("beforeunload", (event) => {
    if (pendingSaves.value || saveError.value) {
      event.preventDefault();
      event.returnValue = "";
    }
  });
  try {
    storage = new BookStorage(await openDatabase());
    books.value = await storage.list();
    if (books.value.length) await openBook(books.value[0]);
  } catch (problem) {
    storageError.value = `本机存储不可用：${problem.message}`;
  } finally {
    ready.value = true;
  }
});
</script>

<template>
  <div class="app-shell">
    <header class="app-header" :inert="modalOpen">
      <button
        class="icon-button mobile-menu"
        title="抄表本列表"
        aria-label="抄表本列表"
        @click="sidebar = !sidebar"
      >
        <Menu :size="20" />
      </button>
      <div class="brand">
        <span class="brand-mark"><Gauge :size="23" /></span
        ><span>电子抄表本<small>ZSLR</small></span>
      </div>
      <div class="save-state" :class="{ failed: saveError }" role="status">
        <AlertCircle v-if="saveError" :size="15" /><LoaderCircle
          v-else-if="pendingSaves"
          class="spin"
          :size="15"
        /><ShieldCheck v-else :size="15" /><span>{{
          saveError
            ? "保存失败"
            : pendingSaves
              ? "正在保存"
              : book
                ? "已保存到本机"
                : "本机存储"
        }}</span>
      </div>
      <button
        class="button header-import"
        :disabled="!ready || busy || !!storageError"
        @click="beginImport"
      >
        <Plus :size="16" />新建抄表本
      </button>
    </header>
    <div v-if="storageError || saveError" class="global-error" role="alert">
      <AlertCircle :size="19" /><span>{{ storageError || saveError }}</span
      ><button v-if="saveError" class="button" @click="retrySave">
        <RotateCcw :size="15" />重试保存</button
      ><button v-if="book" class="button" @click="backup">
        <Download :size="15" />下载备份
      </button>
    </div>
    <div v-if="error" class="global-error" role="alert">
      <AlertCircle :size="18" /><span>{{ error }}</span
      ><button class="icon-button" aria-label="关闭提示" @click="error = ''">
        <X :size="17" />
      </button>
    </div>
    <div class="workspace" :inert="modalOpen">
      <div v-if="sidebar" class="sidebar-scrim" @click="sidebar = false"></div>
      <aside
        class="sidebar"
        :class="{ open: sidebar }"
        :inert="smallScreen && !sidebar"
      >
        <div class="sidebar-heading">
          <span>我的抄表本</span><span>{{ books.length }}</span>
        </div>
        <nav aria-label="抄表本">
          <button
            v-for="item in books"
            :key="item.id"
            class="book-item"
            :class="{ selected: item.id === book?.id }"
            @click="openBook(item)"
          >
            <BookOpen :size="18" /><span
              ><strong>{{ bookTitle(item.book) }}</strong
              ><small
                >{{ item.book.period }} · {{ progress(item.book).recorded }}/{{
                  item.book.rows.length
                }}
                户 · {{ item.book.sheetName }}</small
              >
              <small v-if="item.book.restoredAt"
                >恢复 {{ formatTime(item.book.restoredAt) }}</small
              >
            </span>
          </button>
          <p v-if="!books.length" class="muted sidebar-empty">暂无抄表本</p>
        </nav>
        <div class="sidebar-bottom">
          <div class="local-label" role="status">
            <CloudOff :size="14" />{{
              offline
                ? "当前离线"
                : pwaError ||
                  (offlineReady ? "离线资源已就绪" : "离线资源未就绪")
            }}
          </div>
          <div v-if="updateAvailable" class="local-label">新版待启用</div>
          <button
            v-if="installAvailable"
            class="text-button"
            @click="installApp"
          >
            <Download :size="17" />安装抄表本
          </button>
          <button
            class="text-button"
            :disabled="busy || !ready || !!storageError"
            @click="backupInput?.click()"
          >
            <FolderOpen :size="17" />恢复备份
          </button>
          <div class="local-label">
            <CloudOff :size="14" />数据保存在此浏览器
          </div>
        </div>
      </aside>
      <main
        v-if="!book"
        class="empty-workspace"
        :inert="sidebar && smallScreen"
      >
        <div class="page-heading">
          <div>
            <p class="eyebrow">抄表作业</p>
            <h1>电子抄表本</h1>
          </div>
        </div>
        <div class="empty-table-head">
          <span>用户 / 地址</span><span>上期指数</span><span>本期指数</span
          ><span>状态</span>
        </div>
        <div class="empty-state">
          <div class="empty-icon">
            <FileSpreadsheet :size="40" :stroke-width="1.4" />
          </div>
          <h2>{{ ready ? "还没有抄表本" : "正在恢复进度" }}</h2>
          <div class="empty-actions">
            <button
              class="button primary"
              :disabled="!ready || !!storageError"
              @click="beginImport"
            >
              <Upload :size="17" />导入工作簿</button
            ><button
              class="button"
              :disabled="!ready || !!storageError"
              @click="useExample"
            >
              打开示例
            </button>
          </div>
        </div>
      </main>
      <main v-else class="book-workspace" :inert="sidebar && smallScreen">
        <div class="page-heading">
          <div class="heading-copy">
            <p class="eyebrow">{{ book.period }} · 抄表作业</p>
            <h1>{{ bookTitle(book) }}</h1>
            <p class="source-name" :title="book.filename">
              {{ book.sheetName }} · {{ book.filename }}
            </p>
          </div>
          <div class="heading-actions">
            <button class="button" @click="showHandoff = true">
              <ClipboardCheck :size="16" />收工交接
            </button>
            <button class="button" @click="backup">
              <Download :size="16" />备份</button
            ><button class="button primary" @click="showExport = true">
              <ArrowDownToLine :size="16" />导出结果
            </button>
          </div>
        </div>
        <div class="backup-status">
          <span
            >最近备份生成：{{
              formatTime(book.lastBackupAt) || "尚未备份"
            }}</span
          >
          <span
            >最近交接生成：{{
              formatTime(book.lastHandoffAt) || "尚未交接"
            }}</span
          >
          <span v-if="offline" class="offline-indicator"
            ><CloudOff :size="14" />离线</span
          >
        </div>
        <section class="progress-band" aria-label="录入进度">
          <div class="progress-total">
            <strong
              >{{ stats.recorded }}<span> / {{ stats.total }}</span></strong
            ><small>已录户数</small>
          </div>
          <div class="progress-track">
            <div
              :style="{ width: `${(stats.recorded / stats.total) * 100}%` }"
            ></div>
          </div>
          <span
            ><i class="dot pending"></i>未录 <b>{{ stats.pending }}</b></span
          ><span
            ><i class="dot deferred"></i>暂缓 <b>{{ stats.deferred }}</b></span
          ><span v-if="stats.drafts" class="draft-count"
            >未确认修改 {{ stats.drafts }}</span
          >
        </section>
        <div v-if="notice" class="notice" role="status">
          <Check :size="16" />{{ notice
          }}<button
            class="icon-button"
            aria-label="关闭消息"
            @click="notice = ''"
          >
            <X :size="15" />
          </button>
        </div>
        <div class="record-workspace" :class="{ 'show-detail': mobileDetail }">
          <section class="record-list" aria-label="用户列表">
            <div class="list-toolbar">
              <div class="search-field">
                <Search :size="17" /><input
                  v-model="search"
                  type="search"
                  placeholder="搜索姓名、抄表号、地址"
                  aria-label="搜索用户"
                />
              </div>
            </div>
            <div class="filter-tabs" role="tablist" aria-label="录入状态">
              <button
                v-for="tab in [
                  { key: 'all', label: '全部', count: stats.total },
                  { key: 'pending', label: '未录', count: stats.pending },
                  { key: 'deferred', label: '暂缓', count: stats.deferred },
                  { key: 'recorded', label: '已录', count: stats.recorded },
                  { key: 'exception', label: '异常', count: exceptions.length },
                ]"
                :key="tab.key"
                role="tab"
                :aria-selected="filter === tab.key"
                :class="{ active: filter === tab.key }"
                @click="filter = tab.key"
              >
                {{ tab.label }}<span>{{ tab.count }}</span>
              </button>
            </div>
            <div class="record-scroll">
              <div class="list-columns">
                <span>用户 / 水表地址</span><span>指数</span><span>状态</span>
              </div>
              <button
                v-for="row in rows"
                :key="row.id"
                class="record-row"
                :class="{ active: row.id === active?.id }"
                :aria-current="row.id === active?.id ? 'true' : undefined"
                @click="selectRow(row)"
              >
                <span class="row-person"
                  ><strong
                    >{{ row.name || row.meter }}
                    <small v-if="row.duplicate" class="duplicate-label"
                      >重号</small
                    ></strong
                  ><span>{{ row.address || row.meter }}</span
                  ><small>表 {{ row.meter }} · 行 {{ row.excelRow }}</small
                  ><small v-if="filter === 'exception'"
                    >{{
                      book.entries[row.id]?.status === "deferred"
                        ? DEFER_REASONS[book.entries[row.id].reason || "other"]
                        : modeLabel(book.entries[row.id]?.mode)
                    }}
                    · {{ book.entries[row.id]?.note }}</small
                  ></span
                ><span class="row-reading"
                  ><strong>{{
                    book.entries[row.id]?.status === "recorded"
                      ? book.entries[row.id].value
                      : "—"
                  }}</strong
                  ><small>上期 {{ row.baseline }}</small></span
                ><span
                  class="status-tag"
                  :class="
                    book.drafts[row.id]
                      ? 'draft'
                      : book.entries[row.id]?.status || 'pending'
                  "
                  >{{ rowState(row) }}</span
                >
              </button>
              <div v-if="!rows.length" class="list-empty">
                <Search :size="25" />
                <p>没有匹配的用户</p>
                <button
                  class="text-button"
                  @click="
                    search = '';
                    filter = 'all';
                  "
                >
                  显示全部
                </button>
              </div>
            </div>
            <footer class="list-footer">
              {{ rows.length }} 户<span>{{ formatTime(savedAt) }}</span>
            </footer>
          </section>
          <section v-if="active" class="record-detail" aria-label="录入读数">
            <div class="detail-top">
              <button
                class="text-button back-list"
                @click="mobileDetail = false"
              >
                <ArrowLeft :size="17" />用户列表</button
              ><span
                >第 {{ currentIndex + 1 }} / {{ stats.total }} 户 · Excel 行
                {{ active.excelRow }}</span
              ><button
                class="icon-button"
                title="修改记录"
                aria-label="修改记录"
                @click="showHistory = true"
              >
                <History :size="18" />
              </button>
            </div>
            <div class="detail-body">
              <div class="person-heading">
                <div>
                  <h2>{{ active.name || active.meter }}</h2>
                  <p>{{ active.address || "未填写地址" }}</p>
                </div>
                <span
                  class="status-tag"
                  :class="activeEntry?.status || 'pending'"
                  >{{ rowState(active) }}</span
                >
              </div>
              <dl class="identity">
                <div>
                  <dt>抄表号</dt>
                  <dd>{{ active.meter }}</dd>
                </div>
                <div>
                  <dt>计费号</dt>
                  <dd>{{ active.account || "—" }}</dd>
                </div>
              </dl>
              <div v-if="active.duplicate" class="warning compact">
                <AlertCircle :size="17" /><span
                  >存在同号水表，请按地址和 Excel 行号核对。</span
                >
              </div>
              <p v-if="active.baselineFormula" class="formula-note">
                上期指数取自原表公式的已保存结果。
              </p>
              <div class="baseline">
                <div>
                  <span>上期指数</span><strong>{{ active.baseline }}</strong
                  ><small>{{ active.month || "月份未提供" }}</small>
                </div>
                <div>
                  <span>上期用量</span
                  ><strong
                    >{{ active.previousUsage ?? "—"
                    }}<small v-if="active.previousUsage !== null">
                      m³</small
                    ></strong
                  ><small>{{ active.historyMonth || "无有效历史记录" }}</small>
                </div>
              </div>
              <form id="reading-form" @submit.prevent="saveReading(true)">
                <label class="field-label" for="current-reading"
                  >本期指数 <span>{{ book.period }}</span></label
                >
                <div class="reading-field">
                  <input
                    id="current-reading"
                    ref="readingInput"
                    :value="draft.value"
                    inputmode="decimal"
                    maxlength="40"
                    autocomplete="off"
                    placeholder="输入读数"
                    :aria-invalid="!!fieldError"
                    @input="editDraft('value', $event.target.value)"
                  /><span>m³</span>
                </div>
                <div class="usage-line">
                  <span>本期用量</span
                  ><strong
                    >{{ assessment?.usage ?? "—" }}
                    <small v-if="assessment?.usage != null">m³</small></strong
                  ><span v-if="draft.mode === 'replaced'">换表需单独核算</span>
                </div>
                <fieldset class="mode-control">
                  <legend>抄表情况</legend>
                  <label
                    v-for="mode in ['normal', 'replaced', 'removed']"
                    :key="mode"
                    :class="{ selected: draft.mode === mode }"
                    ><input
                      type="radio"
                      name="mode"
                      :value="mode"
                      :checked="draft.mode === mode"
                      @change="editDraft('mode', mode)"
                    />{{ modeLabel(mode) }}</label
                  >
                </fieldset>
                <label class="field-label" for="remark"
                  >备注<span>{{
                    draft.mode === "normal" ? "选填" : "必填"
                  }}</span></label
                ><textarea
                  id="remark"
                  :value="draft.note"
                  rows="2"
                  maxlength="500"
                  placeholder="现场情况、换表或暂缓原因"
                  @input="editDraft('note', $event.target.value)"
                ></textarea>
                <p v-if="active.fields.备注信息" class="previous-note">
                  原备注：{{ active.fields.备注信息 }}
                </p>
                <div v-if="assessment?.warning" class="warning">
                  <AlertCircle :size="18" />
                  <div>
                    <p>{{ assessment.warning }}</p>
                    <label class="acknowledgment"
                      ><input
                        type="checkbox"
                        :checked="draft.acknowledged"
                        @change="
                          editDraft('acknowledged', $event.target.checked)
                        "
                      />已核对现场读数，确认录入</label
                    >
                  </div>
                </div>
                <div
                  v-if="fieldError"
                  class="field-error"
                  role="alert"
                  tabindex="-1"
                >
                  <AlertCircle :size="16" />{{ fieldError }}
                </div>
                <div
                  v-if="activeEntry?.status === 'recorded'"
                  class="submit-actions"
                >
                  <button
                    class="button"
                    type="button"
                    :disabled="!!saveError"
                    @click="saveReading(false)"
                  >
                    <Check :size="17" />保存修改
                  </button>
                </div>
                <div class="secondary-actions">
                  <select
                    aria-label="暂缓类型"
                    :value="draft.deferReason || 'inaccessible'"
                    @change="editDraft('deferReason', $event.target.value)"
                  >
                    <option
                      v-for="(label, key) in DEFER_REASONS"
                      :key="key"
                      :value="key"
                    >
                      {{ label }}
                    </option>
                  </select>
                  <button
                    type="button"
                    class="text-button"
                    :disabled="!!saveError"
                    @click="defer"
                  >
                    <Clock3 :size="16" />暂缓</button
                  ><button
                    v-if="dirty"
                    type="button"
                    class="text-button"
                    @click="discardDraft"
                  >
                    <RotateCcw :size="15" />撤销未确认修改
                  </button>
                </div>
              </form>
              <label class="threshold-control"
                >异常阈值
                <input
                  type="number"
                  min="1"
                  max="1000"
                  :value="book.ratio"
                  aria-label="异常阈值"
                  @change="changeRatio"
                />
                %
              </label>
              <details class="source-fields">
                <summary>原始表格字段<ChevronRight :size="15" /></summary>
                <dl>
                  <div v-for="(value, key) in active.fields" :key="key">
                    <dt>{{ key }}</dt>
                    <dd>{{ value || "—" }}</dd>
                  </div>
                </dl>
              </details>
            </div>
            <footer class="detail-footer">
              <button
                class="button"
                :disabled="currentIndex === 0"
                @click="adjacent(-1)"
              >
                <ArrowLeft :size="16" />上一户
              </button>
              <button
                class="button primary"
                type="submit"
                form="reading-form"
                :disabled="!!saveError"
              >
                <CheckCheck :size="17" />确认并下一户
              </button>
              <button
                class="icon-button"
                title="下一户"
                aria-label="下一户"
                :disabled="currentIndex === stats.total - 1"
                @click="adjacent(1)"
              >
                <ArrowRight :size="18" />
              </button>
            </footer>
          </section>
        </div>
      </main>
    </div>
    <input
      ref="uploadInput"
      class="hidden"
      type="file"
      accept=".xlsx"
      aria-label="选择工作簿"
      @change="chooseFile"
    /><input
      ref="backupInput"
      class="hidden"
      type="file"
      accept=".zip"
      aria-label="选择进度备份"
      @change="importBackup"
    />
    <div
      v-if="showImport"
      class="modal-backdrop"
      @keydown.esc="!busy && (showImport = false)"
    >
      <section
        v-modal
        class="modal"
        role="dialog"
        aria-modal="true"
        aria-labelledby="import-title"
      >
        <header>
          <h2 id="import-title">新建抄表本</h2>
          <button
            class="icon-button"
            aria-label="关闭导入"
            :disabled="busy"
            @click="showImport = false"
          >
            <X :size="20" />
          </button>
        </header>
        <div class="modal-body">
          <button
            class="file-picker"
            :disabled="busy"
            @click="uploadInput?.click()"
          >
            <FileSpreadsheet :size="30" /><strong>{{
              candidate?.filename || "选择工作簿"
            }}</strong
            ><span>.xlsx · 最大 20 MB</span>
          </button>
          <p v-if="importError" class="field-error" role="alert">
            {{ importError }}
          </p>
          <template v-if="candidate"
            ><label class="field-label" for="sheet-select">录入工作表</label
            ><select id="sheet-select" v-model="selectedSheet">
              <option
                v-for="sheet in candidate.sheets"
                :key="sheet.name"
                :value="sheet.name"
                :disabled="!!sheet.error"
              >
                {{ sheet.name }} · {{ sheet.error || `${sheet.count} 户` }}
              </option>
            </select>
            <p class="import-note">仅更新所选工作表，其他工作表保留原样。</p>
            <label class="field-label" for="period">本期月份</label
            ><input id="period" v-model="period" type="month" />
            <div
              v-if="
                candidate.sheets.find((sheet) => sheet.name === selectedSheet)
                  ?.duplicates
              "
              class="warning compact"
            >
              <AlertCircle :size="17" /><span
                >含重复抄表号，将按工作表行号分别记录。</span
              >
            </div></template
          >
        </div>
        <footer>
          <button class="button" :disabled="busy" @click="showImport = false">
            取消</button
          ><button
            class="button primary"
            :disabled="!candidate || busy"
            @click="finishImport"
          >
            <LoaderCircle v-if="busy" class="spin" :size="17" /><Plus
              v-else
              :size="17"
            />开始录入
          </button>
        </footer>
      </section>
    </div>
    <div
      v-if="showExport"
      class="modal-backdrop"
      @keydown.esc="showExport = false"
    >
      <section
        v-modal
        class="modal"
        role="dialog"
        aria-modal="true"
        aria-labelledby="export-title"
      >
        <header>
          <h2 id="export-title">导出本期结果</h2>
          <button
            class="icon-button"
            aria-label="关闭导出"
            @click="showExport = false"
          >
            <X :size="20" />
          </button>
        </header>
        <div class="modal-body">
          <dl class="export-summary">
            <div>
              <dt>工作表</dt>
              <dd>{{ book.sheetName }}</dd>
            </div>
            <div>
              <dt>本期月份</dt>
              <dd>{{ book.period }}</dd>
            </div>
            <div>
              <dt>已录入</dt>
              <dd>{{ stats.recorded }} / {{ stats.total }} 户</dd>
            </div>
            <div>
              <dt>未录 / 暂缓</dt>
              <dd>{{ stats.pending }} / {{ stats.deferred }} 户</dd>
            </div>
            <div>
              <dt>未确认修改</dt>
              <dd>{{ stats.drafts }} 户</dd>
            </div>
          </dl>
          <p v-if="!stats.complete" class="warning">
            尚有未完成记录。可先下载进度备份。
          </p>
          <p v-else class="notice"><CheckCheck :size="18" />全部记录已确认。</p>
        </div>
        <footer>
          <button class="button" @click="backup">
            <Download :size="16" />备份进度</button
          ><button
            class="button primary"
            :disabled="!stats.complete || !!saveError"
            @click="doExport"
          >
            <ArrowDownToLine :size="16" />生成 Excel
          </button>
        </footer>
      </section>
    </div>
    <div
      v-if="showHandoff"
      class="modal-backdrop"
      @keydown.esc="showHandoff = false"
    >
      <section
        v-modal
        class="modal"
        role="dialog"
        aria-modal="true"
        aria-labelledby="handoff-title"
      >
        <header>
          <h2 id="handoff-title">收工交接</h2>
          <button
            class="icon-button"
            aria-label="关闭交接"
            @click="showHandoff = false"
          >
            <X :size="20" />
          </button>
        </header>
        <div class="modal-body">
          <dl class="export-summary">
            <div>
              <dt>本期 / 工作表</dt>
              <dd>{{ book.period }} / {{ book.sheetName }}</dd>
            </div>
            <div>
              <dt>已录 / 总户数</dt>
              <dd>{{ stats.recorded }} / {{ stats.total }}</dd>
            </div>
            <div>
              <dt>未录 / 暂缓</dt>
              <dd>{{ stats.pending }} / {{ stats.deferred }}</dd>
            </div>
            <div>
              <dt>未确认修改 / 异常复核</dt>
              <dd>{{ stats.drafts }} / {{ exceptions.length }}</dd>
            </div>
            <div>
              <dt>交接状态</dt>
              <dd>{{ stats.complete ? "已完成" : "未完成" }}</dd>
            </div>
          </dl>
          <p v-if="!stats.complete" class="warning">
            未完成交接，不含正式计费结果。
          </p>
          <ul class="handoff-files">
            <li>原始工作簿与可恢复进度</li>
            <li>交接明细、未完成清单、异常复核清单</li>
            <li v-if="stats.complete">已完成结果工作簿</li>
          </ul>
        </div>
        <footer>
          <button class="button" @click="showHandoff = false">返回录入</button
          ><button
            class="button primary"
            :disabled="!!saveError"
            @click="handoff"
          >
            <Download :size="17" />生成交接包
          </button>
        </footer>
      </section>
    </div>
    <div
      v-if="showHistory"
      class="modal-backdrop"
      @keydown.esc="showHistory = false"
    >
      <section
        v-modal
        class="modal"
        role="dialog"
        aria-modal="true"
        aria-labelledby="history-title"
      >
        <header>
          <h2 id="history-title">修改记录 · 行 {{ active.excelRow }}</h2>
          <button
            class="icon-button"
            aria-label="关闭修改记录"
            @click="showHistory = false"
          >
            <X :size="20" />
          </button>
        </header>
        <div class="modal-body">
          <p v-if="!audit.length" class="muted">暂无已确认的修改</p>
          <ol class="audit-list">
            <li v-for="(item, index) in audit" :key="index">
              <time>{{ formatTime(item.at) }}</time
              ><strong
                >{{ item.before?.value ?? "未录" }} →
                {{ item.after?.value ?? "暂缓" }}</strong
              >
              <p>
                {{ modeLabel(item.after?.mode)
                }}{{ item.after?.note ? ` · ${item.after.note}` : "" }}
              </p>
            </li>
          </ol>
        </div>
      </section>
    </div>
  </div>
</template>
