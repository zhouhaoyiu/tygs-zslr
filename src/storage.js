const DB_NAME = "zslr-meter-books";
export function openDatabase() {
  return new Promise((resolve, reject) => {
    const request = indexedDB.open(DB_NAME, 1);
    request.onupgradeneeded = () => {
      request.result.createObjectStore("books", { keyPath: "id" });
      request.result.createObjectStore("sources");
    };
    request.onsuccess = () => resolve(request.result);
    request.onerror = () => reject(request.error);
    request.onblocked = () =>
      reject(new Error("请关闭其他旧版抄表本页面后重试。"));
  });
}

export class BookStorage {
  constructor(db) {
    this.db = db;
  }
  async list() {
    return new Promise((resolve, reject) => {
      const request = this.db
        .transaction("books")
        .objectStore("books")
        .getAll();
      request.onsuccess = () =>
        resolve(
          request.result.sort((a, b) =>
            b.book.updatedAt.localeCompare(a.book.updatedAt),
          ),
        );
      request.onerror = () => reject(request.error);
    });
  }
  async source(id) {
    return new Promise((resolve, reject) => {
      const request = this.db
        .transaction("sources")
        .objectStore("sources")
        .get(id);
      request.onsuccess = () =>
        request.result
          ? resolve(request.result)
          : reject(new Error("原始工作簿丢失，请从备份恢复。"));
      request.onerror = () => reject(request.error);
    });
  }
  async save(book, revision, source) {
    const snapshot = JSON.parse(JSON.stringify(book));
    return new Promise((resolve, reject) => {
      const tx = this.db.transaction(
        source ? ["books", "sources"] : ["books"],
        "readwrite",
      );
      const store = tx.objectStore("books");
      let problem = null;
      const get = store.get(book.id);
      get.onsuccess = () => {
        if ((get.result?.revision ?? 0) !== revision) {
          problem = new Error(
            "另一页面已修改这个抄表本。请先备份当前进度，再刷新载入最新记录。",
          );
          tx.abort();
          return;
        }
        if (!get.result && !source) {
          problem = new Error("原始工作簿尚未保存。");
          tx.abort();
          return;
        }
        try {
          store.put({ id: book.id, revision: revision + 1, book: snapshot });
          if (source)
            tx.objectStore("sources").put(new Uint8Array(source), book.id);
        } catch (error) {
          problem = error;
          tx.abort();
        }
      };
      tx.oncomplete = () => resolve(revision + 1);
      tx.onerror = () =>
        reject(problem || tx.error || new Error("无法保存到本机。"));
      tx.onabort = () =>
        reject(problem || tx.error || new Error("保存被中断。"));
    });
  }
}
