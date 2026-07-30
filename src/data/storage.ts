import { ProjectData } from "./types";

const DB_NAME = "rpa-project-control-center";
const STORE = "active-project";
const BACKUPS = "backups";

function openDb(): Promise<IDBDatabase> {
  return new Promise((resolve, reject) => {
    const request = indexedDB.open(DB_NAME, 1);
    request.onupgradeneeded = () => {
      const db = request.result;
      if (!db.objectStoreNames.contains(STORE)) db.createObjectStore(STORE);
      if (!db.objectStoreNames.contains(BACKUPS)) db.createObjectStore(BACKUPS, { keyPath: "uploadedAt" });
    };
    request.onerror = () => reject(request.error);
    request.onsuccess = () => resolve(request.result);
  });
}

export async function getActiveProject(): Promise<ProjectData | undefined> {
  const db = await openDb();
  return new Promise((resolve, reject) => {
    const tx = db.transaction(STORE, "readonly");
    const request = tx.objectStore(STORE).get("current");
    request.onerror = () => reject(request.error);
    request.onsuccess = () => resolve(request.result as ProjectData | undefined);
  });
}

export async function saveActiveProject(data: ProjectData, previous?: ProjectData): Promise<void> {
  const db = await openDb();
  await new Promise<void>((resolve, reject) => {
    const tx = db.transaction([STORE, BACKUPS], "readwrite");
    tx.objectStore(STORE).put(data, "current");
    if (previous) tx.objectStore(BACKUPS).put(previous);
    tx.onerror = () => reject(tx.error);
    tx.oncomplete = () => resolve();
  });
}

export async function clearActiveProject(): Promise<void> {
  const db = await openDb();
  await new Promise<void>((resolve, reject) => {
    const tx = db.transaction(STORE, "readwrite");
    tx.objectStore(STORE).delete("current");
    tx.onerror = () => reject(tx.error);
    tx.oncomplete = () => resolve();
  });
}
