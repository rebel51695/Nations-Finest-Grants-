import { doc, getDoc, setDoc, deleteDoc, collection, getDocs } from "firebase/firestore";
import { db } from "./firebaseConfig";

const COLLECTION = "grantflow_data";

// Mimics the Claude-artifact window.storage API so the existing app code
// doesn't need to change: get/set/delete/list, each with a `shared` flag.
// shared:true  -> stored in Firestore, visible to everyone
// shared:false -> stored in this browser only (localStorage), e.g. "who am I"

async function storageGet(key, shared = false) {
  if (!shared) {
    const v = localStorage.getItem(key);
    return v !== null ? { key, value: v, shared: false } : null;
  }
  const safeKey = encodeURIComponent(key);
  const snap = await getDoc(doc(db, COLLECTION, safeKey));
  if (!snap.exists()) return null;
  return { key, value: snap.data().value, shared: true };
}

async function storageSet(key, value, shared = false) {
  if (!shared) {
    localStorage.setItem(key, value);
    return { key, value, shared: false };
  }
  const safeKey = encodeURIComponent(key);
  await setDoc(doc(db, COLLECTION, safeKey), { value, updatedAt: new Date().toISOString() });
  return { key, value, shared: true };
}

async function storageDelete(key, shared = false) {
  if (!shared) {
    localStorage.removeItem(key);
    return { key, deleted: true, shared: false };
  }
  const safeKey = encodeURIComponent(key);
  await deleteDoc(doc(db, COLLECTION, safeKey));
  return { key, deleted: true, shared: true };
}

async function storageList(prefix = "", shared = false) {
  if (!shared) {
    const keys = Object.keys(localStorage).filter((k) => k.startsWith(prefix));
    return { keys, prefix, shared: false };
  }
  const snap = await getDocs(collection(db, COLLECTION));
  const keys = snap.docs.map((d) => decodeURIComponent(d.id)).filter((k) => k.startsWith(prefix));
  return { keys, prefix, shared: true };
}

// Attach to window so the existing GrantFlow.jsx code (written for the
// Claude artifact environment) works completely unchanged.
if (typeof window !== "undefined") {
  window.storage = {
    get: storageGet,
    set: storageSet,
    delete: storageDelete,
    list: storageList,
  };
}
