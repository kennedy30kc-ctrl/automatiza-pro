import { initializeApp } from "firebase/app";
import { getFirestore, doc, getDoc, setDoc, deleteDoc, collection, getDocs } from "firebase/firestore";

const firebaseConfig = {
  apiKey: "AIzaSyCihidpoOeuc4kOKRJBplWBQpXgJHHluGk",
  authDomain: "automatiza-pro.firebaseapp.com",
  projectId: "automatiza-pro",
  storageBucket: "automatiza-pro.firebasestorage.app",
  messagingSenderId: "470358165821",
  appId: "1:470358165821:web:db093869626cce096f1d54"
};

const app = initializeApp(firebaseConfig);
const db = getFirestore(app);

export const storage = {
  async get(key, shared = true) {
    try {
      const col = shared ? "shared" : "personal";
      const ref = doc(db, col, key.replace(/[/\\]/g, "_"));
      const snap = await getDoc(ref);
      return snap.exists() ? { key, value: snap.data().value, shared } : null;
    } catch { return null; }
  },
  async set(key, value, shared = true) {
    try {
      const col = shared ? "shared" : "personal";
      const ref = doc(db, col, key.replace(/[/\\]/g, "_"));
      await setDoc(ref, { key, value, shared, updatedAt: Date.now() });
      return { key, value, shared };
    } catch { return null; }
  },
  async delete(key, shared = true) {
    try {
      const col = shared ? "shared" : "personal";
      const ref = doc(db, col, key.replace(/[/\\]/g, "_"));
      await deleteDoc(ref);
      return { key, deleted: true, shared };
    } catch { return null; }
  },
  async list(prefix = "", shared = true) {
    try {
      const col = shared ? "shared" : "personal";
      const snap = await getDocs(collection(db, col));
      const keys = snap.docs.map(d => d.data().key).filter(k => k && k.startsWith(prefix));
      return { keys, prefix, shared };
    } catch { return { keys: [] }; }
  }
};

export default db;