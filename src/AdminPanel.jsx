import { useState, useEffect } from "react";
import { collection, getDocs, doc, setDoc, getDoc, deleteDoc } from "firebase/firestore";
import { db } from "./firebaseConfig";
import { ADMIN_EMAILS } from "./adminEmails";

const MODULES = [
  { key: "dashboard", label: "Dashboard" },
  { key: "grants", label: "Grants" },
  { key: "budgets", label: "Budgets" },
  { key: "invoicing", label: "Invoicing" },
  { key: "tasks", label: "Tasks" },
  { key: "grant-reports", label: "Grant Reports" },
  { key: "reporting", label: "Reporting" },
  { key: "org-budget", label: "Org Budget" },
  { key: "scenarios", label: "Scenarios" },
  { key: "burn-rate", label: "Burn Rate" },
  { key: "personnel", label: "Personnel" },
  { key: "activity-log", label: "Activity Log" },
  { key: "trash", label: "Trash" },
  { key: "data", label: "Data & Backup" },
];

export default function AdminPanel({ currentUserEmail }) {
  const [users, setUsers] = useState([]);
  const [loading, setLoading] = useState(true);
  const [error, setError] = useState("");
  const [expandedId, setExpandedId] = useState(null);
  const [announcement, setAnnouncement] = useState(null);
  const [draftMessage, setDraftMessage] = useState("");
  const [savingAnnouncement, setSavingAnnouncement] = useState(false);

  const loadAnnouncement = async () => {
    try {
      const snap = await getDoc(doc(db, "app_config", "announcement"));
      if (snap.exists() && snap.data().message) {
        const parsed = snap.data();
        setAnnouncement(parsed);
        setDraftMessage(parsed?.message || "");
      }
    } catch (e) { /* none set yet */ }
  };

  const postAnnouncement = async () => {
    if (!draftMessage.trim()) return;
    setSavingAnnouncement(true);
    const payload = { message: draftMessage.trim(), setBy: currentUserEmail || "", setAt: new Date().toISOString() };
    await setDoc(doc(db, "app_config", "announcement"), payload);
    setAnnouncement(payload);
    setSavingAnnouncement(false);
  };

  const clearAnnouncement = async () => {
    setSavingAnnouncement(true);
    await deleteDoc(doc(db, "app_config", "announcement"));
    setAnnouncement(null);
    setDraftMessage("");
    setSavingAnnouncement(false);
  };

  const load = async () => {
    setLoading(true);
    setError("");
    try {
      const snap = await getDocs(collection(db, "user_access"));
      const list = snap.docs.map((d) => ({ id: d.id, role: "editor", disabledModules: [], ...d.data() }));
      list.sort((a, b) => (b.requestedAt || "").localeCompare(a.requestedAt || ""));
      setUsers(list);
    } catch (e) {
      setError("Couldn't load the user list.");
    }
    setLoading(false);
  };

  useEffect(() => { load(); loadAnnouncement(); }, []);

  const setAllowed = async (id, allowed) => {
    await setDoc(doc(db, "user_access", id), { allowed }, { merge: true });
    setUsers((prev) => prev.map((u) => (u.id === id ? { ...u, allowed } : u)));
  };

  const setRole = async (id, role) => {
    await setDoc(doc(db, "user_access", id), { role }, { merge: true });
    setUsers((prev) => prev.map((u) => (u.id === id ? { ...u, role } : u)));
  };

  const toggleModule = async (id, moduleKey) => {
    const user = users.find((u) => u.id === id);
    const current = user.disabledModules || [];
    const next = current.includes(moduleKey) ? current.filter((k) => k !== moduleKey) : [...current, moduleKey];
    await setDoc(doc(db, "user_access", id), { disabledModules: next }, { merge: true });
    setUsers((prev) => prev.map((u) => (u.id === id ? { ...u, disabledModules: next } : u)));
  };

  const fmtDate = (iso) => (iso ? new Date(iso).toLocaleString("en-US", { month: "short", day: "numeric", year: "numeric", hour: "numeric", minute: "2-digit" }) : "—");

  return (
    <div className="space-y-4">
      <div>
        <h1 className="font-display text-2xl" style={{ color: "#1C2624" }}>User Access</h1>
        <p className="text-sm mt-1" style={{ color: "#5B6B66" }}>Approve access, set roles, and control which modules each person can see</p>
      </div>

      <div className="bg-white rounded-lg border p-4 text-sm" style={{ borderColor: "#E1E5DE" }}>
        <div className="font-medium mb-1" style={{ color: "#1C2624" }}>Admins (always have full access, fixed and cannot be changed here)</div>
        <div style={{ color: "#5B6B66" }}>{ADMIN_EMAILS.join(", ")}</div>
      </div>

      <div className="bg-white rounded-lg border p-4" style={{ borderColor: "#E1E5DE" }}>
        <div className="font-medium mb-1 text-sm" style={{ color: "#1C2624" }}>Global announcement banner</div>
        <p className="text-xs mb-3" style={{ color: "#8A8F87" }}>
          Shows to everyone the next time they open or refresh the app. Stays up until an admin removes it.
        </p>
        {announcement && (
          <div className="rounded-md px-3 py-2 mb-3 text-sm" style={{ background: "#FFF7E6", border: "1px solid #F0B21E", color: "#5B4A0F" }}>
            Currently showing: "{announcement.message}"
          </div>
        )}
        <div className="flex items-center gap-2">
          <input
            value={draftMessage}
            onChange={(e) => setDraftMessage(e.target.value)}
            placeholder="e.g. System will be briefly unavailable Friday 5pm for maintenance"
            className="flex-1 rounded-md border px-3 py-2 text-sm"
            style={{ borderColor: "#E1E5DE" }}
          />
          <button
            onClick={postAnnouncement}
            disabled={savingAnnouncement || !draftMessage.trim()}
            className="text-xs px-3 py-2 rounded-md text-white shrink-0"
            style={{ background: "#1F5C6B", opacity: savingAnnouncement || !draftMessage.trim() ? 0.6 : 1 }}
          >
            {announcement ? "Update message" : "Post message"}
          </button>
          {announcement && (
            <button
              onClick={clearAnnouncement}
              disabled={savingAnnouncement}
              className="text-xs px-3 py-2 rounded-md border shrink-0"
              style={{ borderColor: "#B5443A", color: "#B5443A" }}
            >
              Remove
            </button>
          )}
        </div>
      </div>

      <button onClick={load} className="text-xs px-3 py-1.5 rounded-md border" style={{ borderColor: "#E1E5DE", color: "#1C2624" }}>
        Refresh list
      </button>

      {error && <p className="text-sm" style={{ color: "#B5443A" }}>{error}</p>}

      {loading ? (
        <div className="bg-white rounded-lg border p-10 text-center" style={{ borderColor: "#E1E5DE", color: "#8A8F87" }}>Loading…</div>
      ) : users.length === 0 ? (
        <div className="bg-white rounded-lg border p-10 text-center" style={{ borderColor: "#E1E5DE", color: "#8A8F87" }}>
          No one has requested access yet — once someone creates an account, they'll show up here.
        </div>
      ) : (
        <div className="bg-white rounded-lg border divide-y" style={{ borderColor: "#E1E5DE" }}>
          {users.map((u) => (
            <div key={u.id} className="px-4 py-3 text-sm">
              <div className="flex items-center justify-between flex-wrap gap-2">
                <div>
                  <div style={{ color: "#1C2624" }}>{u.email || u.id}</div>
                  <div className="text-xs" style={{ color: "#8A8F87" }}>Requested {fmtDate(u.requestedAt)}</div>
                </div>
                <div className="flex items-center gap-2 flex-wrap">
                  <span
                    className="text-xs px-2 py-0.5 rounded-full"
                    style={{
                      background: u.allowed ? "#EAF1EC" : "#FBEAE8",
                      color: u.allowed ? "#2F6F53" : "#B5443A",
                    }}
                  >
                    {u.allowed ? "Approved" : "Pending / Denied"}
                  </span>
                  {u.allowed && (
                    <select
                      value={u.role || "editor"}
                      onChange={(e) => setRole(u.id, e.target.value)}
                      className="text-xs rounded-md border px-2 py-1.5"
                      style={{ borderColor: "#E1E5DE", color: "#1C2624" }}
                    >
                      <option value="editor">Editor (read + edit)</option>
                      <option value="viewer">Viewer (read-only)</option>
                    </select>
                  )}
                  {u.allowed && (
                    <button
                      onClick={() => setExpandedId(expandedId === u.id ? null : u.id)}
                      className="text-xs px-3 py-1.5 rounded-md border"
                      style={{ borderColor: "#E1E5DE", color: "#1F5C6B" }}
                    >
                      {expandedId === u.id ? "Hide modules" : "Manage modules"}
                      {u.disabledModules?.length > 0 ? ` (${u.disabledModules.length} hidden)` : ""}
                    </button>
                  )}
                  {!u.allowed && (
                    <button onClick={() => setAllowed(u.id, true)} className="text-xs px-3 py-1.5 rounded-md text-white" style={{ background: "#2F6F53" }}>
                      Approve
                    </button>
                  )}
                  {u.allowed && (
                    <button onClick={() => setAllowed(u.id, false)} className="text-xs px-3 py-1.5 rounded-md border" style={{ borderColor: "#B5443A", color: "#B5443A" }}>
                      Revoke
                    </button>
                  )}
                </div>
              </div>
              {expandedId === u.id && u.allowed && (
                <div className="mt-3 pt-3 border-t grid grid-cols-2 sm:grid-cols-3 gap-2" style={{ borderColor: "#E1E5DE" }}>
                  {MODULES.map((m) => {
                    const hidden = (u.disabledModules || []).includes(m.key);
                    return (
                      <label key={m.key} className="flex items-center gap-2 text-xs" style={{ color: "#1C2624" }}>
                        <input type="checkbox" checked={!hidden} onChange={() => toggleModule(u.id, m.key)} />
                        {m.label}
                      </label>
                    );
                  })}
                </div>
              )}
            </div>
          ))}
        </div>
      )}
    </div>
  );
}
