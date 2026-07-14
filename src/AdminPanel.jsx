import { useState, useEffect } from "react";
import { collection, getDocs, doc, setDoc } from "firebase/firestore";
import { db } from "./firebaseConfig";
import { ADMIN_EMAILS } from "./adminEmails";

export default function AdminPanel() {
  const [users, setUsers] = useState([]);
  const [loading, setLoading] = useState(true);
  const [error, setError] = useState("");

  const load = async () => {
    setLoading(true);
    setError("");
    try {
      const snap = await getDocs(collection(db, "user_access"));
      const list = snap.docs.map((d) => ({ id: d.id, ...d.data() }));
      list.sort((a, b) => (b.requestedAt || "").localeCompare(a.requestedAt || ""));
      setUsers(list);
    } catch (e) {
      setError("Couldn't load the user list.");
    }
    setLoading(false);
  };

  useEffect(() => { load(); }, []);

  const setAllowed = async (id, allowed) => {
    await setDoc(doc(db, "user_access", id), { allowed }, { merge: true });
    setUsers((prev) => prev.map((u) => (u.id === id ? { ...u, allowed } : u)));
  };

  const fmtDate = (iso) => (iso ? new Date(iso).toLocaleString("en-US", { month: "short", day: "numeric", year: "numeric", hour: "numeric", minute: "2-digit" }) : "—");

  return (
    <div className="space-y-4">
      <div>
        <h1 className="font-display text-2xl" style={{ color: "#1C2624" }}>User Access</h1>
        <p className="text-sm mt-1" style={{ color: "#5B6B66" }}>Approve or deny who can sign in to the Grant Portal</p>
      </div>

      <div className="bg-white rounded-lg border p-4 text-sm" style={{ borderColor: "#E1E5DE" }}>
        <div className="font-medium mb-1" style={{ color: "#1C2624" }}>Admins (always have full access)</div>
        <div style={{ color: "#5B6B66" }}>{ADMIN_EMAILS.join(", ")}</div>
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
            <div key={u.id} className="px-4 py-3 flex items-center justify-between text-sm">
              <div>
                <div style={{ color: "#1C2624" }}>{u.email || u.id}</div>
                <div className="text-xs" style={{ color: "#8A8F87" }}>Requested {fmtDate(u.requestedAt)}</div>
              </div>
              <div className="flex items-center gap-2">
                <span
                  className="text-xs px-2 py-0.5 rounded-full"
                  style={{
                    background: u.allowed ? "#EAF1EC" : "#FBEAE8",
                    color: u.allowed ? "#2F6F53" : "#B5443A",
                  }}
                >
                  {u.allowed ? "Approved" : "Pending / Denied"}
                </span>
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
          ))}
        </div>
      )}
    </div>
  );
}
