import { useState, useEffect } from "react";
import { doc, getDoc, setDoc } from "firebase/firestore";
import { db } from "./firebaseConfig";
import LoginScreen, { useAuthUser, signOutUser } from "./LoginScreen.jsx";
import PendingScreen from "./PendingScreen.jsx";
import GrantFlow from "./GrantFlow.jsx";
import { isAdminEmail } from "./adminEmails";

async function fetchAccessRecord(email) {
  const id = email.trim().toLowerCase();
  const ref = doc(db, "user_access", id);
  const snap = await getDoc(ref);
  if (!snap.exists()) {
    const fresh = { email: id, allowed: false, role: "editor", disabledModules: [], requestedAt: new Date().toISOString() };
    await setDoc(ref, fresh);
    return fresh;
  }
  const data = snap.data();
  return {
    allowed: data.allowed === true,
    role: data.role || "editor",
    disabledModules: Array.isArray(data.disabledModules) ? data.disabledModules : [],
  };
}

export default function AppGate() {
  const user = useAuthUser();
  const [accessState, setAccessState] = useState("checking"); // checking | allowed | pending
  const [record, setRecord] = useState(null);

  useEffect(() => {
    if (!user) return;
    if (isAdminEmail(user.email)) {
      setRecord({ allowed: true, role: "admin", disabledModules: [] });
      setAccessState("allowed");
      return;
    }
    setAccessState("checking");
    fetchAccessRecord(user.email).then((rec) => {
      setRecord(rec);
      setAccessState(rec.allowed ? "allowed" : "pending");
    });
  }, [user]);

  if (user === undefined) {
    return (
      <div className="min-h-screen flex items-center justify-center" style={{ background: "#F6F7F3", color: "#8A8F87" }}>
        Loading…
      </div>
    );
  }

  if (!user) {
    return <LoginScreen />;
  }

  if (accessState === "checking") {
    return (
      <div className="min-h-screen flex items-center justify-center" style={{ background: "#F6F7F3", color: "#8A8F87" }}>
        Loading…
      </div>
    );
  }

  if (accessState === "pending") {
    return (
      <PendingScreen
        email={user.email}
        onSignOut={signOutUser}
        onCheckAgain={async () => {
          const rec = await fetchAccessRecord(user.email);
          setRecord(rec);
          setAccessState(rec.allowed ? "allowed" : "pending");
        }}
      />
    );
  }

  return (
    <GrantFlow
      currentUserEmail={user.email}
      isAdmin={isAdminEmail(user.email)}
      userRole={isAdminEmail(user.email) ? "admin" : (record?.role || "editor")}
      disabledModules={isAdminEmail(user.email) ? [] : (record?.disabledModules || [])}
      onSignOut={signOutUser}
    />
  );
}
