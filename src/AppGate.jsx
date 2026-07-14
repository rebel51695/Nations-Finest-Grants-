import { useState, useEffect } from "react";
import { doc, getDoc, setDoc } from "firebase/firestore";
import { db } from "./firebaseConfig";
import LoginScreen, { useAuthUser, signOutUser } from "./LoginScreen.jsx";
import PendingScreen from "./PendingScreen.jsx";
import GrantFlow from "./GrantFlow.jsx";
import { isAdminEmail } from "./adminEmails";

async function checkAccess(email) {
  const id = email.trim().toLowerCase();
  const ref = doc(db, "user_access", id);
  const snap = await getDoc(ref);
  if (!snap.exists()) {
    await setDoc(ref, { email: id, allowed: false, requestedAt: new Date().toISOString() });
    return false;
  }
  return snap.data().allowed === true;
}

export default function AppGate() {
  const user = useAuthUser();
  const [accessState, setAccessState] = useState("checking"); // checking | allowed | pending

  useEffect(() => {
    if (!user) return;
    if (isAdminEmail(user.email)) {
      setAccessState("allowed");
      return;
    }
    setAccessState("checking");
    checkAccess(user.email).then((allowed) => setAccessState(allowed ? "allowed" : "pending"));
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
          const allowed = await checkAccess(user.email);
          setAccessState(allowed ? "allowed" : "pending");
        }}
      />
    );
  }

  return <GrantFlow currentUserEmail={user.email} isAdmin={isAdminEmail(user.email)} onSignOut={signOutUser} />;
}
