import { useState, useEffect } from "react";
import {
  onAuthStateChanged,
  signInWithEmailAndPassword,
  createUserWithEmailAndPassword,
  sendPasswordResetEmail,
  signOut,
} from "firebase/auth";
import { auth, ALLOWED_EMAIL_DOMAIN } from "./firebaseConfig";

function isAllowedEmail(email) {
  return email.trim().toLowerCase().endsWith("@" + ALLOWED_EMAIL_DOMAIN);
}

function LoginScreen() {
  const [mode, setMode] = useState("signin"); // signin | signup | reset
  const [email, setEmail] = useState("");
  const [password, setPassword] = useState("");
  const [error, setError] = useState("");
  const [info, setInfo] = useState("");
  const [busy, setBusy] = useState(false);

  const friendlyError = (code) => {
    if (code === "auth/invalid-email") return "That doesn't look like a valid email address.";
    if (code === "auth/user-not-found" || code === "auth/invalid-credential") return "No account found with that email and password.";
    if (code === "auth/wrong-password") return "Incorrect password.";
    if (code === "auth/email-already-in-use") return "An account with that email already exists — try signing in instead.";
    if (code === "auth/weak-password") return "Password should be at least 6 characters.";
    if (code === "auth/too-many-requests") return "Too many attempts — please wait a moment and try again.";
    return "Something went wrong. Please try again.";
  };

  const handleSubmit = async (e) => {
    e.preventDefault();
    setError("");
    setInfo("");
    if (!isAllowedEmail(email)) {
      setError(`Only @${ALLOWED_EMAIL_DOMAIN} email addresses can access this portal.`);
      return;
    }
    setBusy(true);
    try {
      if (mode === "signup") {
        await createUserWithEmailAndPassword(auth, email.trim(), password);
      } else if (mode === "reset") {
        await sendPasswordResetEmail(auth, email.trim());
        setInfo("Password reset email sent — check your inbox.");
        setBusy(false);
        return;
      } else {
        await signInWithEmailAndPassword(auth, email.trim(), password);
      }
    } catch (err) {
      setError(friendlyError(err.code));
    }
    setBusy(false);
  };

  return (
    <div className="min-h-screen flex items-center justify-center" style={{ background: "#F6F7F3" }}>
      <style>{`
        @import url('https://fonts.googleapis.com/css2?family=Oswald:wght@600;700&family=Inter:wght@400;500;600&display=swap');
        .login-title { font-family: 'Oswald', sans-serif; text-transform: uppercase; letter-spacing: 0.02em; }
        body, input, button { font-family: 'Inter', system-ui, sans-serif; }
      `}</style>
      <div className="w-full max-w-sm mx-4">
        <div className="flex flex-col items-center mb-6">
          <svg width="44" height="44" viewBox="0 0 40 40" className="mb-2">
            <path d="M6 4 H34 V24 L20 36 L6 24 Z" fill="none" stroke="#F0B21E" strokeWidth="2" strokeLinejoin="round" />
            <path d="M11 11 L12.3 13.8 L15.3 14.2 L13.1 16.3 L13.7 19.3 L11 17.8 L8.3 19.3 L8.9 16.3 L6.7 14.2 L9.7 13.8 Z" fill="#F0B21E" />
            <rect x="17" y="10" width="14" height="2.2" fill="#F0B21E" />
            <rect x="17" y="15" width="14" height="2.2" fill="#F0B21E" />
            <rect x="8" y="20" width="23" height="2.2" fill="#F0B21E" />
          </svg>
          <div className="login-title text-lg" style={{ color: "#17313A" }}>Nation's Finest</div>
          <div className="text-xs" style={{ color: "#B8860B", letterSpacing: "0.05em" }}>GRANT PORTAL</div>
        </div>

        <div className="bg-white rounded-xl border shadow-sm p-6" style={{ borderColor: "#E1E5DE" }}>
          <h1 className="login-title text-base mb-4" style={{ color: "#1C2624" }}>
            {mode === "signup" ? "Create your account" : mode === "reset" ? "Reset password" : "Sign in"}
          </h1>

          <form onSubmit={handleSubmit} className="space-y-3">
            <div>
              <label className="block text-xs font-medium mb-1" style={{ color: "#5B6B66" }}>Email</label>
              <input
                type="email"
                value={email}
                onChange={(e) => setEmail(e.target.value)}
                placeholder={`you@${ALLOWED_EMAIL_DOMAIN}`}
                required
                className="w-full rounded-md border px-3 py-2 text-sm outline-none"
                style={{ borderColor: "#E1E5DE", color: "#1C2624" }}
              />
            </div>
            {mode !== "reset" && (
              <div>
                <label className="block text-xs font-medium mb-1" style={{ color: "#5B6B66" }}>Password</label>
                <input
                  type="password"
                  value={password}
                  onChange={(e) => setPassword(e.target.value)}
                  placeholder={mode === "signup" ? "Choose a password (6+ characters)" : "Your password"}
                  required
                  minLength={6}
                  className="w-full rounded-md border px-3 py-2 text-sm outline-none"
                  style={{ borderColor: "#E1E5DE", color: "#1C2624" }}
                />
              </div>
            )}

            {error && <p className="text-xs" style={{ color: "#B5443A" }}>{error}</p>}
            {info && <p className="text-xs" style={{ color: "#2F6F53" }}>{info}</p>}

            <button
              type="submit"
              disabled={busy}
              className="w-full py-2 rounded-md text-sm text-white font-medium"
              style={{ background: "#17313A", opacity: busy ? 0.6 : 1 }}
            >
              {busy ? "Please wait…" : mode === "signup" ? "Create account" : mode === "reset" ? "Send reset email" : "Sign in"}
            </button>
          </form>

          <div className="mt-4 text-center text-xs space-y-1" style={{ color: "#5B6B66" }}>
            {mode === "signin" && (
              <>
                <div>
                  New here?{" "}
                  <button onClick={() => { setMode("signup"); setError(""); setInfo(""); }} className="underline" style={{ color: "#17313A" }}>
                    Create an account
                  </button>
                </div>
                <div>
                  <button onClick={() => { setMode("reset"); setError(""); setInfo(""); }} className="underline" style={{ color: "#5B6B66" }}>
                    Forgot your password?
                  </button>
                </div>
              </>
            )}
            {mode !== "signin" && (
              <button onClick={() => { setMode("signin"); setError(""); setInfo(""); }} className="underline" style={{ color: "#17313A" }}>
                Back to sign in
              </button>
            )}
          </div>
        </div>

        <p className="text-center text-xs mt-4" style={{ color: "#8A8F87" }}>
          Access is limited to @{ALLOWED_EMAIL_DOMAIN} email addresses.
        </p>
      </div>
    </div>
  );
}

export function useAuthUser() {
  const [user, setUser] = useState(undefined); // undefined = still checking, null = signed out
  useEffect(() => {
    const unsub = onAuthStateChanged(auth, (u) => setUser(u));
    return unsub;
  }, []);
  return user;
}

export function signOutUser() {
  return signOut(auth);
}

export default LoginScreen;
