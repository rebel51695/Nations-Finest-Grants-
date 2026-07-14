import { useState } from "react";

export default function PendingScreen({ email, onSignOut, onCheckAgain, denied }) {
  const [checking, setChecking] = useState(false);

  const handleCheck = async () => {
    setChecking(true);
    await onCheckAgain();
    setChecking(false);
  };

  return (
    <div className="min-h-screen flex items-center justify-center" style={{ background: "#F6F7F3" }}>
      <style>{`
        @import url('https://fonts.googleapis.com/css2?family=Oswald:wght@600;700&family=Inter:wght@400;500;600&display=swap');
        .pending-title { font-family: 'Oswald', sans-serif; text-transform: uppercase; letter-spacing: 0.02em; }
        body, button { font-family: 'Inter', system-ui, sans-serif; }
      `}</style>
      <div className="w-full max-w-sm mx-4 bg-white rounded-xl border shadow-sm p-6 text-center" style={{ borderColor: "#E1E5DE" }}>
        <svg width="40" height="40" viewBox="0 0 40 40" className="mx-auto mb-3">
          <path d="M6 4 H34 V24 L20 36 L6 24 Z" fill="none" stroke="#F0B21E" strokeWidth="2" strokeLinejoin="round" />
          <path d="M11 11 L12.3 13.8 L15.3 14.2 L13.1 16.3 L13.7 19.3 L11 17.8 L8.3 19.3 L8.9 16.3 L6.7 14.2 L9.7 13.8 Z" fill="#F0B21E" />
          <rect x="17" y="10" width="14" height="2.2" fill="#F0B21E" />
          <rect x="17" y="15" width="14" height="2.2" fill="#F0B21E" />
          <rect x="8" y="20" width="23" height="2.2" fill="#F0B21E" />
        </svg>
        <h1 className="pending-title text-base mb-2" style={{ color: "#1C2624" }}>
          {denied ? "Access denied" : "Waiting for approval"}
        </h1>
        <p className="text-sm mb-1" style={{ color: "#5B6B66" }}>
          Signed in as <strong>{email}</strong>
        </p>
        <p className="text-sm mb-5" style={{ color: "#5B6B66" }}>
          {denied
            ? "An administrator has not granted this account access to the Grant Portal."
            : "Your account has been created, but an administrator needs to approve it before you can access the Grant Portal."}
        </p>
        <div className="flex flex-col gap-2">
          {!denied && (
            <button
              onClick={handleCheck}
              disabled={checking}
              className="py-2 rounded-md text-sm text-white font-medium"
              style={{ background: "#17313A", opacity: checking ? 0.6 : 1 }}
            >
              {checking ? "Checking…" : "Check again"}
            </button>
          )}
          <button onClick={onSignOut} className="py-2 rounded-md text-sm border" style={{ borderColor: "#E1E5DE", color: "#1C2624" }}>
            Sign out
          </button>
        </div>
      </div>
    </div>
  );
}
