# Nation's Finest Grant Portal (GrantFlow)

A standalone version of GrantFlow, backed by a real Firebase Firestore database
instead of Claude's artifact storage. This means everyone who visits the
deployed URL sees and edits the same shared data — no more separate copies.

## What changed from the Claude artifact version

- `src/storageShim.js` recreates the same `window.storage` API the app already
  used, but backed by Firestore (for shared data) and `localStorage` (for the
  personal "who am I" name). The rest of the app (`GrantFlow.jsx`) is
  unchanged — same features, same code, same UI.
- `src/firebaseConfig.js` holds the Firebase project connection details.

## Running it locally (optional, to test before deploying)

```
npm install
npm run dev
```

Then open the URL it prints (usually http://localhost:5173).

## Deploying for real (so your team has one shared URL)

1. **Push this project to a GitHub repository.**
   - Create a new empty repo on github.com (no README/license, just empty)
   - From this folder, run:
     ```
     git init
     git add .
     git commit -m "Initial GrantFlow standalone app"
     git branch -M main
     git remote add origin <your-repo-URL>
     git push -u origin main
     ```

2. **Deploy on Vercel** (free tier is enough for this):
   - Go to vercel.com, sign in with your GitHub account
   - Click "Add New" → "Project"
   - Select the repository you just pushed
   - Vercel will auto-detect it's a Vite project — no configuration needed
   - Click "Deploy"
   - After a minute or two, you'll get a real URL (e.g. `grantflow.vercel.app`)
     that anyone on your team can open, and everyone will see the same data.

3. **(Optional) Custom domain** — in the Vercel project settings, under
   "Domains," you can point a domain you own (e.g. `grants.nationsfinest.org`)
   at this deployment instead of using the default `.vercel.app` address.

## Firestore security note

The current setup allows anyone with the app's URL to read and write data —
matching how the tool has worked so far (shared, no login). If you ever want
to restrict who can access it, that would mean adding real user accounts
(Firebase Authentication) and tightening the Firestore security rules to
require being signed in. That's a bigger follow-up project, not something
needed to get this working today.
