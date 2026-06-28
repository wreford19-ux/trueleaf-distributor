# TrueLeaf Distributor Ordering App

React + Vite app. Firebase Auth/Firestore for accounts, orders and stock.
Netlify Function (`netlify/functions/submit-order`) sends order emails via Resend.

## Run locally
```
npm install
npm run dev
```

## Deploy (Netlify, Git-connected)
1. Push this folder to a GitHub repo.
2. In Netlify, open the existing `trueleaf-distributor` site → Project configuration →
   Build & deploy → link this repository (build command `npm run build`, publish `dist`).
3. Add Environment variables (Project configuration → Environment variables):
   - `RESEND_API_KEY` = your Resend API key (re_…)
   - `SECRETS_SCAN_SMART_DETECTION_ENABLED` = false   (Firebase public key false-positive)
   NOW (no verified domain yet) — alerts go to your own inbox only:
   - leave `FROM_EMAIL` unset (defaults to onboarding@resend.dev)
   - `ORDER_NOTIFY_TO` = wreford19@gmail.com
   LATER (once a domain is verified in Resend) — full emails incl. distributor confirmation:
   - `FROM_EMAIL` = TrueLeaf Seed Co. <orders@donedigital.co.za>
   - `ORDER_NOTIFY_TO` = orders@trueleafseeds.co.za,wreford@donedigital.co.za
4. Trigger a deploy. Functions auto-deploy from `netlify/functions`.

## Notes
- Firebase web config in `src/firebase.js` is public by design; security is in Firestore rules.
- Orders save to Firestore immediately; the email function is best-effort on top.
