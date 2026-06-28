// firebase.js — TrueLeaf Distributor Ordering App
// Config is filled in and ready. Just place this in your project's src/ folder.

import { initializeApp } from "firebase/app";
import { getAuth } from "firebase/auth";
import { getFirestore } from "firebase/firestore";

const firebaseConfig = {
  apiKey: "AIzaSyDtSRcyUggDQl2ge_kbDXTPIACCHDfZiRU",
  authDomain: "trueleaf-distributors.firebaseapp.com",
  projectId: "trueleaf-distributors",
  storageBucket: "trueleaf-distributors.firebasestorage.app",
  messagingSenderId: "173196707811",
  appId: "1:173196707811:web:cf8b438061a85b513942d6",
};

const app = initializeApp(firebaseConfig);
export const auth = getAuth(app);
export const db = getFirestore(app);
