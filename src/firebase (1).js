// firebase.js — TrueLeaf Distributor Ordering App
import { initializeApp, deleteApp } from "firebase/app";
import { getAuth, createUserWithEmailAndPassword, signOut } from "firebase/auth";
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

// Create a distributor auth account WITHOUT disturbing the admin's current
// session, using a throwaway secondary Firebase app instance. Returns the new uid.
export async function createDistributorAuth(email, password) {
  const secondary = initializeApp(firebaseConfig, "secondary-" + Date.now());
  try {
    const sAuth = getAuth(secondary);
    const cred = await createUserWithEmailAndPassword(sAuth, email, password);
    const uid = cred.user.uid;
    await signOut(sAuth);
    return uid;
  } finally {
    await deleteApp(secondary);
  }
}
