// Firebase config and Google Auth logic for React
import { initializeApp } from "firebase/app";
import { getAuth, GoogleAuthProvider, signInWithPopup, signOut } from "firebase/auth";

const firebaseConfig = {
  apiKey: "AIzaSyBW3biS7ROg58IH21vvPZ42DbpsaJMCQ-Y",
  authDomain: "foreclosure-15f09.firebaseapp.com",
  projectId: "foreclosure-15f09",
  storageBucket: "foreclosure-15f09.firebasestorage.app",
  messagingSenderId: "719554951830",
  appId: "1:719554951830:web:ac5f63a889ea95c95460b7",
  measurementId: "G-0F2LF33GM3"
};

const app = initializeApp(firebaseConfig);
export const auth = getAuth(app);
export const provider = new GoogleAuthProvider();

export function signInWithGoogle() {
  return signInWithPopup(auth, provider);
}

export function signOutGoogle() {
  return signOut(auth);
}
