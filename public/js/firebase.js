// Firebase Auth only (Google sign-in). All data goes through the Apps Script (js/api.js).
import { initializeApp } from 'https://www.gstatic.com/firebasejs/11.0.2/firebase-app.js';
import {
  getAuth, GoogleAuthProvider, signInWithPopup, signOut, onAuthStateChanged,
} from 'https://www.gstatic.com/firebasejs/11.0.2/firebase-auth.js';
import { firebaseConfig } from './config.js';

const auth = getAuth(initializeApp(firebaseConfig));

// Always show Google's account chooser, so people can pick another account after signing out.
export const signIn = () => { const p = new GoogleAuthProvider(); p.setCustomParameters({ prompt: 'select_account' }); return signInWithPopup(auth, p); };
export const logOut = () => signOut(auth);
export const onUser = (cb) => onAuthStateChanged(auth, cb);
export const idToken = () => auth.currentUser?.getIdToken();
