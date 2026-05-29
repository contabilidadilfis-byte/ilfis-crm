import { initializeApp } from 'firebase/app';
import { getFirestore } from 'firebase/firestore';

const firebaseConfig = {
  apiKey: "AIzaSyD4ituHS6aDE74FBKYsFtUT6suu49WWN-E",
  authDomain: "ilfis-crm.firebaseapp.com",
  projectId: "ilfis-crm",
  storageBucket: "ilfis-crm.firebasestorage.app",
  messagingSenderId: "706606962641",
  appId: "1:706606962641:web:3a481e42c9bc40b952b2f0"
};

const app = initializeApp(firebaseConfig);
export const db = getFirestore(app);
