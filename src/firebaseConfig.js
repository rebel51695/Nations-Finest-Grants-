import { initializeApp } from "firebase/app";
import { getFirestore } from "firebase/firestore";

const firebaseConfig = {
  apiKey: "AIzaSyAsXf-nsgHyGxACNm1DpDI8A0-zg4q1MUQ",
  authDomain: "nation-s-finest-grants.firebaseapp.com",
  projectId: "nation-s-finest-grants",
  storageBucket: "nation-s-finest-grants.firebasestorage.app",
  messagingSenderId: "611075575486",
  appId: "1:611075575486:web:c7a96f4314a5b4c43c98ef",
};

const app = initializeApp(firebaseConfig);
export const db = getFirestore(app);
