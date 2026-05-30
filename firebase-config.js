/**
 * ⚙️ MCHMUK Firebase Configuration
 * กรอกค่าจาก Firebase Console → Project Settings → Your Apps → Web App
 * https://console.firebase.google.com
 */
const firebaseConfig = {
    apiKey:            "YOUR_API_KEY",
    authDomain:        "YOUR_PROJECT_ID.firebaseapp.com",
    projectId:         "YOUR_PROJECT_ID",
    storageBucket:     "YOUR_PROJECT_ID.appspot.com",
    messagingSenderId: "YOUR_SENDER_ID",
    appId:             "YOUR_APP_ID"
};

// Initialize Firebase
firebase.initializeApp(firebaseConfig);

// Global service references (used throughout app)
const fbAuth    = firebase.auth();
const fbDb      = firebase.firestore();
const fbStorage = firebase.storage();
