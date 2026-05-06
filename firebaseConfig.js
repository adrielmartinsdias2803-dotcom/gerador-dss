// =========================================
// CONFIGURAÇÃO DO FIREBASE — GERADOR DSS
// =========================================
// ATENÇÃO: Preencha com as credenciais do seu projeto Firebase
// Console: https://console.firebase.google.com
// Projeto → Configurações → Seus aplicativos → Web → firebaseConfig
// =========================================

const firebaseConfig = {
    apiKey: "AIzaSyA6UWiKfQp5-niHVZUqMvULk1sM5anwm7Q",
    authDomain: "mococa-dss.firebaseapp.com",
    projectId: "mococa-dss",
    storageBucket: "mococa-dss.firebasestorage.app",
    messagingSenderId: "208122016697",
    appId: "1:208122016697:web:f5bb2df6b110335829bc68",
    measurementId: "G-F901LZE4RK"
};

// Inicializa o Firebase (usado por todos os arquivos que importarem este script)
if (!firebase.apps || !firebase.apps.length) {
    firebase.initializeApp(firebaseConfig);
}

// Exporta a instância do Firestore
const db = firebase.firestore();
