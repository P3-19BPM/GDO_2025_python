import { initializeApp } from "https://www.gstatic.com/firebasejs/10.12.2/firebase-app.js";
import { getAuth, GoogleAuthProvider, signInWithPopup, signOut, onAuthStateChanged } from "https://www.gstatic.com/firebasejs/10.12.2/firebase-auth.js";
import { getFirestore, doc, getDoc, setDoc, onSnapshot, collection, updateDoc, serverTimestamp } from "https://www.gstatic.com/firebasejs/10.12.2/firebase-firestore.js";

// --- CONFIGURAÇÃO E ESTADO GLOBAL ---
const firebaseConfig = {
    apiKey: "AIzaSyCQeD0V91NbUjuhg7-6BRwNOIjw_mdkL24",
    authDomain: "vtcrime-d6a61.firebaseapp.com",
    projectId: "vtcrime-d6a61",
    storageBucket: "vtcrime-d6a61.appspot.com",
    messagingSenderId: "261046458856",
    appId: "1:261046458856:web:3226ca9b456619c85cac5d"
};

const adminEmails = ['novaisufvjm@gmail.com', '19bpm.3@gmail.com'];
const preApprovedViewerEmails = [ /* ... sua longa lista de e-mails ... */ ];

let state = { allCrimeData: [], allVisitaData: [], visitCounts: {}, filteredCrimeData: [], geojsonData: null, map: null, crimeMarkersLayer: null, sectorLayer: null, activeChoropleth: 'SUB_SETOR' };
const palavrasExcluidas = ['autor', 'suspeito', 'co-autor'];

// --- INICIALIZAÇÃO FIREBASE ---
const app = initializeApp(firebaseConfig);
const auth = getAuth(app);
const db = getFirestore(app);
const provider = new GoogleAuthProvider();

// --- FUNÇÕES GLOBAIS (acessíveis pelo HTML) ---
window.authFunctions = { signInWithPopup: () => signInWithPopup(auth, provider), signOut: () => signOut(auth) };
window.adminFunctions = { updateUserStatus: async (uid, newStatus) => { const action = newStatus === 'approved' ? 'aprovar' : 'revogar'; if (!confirm(`Tem certeza que deseja ${action} este usuário?`)) return; try { await updateDoc(doc(db, "user_permissions", uid), { status: newStatus }); alert('Permissão atualizada!'); } catch (error) { console.error("Erro ao atualizar status:", error); alert('Falha ao atualizar status.'); } } };

// --- LÓGICA DE RENDERIZAÇÃO DA UI ---

/**
 * Carrega dinamicamente um arquivo HTML e o insere em um elemento alvo.
 * @param {string} url - O caminho para o arquivo HTML parcial.
 * @param {HTMLElement} targetElement - O elemento do DOM onde o HTML será inserido.
 */
async function loadHTML(url, targetElement) {
    try {
        const response = await fetch(url);
        if (!response.ok) throw new Error(`Não foi possível carregar ${url}`);
        const text = await response.text();
        targetElement.innerHTML = text;
    } catch (error) {
        console.error("Erro ao carregar componente HTML:", error);
        targetElement.innerHTML = `<p class="text-red-500">Erro ao carregar componente.</p>`;
    }
}

/**
 * Controla qual tela é exibida com base no estado do usuário.
 * @param {string} view - O nome da visão a ser mostrada ('login', 'app', 'limited').
 * @param {object} user - (Opcional) O objeto de usuário do Firebase.
 */
async function renderView(view, user = null) {
    const appRoot = document.getElementById('app-root');
    appRoot.innerHTML = ''; // Limpa a tela

    if (view === 'login') {
        await loadHTML('partials/login.html', appRoot);
    } else if (view === 'app') {
        const appShell = document.createElement('div');
        appRoot.appendChild(appShell);
        await loadHTML('partials/app_shell.html', appShell); // 'app_shell.html' seria um novo parcial com o header, etc.
        
        // Agora, dentro do app shell, carregue os componentes corretos
        const userInfoContainer = document.getElementById('user-info');
        const photoURL = user.photoURL || `https://ui-avatars.com/api/?name=${encodeURIComponent(user.displayName || user.email)}&background=random&color=fff&size=32`;
        userInfoContainer.innerHTML = `<div class="flex items-center">...</div>`; // Código do user info

        // Carrega o resto da UI...
        loadAndDisplayData(true);
    }
    // ... e assim por diante para 'limited', 'pending', etc.
}


// --- LÓGICA PRINCIPAL DA APLICAÇÃO ---
// ... (Todas as outras funções do seu script, como initMap, applyFilters, etc., viriam aqui) ...

// --- PONTO DE ENTRADA ---
onAuthStateChanged(auth, (user) => {
    // A lógica de `onAuthStateChanged` que tínhamos antes, mas adaptada para chamar `renderView`.
});
