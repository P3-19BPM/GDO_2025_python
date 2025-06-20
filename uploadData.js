// Arquivo: uploadData.js
// Versão 5 - Faz upload do GeoJSON para o Firebase Storage e salva o link no Firestore.

const admin = require('firebase-admin');
const xlsx = require('xlsx');
const fs = require('fs');
const path = require('path');
const { Storage } = require('@google-cloud/storage'); // <-- Ferramenta para o Storage

// --- CONFIGURAÇÃO ---
const serviceAccountKeyFilename = 'vtcrime-d6a61-firebase-adminsdk-fbsvc-86c34e822c.json';
const excelFilename = 'Interacao_Dados.xlsx';
const geojsonFilename = 'SubSetores_19_13-01-2025.geojson';
const serviceAccount = require(path.resolve(__dirname, serviceAccountKeyFilename));

// --- INICIALIZAÇÃO DO FIREBASE E STORAGE ---
// A configuração agora inclui o 'storageBucket'
admin.initializeApp({
  credential: admin.credential.cert(serviceAccount),
  storageBucket: `${serviceAccount.project_id}.appspot.com` 
});

const db = admin.firestore();
const bucket = admin.storage().bucket(); // Acessa o bucket de armazenamento padrão

console.log('Firebase Admin e Storage inicializados com sucesso.\n');

// --- FUNÇÃO PRINCIPAL ---
async function uploadData() {
    try {
        // --- 1. Upload do GeoJSON para o Storage ---
        console.log('--- Etapa 1: Fazendo upload do GeoJSON para o Storage ---');
        const geojsonFilePath = path.resolve(__dirname, geojsonFilename);
        const destinationPath = `mapas/${geojsonFilename}`; // Salva em uma pasta 'mapas'

        await bucket.upload(geojsonFilePath, {
            destination: destinationPath,
            public: true // Torna o arquivo publicamente legível
        });

        // Gera a URL pública para o arquivo
        const publicUrl = `https://storage.googleapis.com/${bucket.name}/${destinationPath}`;
        console.log(`GeoJSON enviado com sucesso. URL pública: ${publicUrl}\n`);

        // --- 2. Processamento dos Dados do Excel ---
        console.log('--- Etapa 2: Processando dados do Excel ---');
        const excelFilePath = path.resolve(__dirname, excelFilename);
        const workbook = xlsx.readFile(excelFilePath);
        const crimeSheet = workbook.Sheets['VT_CRIME'];
        const visitaSheet = workbook.Sheets['VT_VISITA'];
        
        if (!crimeSheet || !visitaSheet) {
            throw new Error('As abas "VT_CRIME" ou "VT_VISITA" não foram encontradas no Excel.');
        }

        const crimes = xlsx.utils.sheet_to_json(crimeSheet);
        const visitas = xlsx.utils.sheet_to_json(visitaSheet);
        console.log(`Dados processados: ${crimes.length} crimes, ${visitas.length} visitas.\n`);

        // --- 3. Salva os dados e o LINK no Firestore ---
        console.log('--- Etapa 3: Salvando dados no Firestore ---');
        const appDataRef = db.collection('app_data').doc('main');
        
        await appDataRef.set({
            crimes,
            visitas,
            geojsonUrl: publicUrl, // Salva APENAS a URL do arquivo GeoJSON
            lastUpdated: admin.firestore.FieldValue.serverTimestamp()
        });

        console.log('----------------------------------------------------');
        console.log('SUCESSO! Dados e link do mapa atualizados no Firestore.');
        console.log('----------------------------------------------------');

    } catch (error) {
        console.error('ERRO DURANTE O PROCESSO:', error.message);
    }
}

uploadData();
