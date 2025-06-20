// Arquivo: uploadData.js
// Versão 8 - Pré-calcula as estatísticas e as salva em um documento separado.

const admin = require('firebase-admin');
const xlsx = require('xlsx');
const fs = require('fs');
const path = require('path');

// --- CONFIGURAÇÃO ---
const serviceAccountKeyFilename = 'vtcrime-d6a61-firebase-adminsdk-fbsvc-86c34e822c.json';
const excelFilename = 'Interacao_Dados.xlsx';
const geojsonPublicUrl = 'https://raw.githubusercontent.com/P3-19BPM/GDO_2025_python/main/SubSetores_19_13-01-2025.geojson';
const serviceAccount = require(path.resolve(__dirname, serviceAccountKeyFilename));

admin.initializeApp({ credential: admin.credential.cert(serviceAccount) });
const db = admin.firestore();
console.log('Firebase Admin inicializado com sucesso.\n');

async function uploadData() {
    try {
        console.log('--- Etapa 1: Processando dados do Excel ---');
        const excelFilePath = path.resolve(__dirname, excelFilename);
        const workbook = xlsx.readFile(excelFilePath);
        const crimeSheet = workbook.Sheets['VT_CRIME'];
        const visitaSheet = workbook.Sheets['VT_VISITA'];
        if (!crimeSheet || !visitaSheet) throw new Error('Abas "VT_CRIME" ou "VT_VISITA" não encontradas.');

        const crimes = xlsx.utils.sheet_to_json(crimeSheet);
        const visitas = xlsx.utils.sheet_to_json(visitaSheet);

        // --- Etapa 2: Pré-cálculo das estatísticas ---
        console.log('--- Etapa 2: Calculando estatísticas de resumo ---');
        const uniqueOccurrences = [...new Set(crimes.map(c => String(c.numero_ocorrencia)))];
        const visitCounts = visitas.reduce((acc, v) => {
            const id = String(v.numero_reds_furto);
            if (id) acc[id] = (acc[id] || 0) + 1;
            return acc;
        }, {});
        
        let summaryStats = { total: uniqueOccurrences.length, semVisita: 0, comUma: 0, comDuasOuMais: 0 };
        uniqueOccurrences.forEach(id => {
            const count = visitCounts[id] || 0;
            if (count === 0) summaryStats.semVisita++;
            else if (count === 1) summaryStats.comUma++;
            else summaryStats.comDuasOuMais++;
        });
        console.log('Estatísticas calculadas:', summaryStats);
        
        // --- Etapa 3: Salvando dados no Firestore ---
        console.log('\n--- Etapa 3: Enviando dados para o Firestore ---');
        const batch = db.batch();

        // Salva os dados completos
        const mainDataRef = db.collection('app_data').doc('main');
        batch.set(mainDataRef, {
            crimes,
            visitas,
            geojsonUrl: geojsonPublicUrl,
            lastUpdated: admin.firestore.FieldValue.serverTimestamp()
        });

        // Salva os dados de resumo
        const summaryDataRef = db.collection('app_data').doc('summary');
        batch.set(summaryDataRef, summaryStats);

        await batch.commit();

        console.log('----------------------------------------------------');
        console.log('SUCESSO! Dados completos e resumo atualizados no Firestore.');
        console.log('----------------------------------------------------');

    } catch (error) {
        console.error('ERRO DURANTE O PROCESSO:', error.message);
    }
}

uploadData();
