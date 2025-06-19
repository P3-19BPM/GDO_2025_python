// Arquivo: uploadData.js
// Versão 4 - Converte o GeoJSON para string para contornar limitações do Firestore.

const admin = require('firebase-admin');
const xlsx = require('xlsx');
const fs = require('fs');
const path = require('path');

// --- CONFIGURAÇÃO ---
const serviceAccountKeyFilename = 'vtcrime-d6a61-firebase-adminsdk-fbsvc-86c34e822c.json';
const excelFilename = 'Interacao_Dados.xlsx';
const geojsonFilename = 'SubSetores_19_13-01-2025.geojson';

// --- Validação Inicial dos Arquivos ---
console.log('--- Verificando Arquivos de Configuração ---');
const serviceAccountPath = path.resolve(__dirname, serviceAccountKeyFilename);
const excelFilePath = path.resolve(__dirname, excelFilename);
const geojsonFilePath = path.resolve(__dirname, geojsonFilename);

if (!fs.existsSync(serviceAccountPath)) {
    console.error(`\nERRO CRÍTICO: Arquivo de chave do Firebase não encontrado: ${serviceAccountPath}`);
    process.exit(1);
}
if (!fs.existsSync(excelFilePath)) {
    console.error(`\nERRO CRÍTICO: Arquivo Excel não encontrado: ${excelFilePath}`);
    process.exit(1);
}
if (!fs.existsSync(geojsonFilePath)) {
    console.error(`\nERRO CRÍTICO: Arquivo GeoJSON não encontrado: ${geojsonFilePath}`);
    process.exit(1);
}
console.log('Todos os arquivos necessários foram encontrados.\n');

// --- INICIALIZAÇÃO DO FIREBASE ---
try {
    const serviceAccount = require(serviceAccountPath);
    admin.initializeApp({
      credential: admin.credential.cert(serviceAccount)
    });

    const db = admin.firestore();
    console.log('--- Conexão com Firebase ---');
    console.log('Firebase Admin inicializado com sucesso.\n');

    // --- FUNÇÃO PRINCIPAL ---
    async function uploadData() {
        try {
            console.log('--- Processamento de Dados ---');
            const workbook = xlsx.readFile(excelFilePath);
            const crimeSheet = workbook.Sheets['VT_CRIME'];
            const visitaSheet = workbook.Sheets['VT_VISITA'];
            
            if (!crimeSheet || !visitaSheet) {
                throw new Error('As abas "VT_CRIME" ou "VT_VISITA" não foram encontradas no arquivo Excel.');
            }

            const crimes = xlsx.utils.sheet_to_json(crimeSheet);
            const visitas = xlsx.utils.sheet_to_json(visitaSheet);
            const geojsonData = JSON.parse(fs.readFileSync(geojsonFilePath, 'utf-8'));
            
            console.log(`Dados processados: ${crimes.length} crimes, ${visitas.length} visitas.`);
            console.log('Iniciando upload para o Firestore...\n');
            
            const appDataRef = db.collection('app_data').doc('main');
            
            // CORREÇÃO: Converte o objeto GeoJSON em uma string antes de enviar.
            await appDataRef.set({
                crimes,
                visitas,
                geojson: JSON.stringify(geojsonData), // AQUI ESTÁ A MUDANÇA
                lastUpdated: admin.firestore.FieldValue.serverTimestamp()
            });

            console.log('----------------------------------------------------');
            console.log('SUCESSO! Dados atualizados no Firestore.');
            console.log('----------------------------------------------------');

        } catch (error) {
            console.error('ERRO DURANTE O PROCESSAMENTO OU UPLOAD:', error.message);
        }
    }

    uploadData();

} catch (initError) {
    console.error('\nERRO CRÍTICO AO INICIALIZAR O FIREBASE:', initError.message);
}
