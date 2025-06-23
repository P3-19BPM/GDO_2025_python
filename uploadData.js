// Arquivo: uploadData.js
// Versão 11 - Adiciona validação de cabeçalhos para garantir a leitura correta dos dados.

const admin = require('firebase-admin');
const xlsx = require('xlsx');
const path = require('path');

// --- CONFIGURAÇÃO ---
const serviceAccountKeyFilename = 'vtcrime-d6a61-firebase-adminsdk-fbsvc-86c34e822c.json';
const excelFilename = 'Interacao_Dados.xlsx';
const geojsonPublicUrl = 'https://raw.githubusercontent.com/P3-19BPM/GDO_2025_python/main/SubSetores_19_13-01-2025.geojson';
const serviceAccount = require(path.resolve(__dirname, serviceAccountKeyFilename));

admin.initializeApp({ credential: admin.credential.cert(serviceAccount) });
const db = admin.firestore();
console.log('Firebase Admin inicializado com sucesso.\n');

// Função para processar os envios em lotes de 500 operações
async function commitBatchInChunks(collectionRef, dataArray) {
    let batch = db.batch();
    let count = 0;

    for (const item of dataArray) {
        const docRef = collectionRef.doc(); // Cria um novo documento com ID automático
        batch.set(docRef, item);
        count++;

        if (count === 500) {
            console.log(`  - Enviando lote de 500 documentos para a coleção ${collectionRef.id}...`);
            await batch.commit();
            batch = db.batch();
            count = 0;
        }
    }

    if (count > 0) {
        console.log(`  - Enviando lote final de ${count} documentos para a coleção ${collectionRef.id}...`);
        await batch.commit();
    }
}


async function uploadData() {
    try {
        console.log('--- Etapa 1: Processando dados do Excel ---');
        const excelFilePath = path.resolve(__dirname, excelFilename);
        if (!require('fs').existsSync(excelFilePath)) {
            throw new Error(`Arquivo Excel "${excelFilename}" não encontrado no diretório.`);
        }
        
        const workbook = xlsx.readFile(excelFilePath);
        const crimeSheet = workbook.Sheets['VT_CRIME'];
        const visitaSheet = workbook.Sheets['VT_VISITA'];
        if (!crimeSheet || !visitaSheet) throw new Error('Abas "VT_CRIME" ou "VT_VISITA" não encontradas no arquivo Excel.');

        const crimes = xlsx.utils.sheet_to_json(crimeSheet);
        const visitas = xlsx.utils.sheet_to_json(visitaSheet);

        console.log(`  - Registros lidos da aba "VT_CRIME": ${crimes.length}`);
        console.log(`  - Registros lidos da aba "VT_VISITA": ${visitas.length}`);

        // Validação crucial para garantir que os dados foram lidos corretamente
        if (crimes.length === 0 || !crimes[0].hasOwnProperty('numero_ocorrencia')) {
            throw new Error('Falha ao ler dados de "VT_CRIME". Verifique se a primeira linha da planilha contém os cabeçalhos corretos (ex: "numero_ocorrencia") e se não há linhas em branco no topo.');
        }
        if (visitas.length > 0 && !visitas[0].hasOwnProperty('numero_reds_furto')) {
             console.warn('AVISO: A aba "VT_VISITA" parece não ter o cabeçalho "numero_reds_furto". Os cálculos de visitas podem falhar.');
        }


        console.log('\n--- Etapa 2: Calculando estatísticas de resumo ---');
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
        
        console.log('\n--- Etapa 3: Enviando dados para o Firestore ---');
        
        // Envia os crimes para a coleção 'crimes'
        console.log('-> Enviando dados de CRIMES...');
        const crimesCollectionRef = db.collection('crimes');
        await commitBatchInChunks(crimesCollectionRef, crimes);

        // Envia as visitas para a coleção 'visitas'
        console.log('\n-> Enviando dados de VISITAS...');
        const visitasCollectionRef = db.collection('visitas');
        await commitBatchInChunks(visitasCollectionRef, visitas);

        // O documento 'main' agora guarda apenas metadados e a URL do GeoJSON.
        console.log('\n-> Atualizando metadados...');
        const mainDataRef = db.collection('app_data').doc('main');
        await mainDataRef.set({
            geojsonUrl: geojsonPublicUrl,
            lastUpdated: admin.firestore.FieldValue.serverTimestamp()
        });
        
        // Salva os dados de resumo
        const summaryDataRef = db.collection('app_data').doc('summary');
        await summaryDataRef.set(summaryStats);
        
        console.log('----------------------------------------------------');
        console.log('SUCESSO! Dados atualizados no Firestore.');
        console.log('----------------------------------------------------');

    } catch (error) {
        // Agora mostra a mensagem de erro completa para melhor diagnóstico
        console.error('ERRO DURANTE O PROCESSO:', error.message);
    }
}

uploadData();
