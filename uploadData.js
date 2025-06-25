// Arquivo: uploadData.js
// Versão 13 - Limpa coleções antigas antes de fazer o upload para evitar duplicatas.

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

/**
 * Função auxiliar para deletar todos os documentos de uma coleção em lotes.
 * @param {admin.firestore.CollectionReference} collectionRef A referência da coleção a ser limpa.
 * @param {number} batchSize O tamanho do lote para exclusão.
 */
async function deleteCollection(collectionRef, batchSize) {
    const query = collectionRef.orderBy('__name__').limit(batchSize);

    return new Promise((resolve, reject) => {
        deleteQueryBatch(query, resolve).catch(reject);
    });
}

async function deleteQueryBatch(query, resolve) {
    const snapshot = await query.get();

    if (snapshot.size === 0) {
        // Quando não há mais documentos, a exclusão está completa.
        return resolve();
    }

    // Deleta os documentos em um lote.
    const batch = db.batch();
    snapshot.docs.forEach((doc) => {
        batch.delete(doc.ref);
    });
    await batch.commit();

    // Repete o processo para o próximo lote.
    process.nextTick(() => {
        deleteQueryBatch(query, resolve);
    });
}


async function commitBatchInChunks(collectionRef, dataArray) {
    let batch = db.batch();
    let count = 0;

    for (const item of dataArray) {
        const docRef = collectionRef.doc();
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
        const dataVersion = Date.now().toString();
        console.log(`--- Gerada nova versão dos dados: ${dataVersion} ---\n`);

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

        console.log('\n--- Etapa 2: Calculando estatísticas de resumo ---');
        const uniqueOccurrences = [...new Set(crimes.map(c => c.numero_ocorrencia ? String(c.numero_ocorrencia).trim() : null).filter(id => id))];
        
        const visitCounts = visitas.reduce((acc, v) => {
            if (v.numero_reds_furto) {
                const id = String(v.numero_reds_furto).trim();
                acc[id] = (acc[id] || 0) + 1;
            }
            return acc;
        }, {});

         // ALTERAÇÃO 1: Adicionada a propriedade 'totalVisitas' ao objeto.
        let summaryStats = { total: uniqueOccurrences.length, semVisita: 0, comUma: 0, comDuasOuMais: 0, totalVisitas: 0 };
        uniqueOccurrences.forEach(id => {
            const count = visitCounts[id] || 0;
            if (count === 0) summaryStats.semVisita++;
            else if (count === 1) summaryStats.comUma++;
            else summaryStats.comDuasOuMais++;
        });

          // ALTERAÇÃO 2: Adicionada a linha para calcular o total de visitas.
        summaryStats.totalVisitas = Object.values(visitCounts).reduce((sum, count) => sum + count, 0);

        console.log('Estatísticas calculadas:', summaryStats);
        
        // =========================================================================
        // NOVA ETAPA: Limpando as coleções antes de enviar os novos dados.
        // =========================================================================
        console.log('\n--- Etapa 3: Limpando dados antigos do Firestore ---');
        const crimesCollectionRef = db.collection('crimes');
        const visitasCollectionRef = db.collection('visitas');

        console.log('-> Deletando dados antigos de "crimes"...');
        await deleteCollection(crimesCollectionRef, 500);
        console.log('-> Coleção "crimes" limpa.');

        console.log('\n-> Deletando dados antigos de "visitas"...');
        await deleteCollection(visitasCollectionRef, 500);
        console.log('-> Coleção "visitas" limpa.');
        
        console.log('\n--- Etapa 4: Enviando novos dados para o Firestore ---');
        
        console.log('-> Enviando dados de CRIMES...');
        await commitBatchInChunks(crimesCollectionRef, crimes);

        console.log('\n-> Enviando dados de VISITAS...');
        await commitBatchInChunks(visitasCollectionRef, visitas);

        console.log('\n-> Atualizando metadados com a nova versão...');
        const mainDataRef = db.collection('app_data').doc('main');
        await mainDataRef.set({
            geojsonUrl: geojsonPublicUrl,
            lastUpdated: admin.firestore.FieldValue.serverTimestamp(),
            dataVersion: dataVersion
        });
        
        const summaryDataRef = db.collection('app_data').doc('summary');
        await summaryDataRef.set(summaryStats);
        
        console.log('----------------------------------------------------');
        console.log('SUCESSO! Dados atualizados no Firestore.');
        console.log('----------------------------------------------------');

    } catch (error) {
        console.error('ERRO DURANTE O PROCESSO:', error.message);
    }
}

uploadData();