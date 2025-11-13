// server.js
import express from 'express';
import bodyParser from 'body-parser';
import fetch from 'node-fetch';
import dotenv from 'dotenv';
import cors from 'cors';

dotenv.config();

const app = express();

app.use(cors({
    origin: '*',
    methods: ['GET', 'POST', 'PUT', 'DELETE', 'OPTIONS'],
    allowedHeaders: ['Content-Type', 'Authorization', 'X-Requested-With', 'Accept', 'Origin'],
    credentials: true
}));

app.options('*', cors()); 

app.use(bodyParser.json({ limit: '50mb' }));

console.log('🚀 API SharePoint Global Plastic a iniciar...');

const COLUMN_MAPPING = {
    'Title': (row) => row['N° do ticket'] + ' - ' + row.Item + ' - ' + row.Motivo,
    'N_x00b0_doticket': (row) => row['N° do ticket'],
    'NomedoCliente': (row) => row['Nome do Cliente'],
    'Item': (row) => row.Item,
    'Qtde': (row) => String(row.Qtde),
    'Motivo': (row) => row.Motivo,
    'Origemdodefeito': (row) => row['Origem do defeito'],
    'Disposi_x00e7__x00e3_o': (row) => row.Disposição,
    'Disposi_x00e7__x00e3_odaspe_x00e': (row) => row['Disposição das peças'],
    'DatadeGera_x00e7__x00e3_o': (row) => row['Data de Geração'] || '',
    'Foto1': (row) => row['Foto 1'] || null,
    'Foto2': (row) => row['Foto 2'] || null,
    'Foto3': (row) => row['Foto 3'] || null,
    'Foto4': (row) => row['Foto 4'] || null,
    'Foto5': (row) => row['Foto 5'] || null,
    'Foto6': (row) => row['Foto 6'] || null,
    'Foto7': (row) => row['Foto 7'] || null,
    'Foto8': (row) => row['Foto 8'] || null,
    'Foto9': (row) => row['Foto 9'] || null,
    'Foto10': (row) => row['Foto 10'] || null,
};

async function getAccessToken(retries = 3) {
  for (let i = 0; i < retries; i++) {
    try {
      const params = new URLSearchParams();
      params.append('client_id', process.env.CLIENT_ID);
      params.append('scope', 'https://graph.microsoft.com/.default');
      params.append('client_secret', process.env.CLIENT_SECRET);
      params.append('grant_type', 'client_credentials');
      
      const res = await fetch(`https://login.microsoftonline.com/${process.env.TENANT_ID}/oauth2/v2.0/token`, {
        method: 'POST',
        body: params,
        headers: { 'Content-Type': 'application/x-www-form-urlencoded' }
      });
      
      const data = await res.json();
      if (!data.access_token) throw new Error(`Erro na autenticação: ${data.error_description || data.error}`);
      return data.access_token;
    } catch (error) {
      if (i === retries - 1) throw error;
      await new Promise(resolve => setTimeout(resolve, 1000 * (i + 1)));
    }
  }
}

async function getDriveId(accessToken) {
    const url = `https://graph.microsoft.com/v1.0/sites/${process.env.SITE_ID}/drives`;
    const res = await fetch(url, { headers: { 'Authorization': `Bearer ${accessToken}` } });
    if (!res.ok) throw new Error(`Erro ao buscar drives: ${res.status}`);
    const { value: drives } = await res.json();
    const library = drives.find(d => d.name === process.env.LIBRARY_NAME);
    if (!library) throw new Error(`Biblioteca "${process.env.LIBRARY_NAME}" não encontrada.`);
    return library.id;
}

async function getListId(accessToken) {
    const listName = process.env.LIST_NAME;
    const url = `https://graph.microsoft.com/v1.0/sites/${process.env.SITE_ID}/lists?$filter=displayName eq '${encodeURIComponent(listName)}'`;
    const res = await fetch(url, { headers: { 'Authorization': `Bearer ${accessToken}` } });
    if (!res.ok) throw new Error(`Erro ao buscar listas: ${res.status}`);
    const { value: lists } = await res.json();
    if (lists.length > 0) return lists[0].id;
    throw new Error(`Lista "${listName}" não encontrada.`);
}

app.get('/', (req, res) => res.json({ status: 'online', timestamp: new Date().toISOString() }));

// ✅ NOVA ROTA: Verificar Status (Existe na Lista? Existe no PDF?)
app.get('/check-status/:ticketNumber', async (req, res) => {
    const { ticketNumber } = req.params;
    try {
        const accessToken = await getAccessToken();
        const siteId = process.env.SITE_ID;
        const driveId = await getDriveId(accessToken);
        const listId = await getListId(accessToken);

        // 1. Verifica na Lista
        // Filtra pelo campo N_x00b0_doticket
        const listUrl = `https://graph.microsoft.com/v1.0/sites/${siteId}/lists/${listId}/items?expand=fields($select=N_x00b0_doticket)&$filter=fields/N_x00b0_doticket eq '${ticketNumber}'`;
        const listRes = await fetch(listUrl, { headers: { 'Authorization': `Bearer ${accessToken}` } });
        
        let existsInList = false;
        if (listRes.ok) {
             const data = await listRes.json();
             existsInList = data.value && data.value.length > 0;
        }

        // 2. Verifica o PDF
        const encodedFolder = encodeURIComponent(process.env.FOLDER_PATH);
        // Busca arquivos que começam com "Laudo - NUMERO -"
        const pdfNamePart = `Laudo - ${ticketNumber}-`;
        const driveUrl = `https://graph.microsoft.com/v1.0/drives/${driveId}/root:/${encodedFolder}:/search(q='${pdfNamePart}')`;
        const driveRes = await fetch(driveUrl, { headers: { 'Authorization': `Bearer ${accessToken}` } });
        
        let existsInPdf = false;
        if (driveRes.ok) {
            const data = await driveRes.json();
            // Confirmação extra para garantir que não pegou um arquivo parecido
            existsInPdf = data.value && data.value.some(f => f.name.includes(ticketNumber) && f.name.endsWith('.pdf'));
        }

        res.json({ existsInList, existsInPdf });

    } catch (error) {
        console.error(`Erro check-status:`, error.message);
        // Se der erro na verificação, assume false para tentar reenviar (fail-safe) ou retorna erro
        res.status(500).json({ error: error.message });
    }
});

// ROTA 1: Upload PDF
app.post('/upload-pdf', async (req, res) => {
  const { fileName, fileBase64 } = req.body;
  if (!fileName || !fileBase64) return res.status(400).json({ error: 'Dados incompletos' });

  try {
    const accessToken = await getAccessToken();
    const driveId = await getDriveId(accessToken);
    const encodedFolder = encodeURIComponent(process.env.FOLDER_PATH);
    const encodedFileName = encodeURIComponent(fileName);
    const uploadUrl = `https://graph.microsoft.com/v1.0/drives/${driveId}/root:/${encodedFolder}/${encodedFileName}:/content`;
    
    const response = await fetch(uploadUrl, {
      method: 'PUT',
      headers: { 'Authorization': `Bearer ${accessToken}`, 'Content-Type': 'application/pdf' },
      body: Buffer.from(fileBase64, 'base64')
    });

    if (!response.ok) throw new Error(`SharePoint Error ${response.status}`);
    const result = await response.json();
    res.status(200).json({ success: true, sharePointUrl: result.webUrl });
  } catch (error) {
    console.error(`❌ Erro PDF:`, error.message);
    res.status(500).json({ success: false, error: error.message });
  }
});

// ROTA 2: Upload Lista
app.post('/upload-list-data', async (req, res) => {
    const { listData } = req.body;
    if (!listData || listData.length === 0) return res.status(400).json({ success: false, error: 'Sem dados' });

    try {
        const accessToken = await getAccessToken();
        const listId = await getListId(accessToken); 
        const listItemsUrl = `https://graph.microsoft.com/v1.0/sites/${process.env.SITE_ID}/lists/${listId}/items`;

        const insertionPromises = listData.map(async (row) => {
            const itemFields = {};
            for (const key in COLUMN_MAPPING) {
                const val = COLUMN_MAPPING[key](row);
                if (val !== null && val !== '' && val !== undefined) itemFields[key] = val;
            }
            
            const itemResponse = await fetch(listItemsUrl, {
                method: 'POST',
                headers: { 'Authorization': `Bearer ${accessToken}`, 'Content-Type': 'application/json' },
                body: JSON.stringify({ fields: itemFields })
            });

            if (!itemResponse.ok) throw new Error(`Status: ${itemResponse.status}`);
            return itemResponse.json();
        });

        await Promise.all(insertionPromises);
        res.status(200).json({ success: true });
    } catch (error) {
        console.error(`❌ Erro lista:`, error.message);
        res.status(500).json({ success: false, error: error.message });
    }
});

// ROTA 3: Deletar PDF
app.delete('/delete-pdf-by-ticket-number/:ticketNumber', async (req, res) => {
    const { ticketNumber } = req.params;
    if (!ticketNumber) return res.status(400).json({ error: 'Ticket obrigatório' });

    try {
        const accessToken = await getAccessToken();
        const driveId = await getDriveId(accessToken);
        const encodedFolder = encodeURIComponent(process.env.FOLDER_PATH);
        const listUrl = `https://graph.microsoft.com/v1.0/drives/${driveId}/root:/${encodedFolder}:/children`;
        
        const listResponse = await fetch(listUrl, { headers: { 'Authorization': `Bearer ${accessToken}` } });
        if (!listResponse.ok) throw new Error(`Erro listagem`);
        const { value: allFiles } = await listResponse.json();
        
        const filesToDelete = allFiles.filter(file => file.name.startsWith(`Laudo - ${ticketNumber}-`));
        if (filesToDelete.length === 0) return res.json({ success: true, message: 'Nada a excluir.' });

        await Promise.all(filesToDelete.map(file => 
            fetch(`https://graph.microsoft.com/v1.0/drives/${driveId}/items/${file.id}`, { method: 'DELETE', headers: { 'Authorization': `Bearer ${accessToken}` } })
        ));
        res.status(200).json({ success: true });
    } catch (error) {
        res.status(500).json({ success: false, error: error.message });
    }
});

// ROTA 4: Limpar Lista
app.delete('/clear-list', async (req, res) => {
    try {
        const accessToken = await getAccessToken();
        const listId = await getListId(accessToken);
        let itemsToDelete = [];
        let nextLink = `https://graph.microsoft.com/v1.0/sites/${process.env.SITE_ID}/lists/${listId}/items?$select=id`;
        
        while (nextLink) {
            const response = await fetch(nextLink, { headers: { 'Authorization': `Bearer ${accessToken}` } });
            if (!response.ok) throw new Error(`Erro busca`);
            const data = await response.json();
            if (data.value) itemsToDelete = itemsToDelete.concat(data.value);
            nextLink = data['@odata.nextLink'];
        }

        if (itemsToDelete.length === 0) return res.status(200).json({ success: true, message: 'Lista vazia.' });

        const BATCH_SIZE = 10;
        for (let i = 0; i < itemsToDelete.length; i += BATCH_SIZE) {
            const batch = itemsToDelete.slice(i, i + BATCH_SIZE);
            await Promise.all(batch.map(item => 
                fetch(`https://graph.microsoft.com/v1.0/sites/${process.env.SITE_ID}/lists/${listId}/items/${item.id}`, { method: 'DELETE', headers: { 'Authorization': `Bearer ${accessToken}` } })
            ));
        }
        res.status(200).json({ success: true });
    } catch (error) {
        res.status(500).json({ success: false, error: error.message });
    }
});

const PORT = process.env.PORT || 3000;
app.listen(PORT, () => console.log(`🌐 API online na porta ${PORT}`));

export default app;
