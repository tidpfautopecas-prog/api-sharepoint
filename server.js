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

const COLUMN_MAPPING = {
    'Title': (row) => row.Title,
    'N_x00b0_doticket': (row) => row.ticketNumber,
    'NomedoCliente': (row) => row.nomeCliente,
    'Item': (row) => row.item,
    'Qtde': (row) => String(row.qtde),
    'Motivo': (row) => row.motivo,
    'Origemdodefeito': (row) => row.origemDefeito,
    'Disposi_x00e7__x00e3_o': (row) => row.disposicao,
    'Disposi_x00e7__x00e3_odaspe_x00e': (row) => row.disposicaoPecas,
    'DatadeGera_x00e7__x00e3_o': (row) => row.dataGeracao || '',
    'Foto1': (row) => row.foto1 || null,
    'Foto2': (row) => row.foto2 || null,
    'Foto3': (row) => row.foto3 || null,
    'Foto4': (row) => row.foto4 || null,
    'Foto5': (row) => row.foto5 || null,
    'Foto6': (row) => row.foto6 || null,
    'Foto7': (row) => row.foto7 || null,
    'Foto8': (row) => row.foto8 || null,
    'Foto9': (row) => row.foto9 || null,
    'Foto10': (row) => row.foto10 || null,
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
      if (!data.access_token) throw new Error();
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
    if (!res.ok) throw new Error();
    const { value: drives } = await res.json();
    const library = drives.find(d => d.name === process.env.LIBRARY_NAME);
    if (!library) throw new Error();
    return library.id;
}

async function getListId(accessToken) {
    const listName = process.env.LIST_NAME;
    const url = `https://graph.microsoft.com/v1.0/sites/${process.env.SITE_ID}/lists?$filter=displayName eq '${encodeURIComponent(listName)}'`;
    const res = await fetch(url, { headers: { 'Authorization': `Bearer ${accessToken}` } });
    if (!res.ok) throw new Error();
    const { value: lists } = await res.json();
    if (lists.length > 0) return lists[0].id;
    throw new Error();
}

app.get('/', (req, res) => res.json({ status: 'online' }));

app.get('/status', (req, res) => res.json({ status: 'online' }));

app.get('/check-status/:ticketNumber', async (req, res) => {
    const { ticketNumber } = req.params;
    try {
        const accessToken = await getAccessToken();
        const siteId = process.env.SITE_ID;
        const driveId = await getDriveId(accessToken);
        const listId = await getListId(accessToken);

        const listUrl = `https://graph.microsoft.com/v1.0/sites/${siteId}/lists/${listId}/items?$filter=fields/N_x00b0_doticket eq '${ticketNumber}'`;
        const listRes = await fetch(listUrl, { headers: { 'Authorization': `Bearer ${accessToken}`, 'Prefer': 'HonorNonIndexedQueriesWarningMayFailRandomly' } });
        
        let existsInList = false;
        if (listRes.ok) {
             const data = await listRes.json();
             existsInList = data.value && data.value.length > 0;
        }

        const encodedFolder = encodeURIComponent(process.env.FOLDER_PATH);
        const pdfNamePart = `Laudo - ${ticketNumber}-`;
        const driveUrl = `https://graph.microsoft.com/v1.0/drives/${driveId}/root:/${encodedFolder}:/search(q='${pdfNamePart}')`;
        const driveRes = await fetch(driveUrl, { headers: { 'Authorization': `Bearer ${accessToken}` } });
        
        let existsInPdf = false;
        if (driveRes.ok) {
            const data = await driveRes.json();
            existsInPdf = data.value && data.value.some(f => f.name.includes(ticketNumber) && f.name.endsWith('.pdf'));
        }

        res.json({ existsInList, existsInPdf });

    } catch (error) {
        res.status(500).json({ error: error.message });
    }
});

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

    if (!response.ok) throw new Error();
    const result = await response.json();
    res.status(200).json({ success: true, sharePointUrl: result.webUrl });
  } catch (error) {
    res.status(500).json({ success: false, error: error.message });
  }
});

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

            if (!itemResponse.ok) {
                const errText = await itemResponse.text();
                throw new Error(errText);
            }
            return itemResponse.json();
        });

        await Promise.all(insertionPromises);
        res.status(200).json({ success: true });
    } catch (error) {
        res.status(500).json({ success: false, error: error.message });
    }
});

app.delete('/delete-pdf-by-ticket-number/:ticketNumber', async (req, res) => {
    const { ticketNumber } = req.params;
    if (!ticketNumber) return res.status(400).json({ error: 'Ticket obrigatório' });

    try {
        const accessToken = await getAccessToken();
        const driveId = await getDriveId(accessToken);
        const encodedFolder = encodeURIComponent(process.env.FOLDER_PATH);
        const listUrl = `https://graph.microsoft.com/v1.0/drives/${driveId}/root:/${encodedFolder}:/children`;
        
        const listResponse = await fetch(listUrl, { headers: { 'Authorization': `Bearer ${accessToken}` } });
        if (!listResponse.ok) throw new Error();
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

const PORT = process.env.PORT || 3001;
app.listen(PORT, () => {});
