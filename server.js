import express from 'express';
import bodyParser from 'body-parser';
import fetch from 'node-fetch';
import dotenv from 'dotenv';
import cors from 'cors'; // ✅ IMPORTANTE: O pacote 'cors' é essencial

dotenv.config();

const app = express();

// =================================================================================
// 🛡️ CORREÇÃO DE CORS (O QUE RESOLVE O SEU ERRO ATUAL)
// =================================================================================
app.use(cors({
    origin: '*', // Permite conexões de qualquer lugar (incluindo seu localhost)
    methods: ['GET', 'POST', 'PUT', 'DELETE', 'OPTIONS'],
    allowedHeaders: ['Content-Type', 'Authorization', 'X-Requested-With', 'Accept', 'Origin'],
    credentials: true
}));

// Garante que as requisições de verificação (preflight) funcionem
app.options('*', cors()); 
// =================================================================================

app.use(bodyParser.json({ limit: '50mb' }));

console.log('🚀 API SharePoint Global Plastic a iniciar...');
console.log(`📁 Site: ${process.env.SITE_ID}`);
console.log(`📂 Biblioteca: ${process.env.LIBRARY_NAME}`);
console.log(`📄 Lista: ${process.env.LIST_NAME}`);
console.log(`📍 Pasta: ${process.env.FOLDER_PATH}`);

// =================================================================================
// 📋 MAPEAMENTO DOS NOMES INTERNOS (BASEADO NO QUE VOCÊ ENVIOU)
// =================================================================================
const COLUMN_MAPPING = {
    // Título (Padrão)
    'Title': (row) => row['N° do ticket'] + ' - ' + row.Item + ' - ' + row.Motivo,
    
    // Nomes Internos que você encontrou nas configurações:
    'N_x00b0_doticket': (row) => row['N° do ticket'],
    'NomedoCliente': (row) => row['Nome do Cliente'],
    'Item': (row) => row.Item,
    'Qtde': (row) => String(row.Qtde), // Força texto para evitar erro de tipo
    'Motivo': (row) => row.Motivo,
    'Origemdodefeito': (row) => row['Origem do defeito'],
    'Disposi_x00e7__x00e3_o': (row) => row.Disposição,
    'Disposi_x00e7__x00e3_odaspe_x00e': (row) => row['Disposição das peças'],

};
// =================================================================================

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
    if (!res.ok) {
        const errorText = await res.text();
        throw new Error(`Não foi possível encontrar as bibliotecas do site. Status: ${res.status} - ${errorText}`);
    }
    const { value: drives } = await res.json();
    const library = drives.find(d => d.name === process.env.LIBRARY_NAME);
    if (!library) {
        throw new Error(`A biblioteca de documentos chamada "${process.env.LIBRARY_NAME}" não foi encontrada no site.`);
    }
    console.log(`✅ ID da Biblioteca "${library.name}" encontrado: ${library.id}`);
    return library.id;
}

async function getListId(accessToken) {
    const listName = process.env.LIST_NAME;
    if (!listName) {
        throw new Error("Variável de ambiente LIST_NAME não está definida.");
    }
    
    // Busca a lista pelo nome exato ("Laudo")
    const url = `https://graph.microsoft.com/v1.0/sites/${process.env.SITE_ID}/lists?$filter=displayName eq '${encodeURIComponent(listName)}'`;
    
    const res = await fetch(url, { headers: { 'Authorization': `Bearer ${accessToken}` } });
    if (!res.ok) {
        const errorText = await res.text();
        throw new Error(`Não foi possível procurar as Listas do site. Status: ${res.status} - ${errorText}`);
    }
    
    const { value: lists } = await res.json();
    
    if (lists.length > 0) {
        console.log(`✅ ID da Lista "${lists[0].displayName}" encontrado: ${lists[0].id}`);
        return lists[0].id;
    } else {
        console.error(`❌ A Lista "${listName}" não foi encontrada. Verifique se o nome no Render é exatamente "Laudo".`);
        throw new Error(`A Lista "${listName}" não foi encontrada.`);
    }
}

app.get('/', (req, res) => {
    res.json({
      message: 'Hello from Global Plastic SharePoint API!',
      status: 'online',
      timestamp: new Date().toISOString(),
    });
});

// ROTA 1: Upload do PDF
app.post('/upload-pdf', async (req, res) => {
  const { fileName, fileBase64 } = req.body;
  if (!fileName || !fileBase64) {
    return res.status(400).json({ error: 'Dados obrigatórios ausentes' });
  }

  try {
    console.log(`📄 A iniciar upload para: ${fileName}`);
    const accessToken = await getAccessToken();
    const driveId = await getDriveId(accessToken);
    const encodedFolder = encodeURIComponent(process.env.FOLDER_PATH);
    const encodedFileName = encodeURIComponent(fileName);
    const uploadUrl = `https://graph.microsoft.com/v1.0/drives/${driveId}/root:/${encodedFolder}/${encodedFileName}:/content`;
    
    console.log(`⬆️ A enviar para o URL correto: ${uploadUrl}`);

    const response = await fetch(uploadUrl, {
      method: 'PUT',
      headers: { 'Authorization': `Bearer ${accessToken}`, 'Content-Type': 'application/pdf' },
      body: Buffer.from(fileBase64, 'base64')
    });

    if (!response.ok) {
      const errorText = await response.text();
      throw new Error(`SharePoint Error ${response.status}: ${errorText}`);
    }

    const result = await response.json();
    console.log(`✅ Upload concluído com sucesso para: ${result.webUrl}`);
    res.status(200).json({ success: true, sharePointUrl: result.webUrl });

  } catch (error) {
    console.error(`❌ Erro no upload:`, error.message);
    res.status(500).json({ success: false, error: 'Falha ao enviar PDF', details: error.message });
  }
});

// ROTA 2: Upload dos Dados da Lista
app.post('/upload-list-data', async (req, res) => {
    const { listData } = req.body;
    
    if (!listData || listData.length === 0) {
        return res.status(400).json({ success: false, error: 'Nenhum dado de lista fornecido.' });
    }

    try {
        console.log(`📋 A iniciar inserção de ${listData.length} itens na Lista do SharePoint.`);
        const accessToken = await getAccessToken();
        const listId = await getListId(accessToken); 

        const listItemsUrl = `https://graph.microsoft.com/v1.0/sites/${process.env.SITE_ID}/lists/${listId}/items`;

        const insertionPromises = listData.map(async (row) => {
            
            const itemFields = {};
            // Mapeia os dados usando os nomes internos corretos
            for (const key in COLUMN_MAPPING) {
                if (Object.prototype.hasOwnProperty.call(COLUMN_MAPPING, key)) {
                     itemFields[key] = COLUMN_MAPPING[key](row);
                }
            }
            
            const itemResponse = await fetch(listItemsUrl, {
                method: 'POST',
                headers: { 
                    'Authorization': `Bearer ${accessToken}`, 
                    'Content-Type': 'application/json' 
                },
                body: JSON.stringify({ fields: itemFields })
            });

            if (!itemResponse.ok) {
                const errorText = await itemResponse.text();
                console.error(`Detalhe do Erro SharePoint (Item): ${errorText}`);
                throw new Error(`Erro ao inserir item na Lista. Status: ${itemResponse.status}.`);
            }
            return itemResponse.json();
        });

        await Promise.all(insertionPromises);

        console.log(`✅ Inserção de todos os ${listData.length} itens na Lista concluída.`);
        res.status(200).json({ success: true, message: 'Dados da lista enviados e salvos com sucesso.' });

    } catch (error) {
        console.error(`❌ Erro no upload da lista:`, error.message);
        res.status(500).json({ success: false, error: 'Falha ao enviar dados da lista', details: error.message });
    }
});

// ROTA 3: Exclusão do PDF
app.delete('/delete-pdf-by-ticket-number/:ticketNumber', async (req, res) => {
    const { ticketNumber } = req.params;
    if (!ticketNumber) return res.status(400).json({ error: 'Número do ticket é obrigatório.' });

    try {
        const accessToken = await getAccessToken();
        const driveId = await getDriveId(accessToken);
        const encodedFolder = encodeURIComponent(process.env.FOLDER_PATH);
        
        const listUrl = `https://graph.microsoft.com/v1.0/drives/${driveId}/root:/${encodedFolder}:/children`;
        
        const listResponse = await fetch(listUrl, { headers: { 'Authorization': `Bearer ${accessToken}` } });
        if (!listResponse.ok) throw new Error(`Não foi possível listar os ficheiros. Status: ${listResponse.status}`);
        
        const { value: allFiles } = await listResponse.json();
        const fileNamePrefix = `Laudo - ${ticketNumber}-`;
        const filesToDelete = allFiles.filter(file => file.name.startsWith(fileNamePrefix));

        if (filesToDelete.length === 0) {
            return res.status(200).json({ success: true, message: `Nenhum PDF encontrado para o laudo ${ticketNumber}.` });
        }

        const deletePromises = filesToDelete.map(file => {
            const deleteUrl = `https://graph.microsoft.com/v1.0/drives/${driveId}/items/${file.id}`;
            return fetch(deleteUrl, { method: 'DELETE', headers: { 'Authorization': `Bearer ${accessToken}` } });
        });

        await Promise.all(deletePromises);
        res.status(200).json({ success: true, message: `${filesToDelete.length} PDF(s) excluídos com sucesso.` });
    } catch (error) {
        console.error(`❌ Erro na exclusão do laudo ${ticketNumber}:`, error.message);
        res.status(500).json({ success: false, error: `Falha ao excluir PDF(s) do laudo ${ticketNumber}`, details: error.message });
    }
});

const PORT = process.env.PORT || 3000;
app.listen(PORT, () => {
  console.log(`🌐 Servidor a rodar na porta ${PORT}`);
  console.log('✅ API SharePoint Global Plastic pronta!');
});

export default app;
