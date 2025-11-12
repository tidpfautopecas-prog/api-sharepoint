import express from 'express';
import bodyParser from 'body-parser';
import fetch from 'node-fetch';
import dotenv from 'dotenv';
import cors from 'cors';

dotenv.config();

const app = express();

app.use(cors());
app.use(bodyParser.json({ limit: '50mb' }));

console.log('🚀 API SharePoint Global Plastic a iniciar...');
console.log(`📁 Site: ${process.env.SITE_ID}`);
console.log(`📂 Biblioteca: ${process.env.LIBRARY_NAME}`);
console.log(`📄 Lista: ${process.env.LIST_NAME}`); // ✅ Garanta que esta variável de ambiente exista no Render
console.log(`📍 Pasta: ${process.env.FOLDER_PATH}`);

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

// ✅ FUNÇÃO EXISTENTE: Encontra o ID da biblioteca de documentos (Drive)
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

// ✅ NOVA FUNÇÃO: Encontra o ID da Lista do SharePoint
async function getListId(accessToken) {
    const listName = process.env.LIST_NAME;
    if (!listName) {
        throw new Error("Variável de ambiente LIST_NAME não está definida.");
    }
    const url = `https://graph.microsoft.com/v1.0/sites/${process.env.SITE_ID}/lists`;
    
    const res = await fetch(url, { headers: { 'Authorization': `Bearer ${accessToken}` } });
    if (!res.ok) {
        const errorText = await res.text();
        throw new Error(`Não foi possível listar as Listas do site. Status: ${res.status} - ${errorText}`);
    }
    
    const { value: lists } = await res.json();
    // Filtra pela propriedade 'displayName' da Lista
    const targetList = lists.find(l => l.displayName === listName);
    
    if (!targetList) {
        throw new Error(`A Lista do SharePoint chamada "${listName}" não foi encontrada no site.`);
    }
    
    console.log(`✅ ID da Lista "${targetList.displayName}" encontrado: ${targetList.id}`);
    return targetList.id;
}


app.get('/', (req, res) => {
    res.json({
      message: 'Hello from Global Plastic SharePoint API!',
      status: 'online',
      timestamp: new Date().toISOString(),
    });
});

// ✅ ENDPOINT DE UPLOAD DE PDF (EXISTENTE)
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

// ✅ ENDPOINT NOVO E CORRIGIDO: Rota para receber e inserir os dados da Lista
app.post('/upload-list-data', async (req, res) => {
    const { listData } = req.body;
    
    if (!listData || listData.length === 0) {
        return res.status(400).json({ success: false, error: 'Nenhum dado de lista fornecido.' });
    }

    try {
        console.log(`📋 A iniciar inserção de ${listData.length} itens na Lista do SharePoint.`);
        const accessToken = await getAccessToken();
        const listId = await getListId(accessToken); // Obtém o ID da Lista

        const listItemsUrl = `https://graph.microsoft.com/v1.0/sites/${process.env.SITE_ID}/lists/${listId}/items`;

        const insertionPromises = listData.map(async (row) => {
            
            // ⭐⭐⭐ ATENÇÃO AQUI: MAPEAMENTO DE NOMES INTERNOS ⭐⭐⭐
            // Os nomes à ESQUERDA (ex: 'TicketNumber') são SUGESTÕES.
            // Substitua-os pelos Nomes Internos REAIS da sua lista do SharePoint.
            const itemFields = {
                // 'Title' é (geralmente) obrigatório. Mapeie para algo útil.
                'Title': row['N° do ticket'] + ' - ' + row.Item + ' - ' + row.Motivo, 
                
                // --- AJUSTE ESTES NOMES INTERNOS ---
                // 'NomeInternoDaSuaLista': row['NomeDoFrontend']
                
                'TicketNumber': row['N° do ticket'],        // Sugestão para 'N° do ticket'
                'CustomerName': row['Nome do Cliente'],  // Sugestão para 'Nome do Cliente'
                'Item': row.Item,                        // Sugestão para 'Item'
                'Qtde': row.Qtde,                        // Sugestão para 'Qtde'
                'Motivo': row.Motivo,                    // Sugestão para 'Motivo'
                'OrigemDefeito': row['Origem do defeito'], // Sugestão para 'Origem do defeito'
                'Disposicao': row.Disposição,              // Sugestão para 'Disposição'
                'DisposicaoPecas': row['Disposição das peças'], // Sugestão para 'Disposição das peças'
                'GenerationDate': row['Data de Geração'], // Sugestão para 'Data de Geração'
            };
            
            const itemResponse = await fetch(listItemsUrl, {
                method: 'POST',
                headers: { 
                    'Authorization': `Bearer ${accessToken}`, 
                    'Content-Type': 'application/json' 
                },
                body: JSON.stringify({ fields: itemFields })
            });

            if (!itemResponse.ok) {
                // Log detalhado para depuração
                const errorText = await itemResponse.text();
                console.error(`Detalhe do Erro SharePoint para o Ticket ${row['N° do ticket']}:`, errorText);
                throw new Error(`Erro ao inserir item na Lista. Status: ${itemResponse.status}. Verifique os Nomes Internos das colunas.`);
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


// ✅ ENDPOINT DE EXCLUSÃO (EXISTENTE)
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
