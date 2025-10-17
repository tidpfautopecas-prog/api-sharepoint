import express from 'express';
import bodyParser from 'body-parser';
import fetch from 'node-fetch';
import dotenv from 'dotenv';
import cors from 'cors';

dotenv.config();

const app = express();

// Middlewares
app.use(cors());
app.use(bodyParser.json({ limit: '50mb' }));

console.log('🚀 API SharePoint Global Plastic a iniciar...');
console.log(`📁 Site: ${process.env.SITE_ID}`);
console.log(`📂 Biblioteca: ${process.env.LIBRARY_NAME}`);
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
      if (!data.access_token) {
        throw new Error(`Erro na autenticação: ${data.error_description || data.error}`);
      }
      return data.access_token;
    } catch (error) {
      console.error(`❌ Tentativa ${i + 1} de obter token falhou:`, error.message);
      if (i === retries - 1) throw error;
      await new Promise(resolve => setTimeout(resolve, 1000 * (i + 1)));
    }
  }
}

function buildGraphUrl(path) {
  const siteId = process.env.SITE_ID;
  return `https://graph.microsoft.com/v1.0/sites/${siteId}/${path}`;
}

// ✅ NOVA ROTA PRINCIPAL PARA RESOLVER O ERRO 404
app.get('/', (req, res) => {
    res.json({
      message: 'Hello from Global Plastic SharePoint API!',
      status: 'online',
      timestamp: new Date().toISOString(),
    });
});

app.post('/upload-pdf', async (req, res) => {
  const { fileName, fileBase64 } = req.body;
  if (!fileName || !fileBase64) {
    return res.status(400).json({ error: 'Dados obrigatórios ausentes' });
  }
  try {
    const accessToken = await getAccessToken();
    const encodedLibrary = encodeURIComponent(process.env.LIBRARY_NAME);
    const encodedFolder = encodeURIComponent(process.env.FOLDER_PATH);
    const encodedFileName = encodeURIComponent(fileName);
    const uploadPath = `drives/root:/${encodedLibrary}/${encodedFolder}/${encodedFileName}:/content`;
    const uploadUrl = buildGraphUrl(uploadPath);
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
    res.status(200).json({ success: true, sharePointUrl: result.webUrl });
  } catch (error) {
    console.error(`❌ Erro no upload:`, error.message);
    res.status(500).json({ success: false, error: 'Falha ao enviar PDF', details: error.message });
  }
});

app.delete('/delete-pdf-by-ticket-number/:ticketNumber', async (req, res) => {
    const { ticketNumber } = req.params;
    if (!ticketNumber) return res.status(400).json({ error: 'Número do ticket é obrigatório.' });
    try {
        const accessToken = await getAccessToken();
        const encodedLibrary = encodeURIComponent(process.env.LIBRARY_NAME);
        const encodedFolder = encodeURIComponent(process.env.FOLDER_PATH);
        const listPath = `drives/root:/${encodedLibrary}/${encodedFolder}:/children`;
        const listUrl = buildGraphUrl(listPath);
        
        const listResponse = await fetch(listUrl, { headers: { 'Authorization': `Bearer ${accessToken}` } });
        if (!listResponse.ok) throw new Error(`Não foi possível listar os ficheiros. Status: ${listResponse.status}`);
        
        const { value: allFiles } = await listResponse.json();
        const fileNamePrefix = `Laudo - ${ticketNumber}-`;
        const filesToDelete = allFiles.filter(file => file.name.startsWith(fileNamePrefix));

        if (filesToDelete.length === 0) {
            return res.status(200).json({ success: true, message: `Nenhum PDF encontrado para o laudo ${ticketNumber}.` });
        }

        const deletePromises = filesToDelete.map(file => {
            const deletePath = `drives/root/items/${file.id}`;
            const deleteUrl = buildGraphUrl(deletePath);
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
