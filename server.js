// server.js

import express from 'express';
import bodyParser from 'body-parser';
import fetch from 'node-fetch';
import dotenv from 'dotenv';
import cors from 'cors';

dotenv.config();

const app = express();

// Middlewares
app.use(cors()); // Permitir CORS para o frontend
app.use(bodyParser.json({ limit: '50mb' }));

// Log de inicialização
console.log('🚀 API SharePoint Global Plastic a iniciar...');
console.log(`📁 Site: ${process.env.SITE_ID}`);
console.log(`📂 Biblioteca: ${process.env.LIBRARY_NAME}`);
console.log(`📍 Pasta: ${process.env.FOLDER_PATH}`);

// Função para obter token do Azure AD com retry
async function getAccessToken(retries = 3) {
  for (let i = 0; i < retries; i++) {
    try {
      const params = new URLSearchParams();
      params.append('client_id', process.env.CLIENT_ID);
      params.append('scope', 'https://graph.microsoft.com/.default');
      params.append('client_secret', process.env.CLIENT_SECRET);
      params.append('grant_type', 'client_credentials');

      console.log(`🔐 Tentativa ${i + 1} - A obter token de acesso...`);

      const res = await fetch(`https://login.microsoftonline.com/${process.env.TENANT_ID}/oauth2/v2.0/token`, {
        method: 'POST',
        body: params,
        headers: { 'Content-Type': 'application/x-www-form-urlencoded' }
      });

      const data = await res.json();
      
      if (!data.access_token) {
        throw new Error(`Erro na autenticação: ${data.error_description || data.error}`);
      }

      console.log('✅ Token obtido com sucesso');
      return data.access_token;
      
    } catch (error) {
      console.error(`❌ Tentativa ${i + 1} falhou:`, error.message);
      if (i === retries - 1) throw error;
      await new Promise(resolve => setTimeout(resolve, 1000 * (i + 1))); // Delay progressivo
    }
  }
}

// Rota de status da API
app.get('/status', (req, res) => {
  res.json({
    status: 'online',
    timestamp: new Date().toISOString(),
    config: {
      siteId: process.env.SITE_ID,
      library: process.env.LIBRARY_NAME,
      folder: process.env.FOLDER_PATH,
      tenant: process.env.TENANT_ID
    }
  });
});

// Rota principal para upload de PDF
app.post('/upload-pdf', async (req, res) => {
    // ... (esta rota permanece sem alterações)
});

// ✅ NOVA ROTA PARA EXCLUIR PDF POR NÚMERO DE TICKET
app.delete('/delete-pdf-by-ticket-number/:ticketNumber', async (req, res) => {
    const startTime = Date.now();
    const { ticketNumber } = req.params;

    if (!ticketNumber) {
        return res.status(400).json({ error: 'Número do ticket é obrigatório.' });
    }

    console.log(`🗑️ Recebida solicitação para excluir PDFs do laudo: ${ticketNumber}`);

    try {
        const accessToken = await getAccessToken();

        // 1. Listar todos os ficheiros na pasta de laudos
        const listUrl = `https://graph.microsoft.com/v1.0/sites/${process.env.SITE_ID}/drives/root:/${process.env.LIBRARY_NAME}/${process.env.FOLDER_PATH}:/children`;
        
        const listResponse = await fetch(listUrl, {
            headers: { 'Authorization': `Bearer ${accessToken}` },
        });

        if (!listResponse.ok) {
            throw new Error(`Não foi possível listar os ficheiros na pasta Laudos. Status: ${listResponse.status}`);
        }

        const { value: allFiles } = await listResponse.json();

        // 2. Filtrar para encontrar ficheiros que correspondam ao padrão do ticket
        const fileNamePrefix = `Laudo - ${ticketNumber}-`;
        const filesToDelete = allFiles.filter(file => file.name.startsWith(fileNamePrefix));

        if (filesToDelete.length === 0) {
            console.log(`🟡 Nenhum PDF encontrado para o laudo ${ticketNumber}. Nenhuma ação necessária.`);
            return res.status(200).json({
                success: true,
                message: `Nenhum PDF encontrado no SharePoint para o laudo ${ticketNumber}.`,
            });
        }

        console.log(`🔎 Encontrados ${filesToDelete.length} PDFs para excluir...`);
        
        // 3. Excluir cada ficheiro encontrado
        const deletePromises = filesToDelete.map(file => {
            console.log(`   - A excluir: ${file.name}`);
            const deleteUrl = `https://graph.microsoft.com/v1.0/sites/${process.env.SITE_ID}/drives/root/items/${file.id}`;
            return fetch(deleteUrl, {
                method: 'DELETE',
                headers: { 'Authorization': `Bearer ${accessToken}` },
            });
        });

        await Promise.all(deletePromises);

        const deleteTime = Date.now() - startTime;
        console.log(`✅ Exclusão concluída em ${deleteTime}ms`);

        res.status(200).json({
            success: true,
            message: `${filesToDelete.length} PDF(s) do laudo ${ticketNumber} foram excluídos com sucesso do SharePoint.`,
            deletedFiles: filesToDelete.map(f => f.name),
        });

    } catch (error) {
        const deleteTime = Date.now() - startTime;
        console.error(`❌ Erro na exclusão do laudo ${ticketNumber} (${deleteTime}ms):`, error.message);
        res.status(500).json({
            success: false,
            error: `Falha ao excluir PDF(s) do laudo ${ticketNumber}`,
            details: error.message,
        });
    }
});


// Rota para testar conectividade
app.get('/test-connection', async (req, res) => {
    // ... (esta rota permanece sem alterações)
});

// Rota para criar a pasta Laudos se não existir
app.post('/create-folder', async (req, res) => {
    // ... (esta rota permanece sem alterações)
});

// Middleware de erro global
app.use((error, req, res, next) => {
  console.error('💥 Erro não tratado:', error);
  res.status(500).json({
    success: false,
    error: 'Erro interno do servidor',
    timestamp: new Date().toISOString()
  });
});

// Iniciar servidor
const PORT = process.env.PORT || 3000;
app.listen(PORT, () => {
  console.log(`🌐 Servidor a rodar na porta ${PORT}`);
  console.log(`📋 Status: http://localhost:${PORT}/status`);
  console.log(`🧪 Teste: http://localhost:${PORT}/test-connection`);
  console.log(`📁 Criar pasta: http://localhost:${PORT}/create-folder`);
  console.log('✅ API SharePoint Global Plastic pronta!');
});

export default app;
