// server.js - VERSÃO MELHORADA COM MAIS LOGS E ROTA DE TESTE
import express from 'express';
import bodyParser from 'body-parser';
import fetch from 'node-fetch';
import dotenv from 'dotenv';
import cors from 'cors';

dotenv.config();

const app = express();
app.use(cors());
app.use(bodyParser.json({ limit: '50mb' }));

// --- As configurações são lidas diretamente do arquivo .env ---
const SITE_ID = process.env.SITE_ID;
const FOLDER_PATH = process.env.FOLDER_PATH || 'Laudos';

// Verifica se o SITE_ID foi configurado corretamente
if (!SITE_ID || SITE_ID === 'GLB-FS' || !SITE_ID.includes(',')) {
  console.error('❌ ERRO FATAL: A variável SITE_ID não foi configurada corretamente no arquivo .env.');
  console.error('   Por favor, insira o ID composto completo do site (com hostname e vírgulas) no arquivo .env.');
  process.exit(1);
}

async function getAccessToken() {
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
  if (!data.access_token) throw new Error('Não foi possível obter token: ' + JSON.stringify(data));
  return data.access_token;
}

// --- MELHORIA: Rota raiz para teste rápido no navegador ---
app.get('/', (req, res) => {
  res.send('<h1>API SharePoint está online e funcionando!</h1><p>Use a rota POST /upload-pdf para enviar arquivos.</p>');
});

app.get('/status', (req, res) => {
  res.json({ status: 'online', siteId: SITE_ID, folder: FOLDER_PATH });
});

app.post('/upload-pdf', async (req, res) => {
  console.log('✅ Rota /upload-pdf foi chamada.');
  try {
    const { fileName, fileBase64 } = req.body;
    if (!fileName || !fileBase64) return res.status(400).json({ error: 'fileName e fileBase64 são obrigatórios' });

    // --- MELHORIA: Log para saber qual arquivo está sendo processado ---
    console.log(`   📄 Recebido arquivo para upload: ${fileName}`);

    const accessToken = await getAccessToken();

    const uploadUrl = `https://graph.microsoft.com/v1.0/sites/${SITE_ID}/drive/root:/${FOLDER_PATH}/${fileName}:/content`;

    console.log(`   📍 Enviando para URL do SharePoint:`, uploadUrl);

    const response = await fetch(uploadUrl, {
      method: 'PUT',
      headers: {
        'Authorization': `Bearer ${accessToken}`,
        'Content-Type': 'application/pdf'
      },
      body: Buffer.from(fileBase64, 'base64')
    });

    if (!response.ok) {
      const errText = await response.text();
      throw new Error(`SharePoint Error ${response.status}: ${errText}`);
    }

    const result = await response.json();

    // --- MELHORIA: Log de sucesso com a URL final do arquivo ---
    console.log(`   🎉 Sucesso! Arquivo salvo em: ${result.webUrl}`);

    res.json({ success: true, fileName, sharePointUrl: result.webUrl });
  } catch (error) {
    console.error('❌ Erro no upload:', error.message);
    res.status(500).json({ success: false, error: error.message });
  }
});

const PORT = process.env.PORT || 3000;
app.listen(PORT, () => {
  console.log(`🌐 API rodando na porta ${PORT}`);
});