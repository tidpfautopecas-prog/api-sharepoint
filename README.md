
# API SharePoint - Global Plastic

API Node.js para integração com SharePoint, permitindo upload automático de PDFs de laudos na pasta específica.

## 🚀 Configuração

### 1. Instalar dependências
```bash
npm install
```

### 2. Configurar variáveis de ambiente
O arquivo `.env` já está configurado com suas credenciais:

```env
TENANT_ID=<SEU_TENANT_ID>
CLIENT_ID=<SEU_CLIENT_ID>
CLIENT_SECRET=*****SECRET*****
SITE_ID=<SEU_SITE_ID>
LIBRARY_NAME=Documentos%20Compartilhados
FOLDER_PATH=Laudos
PORT=3000
```

### 3. Iniciar o servidor
```bash
# Modo produção
npm start

# Modo desenvolvimento (com auto-reload)
npm run dev
```

## 📋 Endpoints Disponíveis

### `GET /status`
Verifica o status da API e configurações.

### `GET /test-connection`
Testa a conectividade com o SharePoint.

### `POST /create-folder`
Cria a pasta "Laudos" no SharePoint se não existir.

### `POST /upload-pdf`
Upload de PDF para o SharePoint.

**Body:**
```json
{
  "fileName": "Laudo_123_15-01-2024_14h30min.pdf",
  "fileBase64": "base64_do_arquivo...",
  "ticketNumber": "#123",
  "ticketTitle": "Título do laudo",
  "isReport": false
}
```

## 🔧 Como usar no frontend

```javascript
// Enviar PDF para SharePoint
const response = await fetch('http://localhost:3000/upload-pdf', {
  method: 'POST',
  headers: { 
    'Content-Type': 'application/json',
    'Accept': 'application/json'
  },
  body: JSON.stringify({ 
    fileName, 
    fileBase64,
    ticketNumber: ticket.numero,
    ticketTitle: ticket.titulo
  })
});

const result = await response.json();
if (response.ok) {
  console.log('✅ PDF salvo no SharePoint!');
} else {
  console.error('❌ Erro:', result.error);
}
```

## 🧪 Testar a API

1. **Verificar status:**
   ```bash
   curl http://localhost:3000/status
   ```

2. **Testar conexão:**
   ```bash
   curl http://localhost:3000/test-connection
   ```

3. **Criar pasta Laudos:**
   ```bash
   curl -X POST http://localhost:3000/create-folder
   ```

## 📁 Estrutura de Pastas no SharePoint

```
SharePoint Site (GLB-FS)
└── Documentos Compartilhados/
    └── Laudos/
        ├── Laudo_123_15-01-2024_14h30min.pdf
        ├── Relatorio_Laudos_15_01_2024.pdf
        └── ...
```

## 🔒 Segurança

- ✅ Credenciais Microsoft oficiais
- ✅ Token de acesso renovado automaticamente
- ✅ CORS configurado para o frontend
- ✅ Validação de dados de entrada
- ✅ Logs detalhados para monitoramento

## 🚨 Troubleshooting

### Erro de autenticação
- Verifique se as credenciais no `.env` estão corretas
- Confirme se o aplicativo tem permissões no Azure AD

### Erro de upload
- Verifique se a pasta "Laudos" existe (use `/create-folder`)
- Confirme permissões de escrita no SharePoint
- Teste a conectividade com `/test-connection`

### Pasta não encontrada
- Execute `POST /create-folder` para criar a pasta automaticamente
- Verifique se `LIBRARY_NAME` e `FOLDER_PATH` estão corretos

## 📊 Logs

A API gera logs detalhados:
- 🔐 Autenticação Microsoft Graph
- ⬆️ Uploads de arquivos
- ✅ Sucessos e falhas
- 🧪 Testes de conectividade

## 🎯 Próximos Passos

1. Iniciar a API: `npm start`
2. Testar conexão: `GET /test-connection`
3. Criar pasta se necessário: `POST /create-folder`
4. Integrar com o frontend React
5. Monitorar logs de upload
"# api-sharepoint" 
