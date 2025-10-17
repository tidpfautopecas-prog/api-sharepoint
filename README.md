# API SharePoint - Global Plastic

API Node.js para integração com SharePoint, permitindo upload e exclusão automáticos de PDFs de laudos.

## 🚀 Configuração

### 1. Instalar dependências
`bash
npm install
`

### 2. Configurar variáveis de ambiente
[cite_start]O ficheiro `.env` já está configurado com as suas credenciais. [cite: 1]

### 3. Iniciar o servidor
`bash
# Modo produção
npm start

# Modo desenvolvimento (com auto-reload)
npm run dev
`

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
`json
{
  "fileName": "Laudo_123_15-01-2024_14h30min.pdf",
  "fileBase64": "base64_do_ficheiro...",
  "ticketNumber": "#123",
  "ticketTitle": "Título do laudo",
  "isReport": false
}
`

### `DELETE /delete-pdf-by-ticket-number/:ticketNumber`
Exclui todos os PDFs no SharePoint que correspondem a um número de ticket específico.

**Exemplo de uso:**
`bash
curl -X DELETE http://localhost:3000/delete-pdf-by-ticket-number/SR-12345
`

## 🔧 Como usar no frontend
(Exemplos de código para upload e outras operações)

## 🧪 Testar a API

1. **Verificar status:**
   `bash
   curl http://localhost:3000/status
   `

2. **Testar conexão:**
   `bash
   curl http://localhost:3000/test-connection
   `

3. **Criar pasta Laudos:**
   `bash
   curl -X POST http://localhost:3000/create-folder
   `

## 📁 Estrutura de Pastas no SharePoint

`
SharePoint Site (GLB-FS)
└── Documentos Compartilhados/
    └── Laudos/
        ├── Laudo_123_15-01-2024_14h30min.pdf
        ├── Relatorio_Laudos_15_01_2024.pdf
        └── ...
`

## 🔒 Segurança

- ✅ Credenciais Microsoft oficiais
- ✅ Token de acesso renovado automaticamente
- ✅ CORS configurado para o frontend
- ✅ Validação de dados de entrada
- ✅ Logs detalhados para monitorização

## 🚨 Troubleshooting

### Erro de autenticação
- Verifique se as credenciais no `.env` estão corretas
- Confirme se a aplicação tem permissões no Azure AD

### Erro de upload
- Verifique se a pasta "Laudos" existe (use `/create-folder`)
- Confirme permissões de escrita no SharePoint
- Teste a conectividade com `/test-connection`

### Pasta não encontrada
- Execute `POST /create-folder` para criar a pasta automaticamente
- Verifique se `LIBRARY_NAME` e `FOLDER_PATH` estão corretos

## 📊 Logs

A API gera logs detalhados para todas as operações, incluindo autenticação, uploads, exclusões e testes.
