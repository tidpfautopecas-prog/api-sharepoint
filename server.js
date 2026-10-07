import 'dotenv/config';
import express from 'express';
import cors from 'cors';
import axios from 'axios';

const app = express();
const PORT = process.env.PORT || 3000;

// ============================================================
// MIDDLEWARES
// ============================================================

app.use(cors());

app.use(
    express.json({
        limit: '50mb'
    })
);

app.use(
    express.urlencoded({
        extended: true,
        limit: '50mb'
    })
);

// ============================================================
// CONFIGURAÇÕES
// ============================================================

const TENANT_ID = process.env.TENANT_ID;
const CLIENT_ID = process.env.CLIENT_ID;
const CLIENT_SECRET = process.env.CLIENT_SECRET;

const SHAREPOINT_HOSTNAME =
    process.env.SHAREPOINT_HOSTNAME;

const SHAREPOINT_SITE_PATH =
    process.env.SHAREPOINT_SITE_PATH;

const LIBRARY_NAME =
    process.env.LIBRARY_NAME;

const FOLDER_PATH =
    process.env.FOLDER_PATH || 'Laudos';

const LIST_NAME =
    process.env.LIST_NAME;

// ============================================================
// INICIALIZAÇÃO
// ============================================================

console.log(
    '🚀 API SharePoint Global Plastic a iniciar...'
);

// ============================================================
// VALIDAÇÃO DAS VARIÁVEIS
// ============================================================

function validateEnvironment() {

    const requiredVariables = [
        'TENANT_ID',
        'CLIENT_ID',
        'CLIENT_SECRET',
        'SHAREPOINT_HOSTNAME',
        'SHAREPOINT_SITE_PATH',
        'LIBRARY_NAME',
        'LIST_NAME'
    ];

    const missing = requiredVariables.filter(
        variable => !process.env[variable]
    );

    if (missing.length > 0) {

        console.warn(
            '⚠️ Variáveis de ambiente não configuradas:',
            missing.join(', ')
        );
    }
}

validateEnvironment();

// ============================================================
// AUTENTICAÇÃO MICROSOFT
// ============================================================

async function getAccessToken() {

    try {

        const tokenUrl =
            `https://login.microsoftonline.com/` +
            `${TENANT_ID}/oauth2/v2.0/token`;

        const params =
            new URLSearchParams();

        params.append(
            'client_id',
            CLIENT_ID
        );

        params.append(
            'client_secret',
            CLIENT_SECRET
        );

        params.append(
            'scope',
            'https://graph.microsoft.com/.default'
        );

        params.append(
            'grant_type',
            'client_credentials'
        );

        const response =
            await axios.post(
                tokenUrl,
                params.toString(),
                {
                    headers: {
                        'Content-Type':
                            'application/x-www-form-urlencoded'
                    }
                }
            );

        return response.data.access_token;

    } catch (error) {

        const message =
            error.response?.data?.error_description ||
            error.response?.data?.error?.message ||
            error.message;

        console.error(
            '❌ Erro autenticação:',
            message
        );

        throw new Error(
            `Erro na autenticação: ${message}`
        );
    }
}

// ============================================================
// SITE ID
// ============================================================

async function getSiteId(accessToken) {

    try {

        const sitePath =
            SHAREPOINT_SITE_PATH.startsWith('/')
                ? SHAREPOINT_SITE_PATH
                : `/${SHAREPOINT_SITE_PATH}`;

        const url =
            `https://graph.microsoft.com/v1.0/sites/` +
            `${SHAREPOINT_HOSTNAME}:${sitePath}`;

        const response =
            await axios.get(
                url,
                {
                    headers: {
                        Authorization:
                            `Bearer ${accessToken}`
                    }
                }
            );

        return response.data.id;

    } catch (error) {

        const message =
            error.response?.data?.error?.message ||
            error.message;

        console.error(
            '❌ Erro ao localizar site:',
            message
        );

        throw new Error(
            `Erro ao localizar site: ${message}`
        );
    }
}

// ============================================================
// DRIVE / BIBLIOTECA
// ============================================================

async function getDriveId(
    accessToken,
    siteId = null
) {

    try {

        if (!siteId) {

            siteId =
                await getSiteId(
                    accessToken
                );
        }

        const url =
            `https://graph.microsoft.com/v1.0/` +
            `sites/${siteId}/drives`;

        const response =
            await axios.get(
                url,
                {
                    headers: {
                        Authorization:
                            `Bearer ${accessToken}`
                    }
                }
            );

        const drives =
            response.data.value || [];

        const drive =
            drives.find(
                item =>
                    item.name
                        ?.trim()
                        .toLowerCase() ===
                    LIBRARY_NAME
                        ?.trim()
                        .toLowerCase()
            );

        if (!drive) {

            const available =
                drives
                    .map(item => item.name)
                    .join(', ');

            throw new Error(
                `Biblioteca "${LIBRARY_NAME}" ` +
                `não encontrada. ` +
                `Bibliotecas disponíveis: ${available}`
            );
        }

        return drive.id;

    } catch (error) {

        const message =
            error.response?.data?.error?.message ||
            error.message;

        throw new Error(
            `Erro ao localizar biblioteca: ${message}`
        );
    }
}

// ============================================================
// LISTA SHAREPOINT
// ============================================================

async function getListId(
    accessToken,
    siteId
) {

    try {

        const url =
            `https://graph.microsoft.com/v1.0/` +
            `sites/${siteId}/lists`;

        const response =
            await axios.get(
                url,
                {
                    headers: {
                        Authorization:
                            `Bearer ${accessToken}`
                    }
                }
            );

        const lists =
            response.data.value || [];

        const list =
            lists.find(
                item =>
                    item.displayName
                        ?.trim()
                        .toLowerCase() ===
                    LIST_NAME
                        ?.trim()
                        .toLowerCase()
            );

        if (!list) {

            const available =
                lists
                    .map(item => item.displayName)
                    .join(', ');

            throw new Error(
                `Lista "${LIST_NAME}" ` +
                `não encontrada. ` +
                `Listas disponíveis: ${available}`
            );
        }

        return list.id;

    } catch (error) {

        const message =
            error.response?.data?.error?.message ||
            error.message;

        throw new Error(
            `Erro ao localizar lista: ${message}`
        );
    }
}

// ============================================================
// NORMALIZA TICKET
// ============================================================

function normalizeTicket(ticket) {

    if (!ticket) {
        return '';
    }

    return String(ticket)
        .trim()
        .toUpperCase()
        .replace(
            /[^A-Z0-9-]/g,
            ''
        );
}

// ============================================================
// IDENTIFICA O TICKET PELO NOME DO PDF
// ============================================================
//
// Reconhece:
//
// Laudo - SR-7382-17913948365.pdf
// Laudo - SR-7382-179134602920.pdf
// Laudo - SR-7382-20261007 1426.pdf
// Laudo - 12345-20261007 1426.pdf
//
// Retorna:
//
// SR-7382
// 12345
//
// ============================================================

function extractTicketNumber(fileName) {

    if (!fileName) {
        return null;
    }

    const name =
        String(fileName).trim();

    if (
        !name
            .toLowerCase()
            .endsWith('.pdf')
    ) {
        return null;
    }

    // --------------------------------------------
    // SR-7382
    // --------------------------------------------

    let match =
        name.match(
            /^Laudo\s*-\s*(SR-\d+)-.+\.pdf$/i
        );

    if (match) {

        return normalizeTicket(
            match[1]
        );
    }

    // --------------------------------------------
    // Outros prefixos:
    // OS-123
    // TK-123
    // ABC-123
    // --------------------------------------------

    match =
        name.match(
            /^Laudo\s*-\s*([A-Za-z]+-\d+)-.+\.pdf$/i
        );

    if (match) {

        return normalizeTicket(
            match[1]
        );
    }

    // --------------------------------------------
    // Ticket somente numérico
    // --------------------------------------------

    match =
        name.match(
            /^Laudo\s*-\s*(\d+)-.+\.pdf$/i
        );

    if (match) {

        return normalizeTicket(
            match[1]
        );
    }

    return null;
}

// ============================================================
// LISTA TODOS OS ARQUIVOS DA PASTA
// ============================================================

async function getAllFilesFromFolder(
    accessToken,
    driveId
) {

    const files = [];

    const cleanFolderPath =
        FOLDER_PATH
            .replace(/^\/+/, '')
            .replace(/\/+$/, '');

    const encodedPath =
        cleanFolderPath
            .split('/')
            .map(
                part =>
                    encodeURIComponent(part)
            )
            .join('/');

    let url =
        `https://graph.microsoft.com/v1.0/` +
        `drives/${driveId}/root:/` +
        `${encodedPath}:/children` +
        `?$top=200&` +
        `$select=id,name,file,folder,` +
        `createdDateTime,lastModifiedDateTime,size`;

    while (url) {

        const response =
            await axios.get(
                url,
                {
                    headers: {
                        Authorization:
                            `Bearer ${accessToken}`
                    }
                }
            );

        if (
            response.data &&
            Array.isArray(
                response.data.value
            )
        ) {

            files.push(
                ...response.data.value
            );
        }

        url =
            response.data[
                '@odata.nextLink'
            ] || null;
    }

    return files;
}

// ============================================================
// VERIFICA SE O PDF DO TICKET EXISTE
// ============================================================

async function ticketPdfExists(
    accessToken,
    driveId,
    ticketNumber
) {

    const ticket =
        normalizeTicket(
            ticketNumber
        );

    const files =
        await getAllFilesFromFolder(
            accessToken,
            driveId
        );

    return files.some(
        file => {

            if (!file.file) {
                return false;
            }

            const fileTicket =
                extractTicketNumber(
                    file.name
                );

            return (
                fileTicket === ticket
            );
        }
    );
}

// ============================================================
// OBTÉM VALOR DO TICKET DE UM ITEM DA LISTA
// ============================================================

function getTicketFromListFields(fields) {

    if (!fields) {
        return '';
    }

    const possibleFields = [
        fields['N_x00b0__x0020_do_x0020_ticket'],
        fields['N_x00b0_do_x0020_ticket'],
        fields['N_x00b0__x0020_do_x0020_Ticket'],
        fields['NumeroTicket'],
        fields['Ticket'],
        fields['Title']
    ];

    for (const value of possibleFields) {

        if (
            value !== undefined &&
            value !== null &&
            String(value).trim() !== ''
        ) {

            return normalizeTicket(
                value
            );
        }
    }

    return '';
}

// ============================================================
// VERIFICA SE TICKET EXISTE NA LISTA
// ============================================================

async function ticketExistsInList(
    accessToken,
    siteId,
    listId,
    ticketNumber
) {

    const ticket =
        normalizeTicket(
            ticketNumber
        );

    let url =
        `https://graph.microsoft.com/v1.0/` +
        `sites/${siteId}/lists/` +
        `${listId}/items` +
        `?$expand=fields&$top=200`;

    while (url) {

        const response =
            await axios.get(
                url,
                {
                    headers: {
                        Authorization:
                            `Bearer ${accessToken}`
                    }
                }
            );

        const items =
            response.data.value || [];

        const found =
            items.some(
                item =>
                    getTicketFromListFields(
                        item.fields
                    ) === ticket
            );

        if (found) {
            return true;
        }

        url =
            response.data[
                '@odata.nextLink'
            ] || null;
    }

    return false;
}

// ============================================================
// HEALTH CHECK
// ============================================================

app.get(
    '/',
    (req, res) => {

        res.json({
            status: 'online',
            timestamp:
                new Date()
                    .toISOString()
        });
    }
);

// ============================================================
// CHECK STATUS
// ============================================================

app.get(
    '/check-status/:ticketNumber',
    async (req, res) => {

        try {

            const ticketNumber =
                normalizeTicket(
                    req.params.ticketNumber
                );

            if (!ticketNumber) {

                return res
                    .status(400)
                    .json({
                        success: false,
                        error:
                            'Número do ticket não informado.'
                    });
            }

            console.log(
                `🔎 Verificando ticket ${ticketNumber}...`
            );

            const accessToken =
                await getAccessToken();

            const siteId =
                await getSiteId(
                    accessToken
                );

            const driveId =
                await getDriveId(
                    accessToken,
                    siteId
                );

            const listId =
                await getListId(
                    accessToken,
                    siteId
                );

            const [
                existsInPdf,
                existsInList
            ] =
                await Promise.all([
                    ticketPdfExists(
                        accessToken,
                        driveId,
                        ticketNumber
                    ),

                    ticketExistsInList(
                        accessToken,
                        siteId,
                        listId,
                        ticketNumber
                    )
                ]);

            console.log(
                `🔎 ${ticketNumber} | ` +
                `PDF: ${existsInPdf} | ` +
                `Lista: ${existsInList}`
            );

            return res.json({
                success: true,
                ticketNumber,
                existsInPdf,
                existsInList
            });

        } catch (error) {

            console.error(
                '❌ Erro check-status:',
                error.message
            );

            return res
                .status(500)
                .json({
                    success: false,
                    error:
                        error.message
                });
        }
    }
);

// ============================================================
// UPLOAD PDF
// ============================================================

app.post(
    '/upload-pdf',
    async (req, res) => {

        try {

            const {
                fileName,
                fileBase64,
                ticketNumber,
                ticketTitle,
                isReport
            } = req.body;

            if (
                !fileName ||
                !fileBase64
            ) {

                return res
                    .status(400)
                    .json({
                        success: false,
                        error:
                            'fileName e fileBase64 são obrigatórios.'
                    });
            }

            console.log(
                `📄 Upload PDF: ${fileName}`
            );

            if (ticketNumber) {

                console.log(
                    `🎫 Ticket: ${ticketNumber}`
                );
            }

            if (ticketTitle) {

                console.log(
                    `📝 Título: ${ticketTitle}`
                );
            }

            if (isReport) {

                console.log(
                    '📊 Arquivo identificado como relatório.'
                );
            }

            const accessToken =
                await getAccessToken();

            const siteId =
                await getSiteId(
                    accessToken
                );

            const driveId =
                await getDriveId(
                    accessToken,
                    siteId
                );

            const cleanFolderPath =
                FOLDER_PATH
                    .replace(/^\/+/, '')
                    .replace(/\/+$/, '');

            const encodedPath =
                cleanFolderPath
                    .split('/')
                    .map(
                        part =>
                            encodeURIComponent(part)
                    )
                    .join('/');

            const encodedFileName =
                encodeURIComponent(
                    fileName
                );

            const uploadUrl =
                `https://graph.microsoft.com/v1.0/` +
                `drives/${driveId}/root:/` +
                `${encodedPath}/` +
                `${encodedFileName}:/content`;

            const buffer =
                Buffer.from(
                    fileBase64,
                    'base64'
                );

            const response =
                await axios.put(
                    uploadUrl,
                    buffer,
                    {
                        headers: {
                            Authorization:
                                `Bearer ${accessToken}`,
                            'Content-Type':
                                'application/pdf'
                        },

                        maxBodyLength:
                            Infinity,

                        maxContentLength:
                            Infinity
                    }
                );

            console.log(
                `✅ PDF enviado: ${fileName}`
            );

            return res.json({
                success: true,

                message:
                    'PDF enviado com sucesso.',

                file: {
                    id:
                        response.data.id,

                    name:
                        response.data.name,

                    webUrl:
                        response.data.webUrl
                }
            });

        } catch (error) {

            const message =
                error.response?.data?.error?.message ||
                error.message;

            console.error(
                '❌ Erro PDF:',
                message
            );

            return res
                .status(500)
                .json({
                    success: false,
                    error: message
                });
        }
    }
);

// ============================================================
// UPLOAD DOS DADOS PARA A LISTA
// ============================================================
//
// OBSERVAÇÃO:
//
// Os nomes internos das colunas de uma lista SharePoint podem
// ser diferentes dos nomes exibidos na tela.
//
// Se sua API anterior já possuía o mapeamento correto das
// colunas, preserve os nomes internos que já funcionavam.
//
// ============================================================

app.post(
    '/upload-list-data',
    async (req, res) => {

        try {

            const {
                ticketNumber,
                listData
            } = req.body;

            if (
                !Array.isArray(
                    listData
                ) ||
                listData.length === 0
            ) {

                return res
                    .status(400)
                    .json({
                        success: false,
                        error:
                            'listData não informado.'
                    });
            }

            const accessToken =
                await getAccessToken();

            const siteId =
                await getSiteId(
                    accessToken
                );

            const listId =
                await getListId(
                    accessToken,
                    siteId
                );

            let inserted = 0;

            for (
                const row
                of listData
            ) {

                /*
                 * ATENÇÃO:
                 *
                 * Se esses nomes internos forem diferentes
                 * na sua lista atual, mantenha os nomes que
                 * seu server.js antigo já utilizava.
                 */

                const fields = {
                    Title:
                        String(
                            row[
                                'N° do ticket'
                            ] ||
                            ticketNumber ||
                            ''
                        ),

                    NumeroTicket:
                        String(
                            row[
                                'N° do ticket'
                            ] ||
                            ticketNumber ||
                            ''
                        ),

                    NomeCliente:
                        row[
                            'Nome do Cliente'
                        ] || '',

                    Item:
                        String(
                            row.Item || ''
                        ),

                    Qtde:
                        row.Qtde || 0,

                    Motivo:
                        Array.isArray(
                            row.Motivo
                        )
                            ? row.Motivo.join(
                                ', '
                            )
                            : (
                                row.Motivo ||
                                ''
                            ),

                    OrigemDefeito:
                        row[
                            'Origem do defeito'
                        ] || '',

                    Disposicao:
                        row[
                            'Disposição'
                        ] || '',

                    DisposicaoPecas:
                        row[
                            'Disposição das peças'
                        ] || '',

                    DataGeracao:
                        row[
                            'Data de Geração'
                        ] || ''
                };

                for (
                    let i = 1;
                    i <= 10;
                    i++
                ) {

                    const value =
                        row[
                            `Foto ${i}`
                        ];

                    if (value) {

                        fields[
                            `Foto${i}`
                        ] = value;
                    }
                }

                const url =
                    `https://graph.microsoft.com/v1.0/` +
                    `sites/${siteId}/lists/` +
                    `${listId}/items`;

                await axios.post(
                    url,
                    {
                        fields
                    },
                    {
                        headers: {
                            Authorization:
                                `Bearer ${accessToken}`,

                            'Content-Type':
                                'application/json'
                        }
                    }
                );

                inserted++;
            }

            console.log(
                `✅ Lista: ${ticketNumber} | ` +
                `${inserted} linha(s)`
            );

            return res.json({
                success: true,
                inserted
            });

        } catch (error) {

            const message =
                error.response?.data?.error?.message ||
                error.message;

            console.error(
                '❌ Erro lista:',
                message
            );

            return res
                .status(500)
                .json({
                    success: false,
                    error: message
                });
        }
    }
);

// ============================================================
// EXCLUI TODOS OS PDFs DE UM TICKET
// ============================================================

app.delete(
    '/delete-pdf-by-ticket-number/:ticketNumber',
    async (req, res) => {

        try {

            const ticketNumber =
                normalizeTicket(
                    req.params.ticketNumber
                );

            if (!ticketNumber) {

                return res
                    .status(400)
                    .json({
                        success: false,
                        error:
                            'Número do ticket não informado.'
                    });
            }

            const accessToken =
                await getAccessToken();

            const siteId =
                await getSiteId(
                    accessToken
                );

            const driveId =
                await getDriveId(
                    accessToken,
                    siteId
                );

            const files =
                await getAllFilesFromFolder(
                    accessToken,
                    driveId
                );

            const matchingFiles =
                files.filter(
                    file =>
                        file.file &&
                        extractTicketNumber(
                            file.name
                        ) ===
                        ticketNumber
                );

            let deleted = 0;

            for (
                const file
                of matchingFiles
            ) {

                const deleteUrl =
                    `https://graph.microsoft.com/v1.0/` +
                    `drives/${driveId}/items/` +
                    `${file.id}`;

                await axios.delete(
                    deleteUrl,
                    {
                        headers: {
                            Authorization:
                                `Bearer ${accessToken}`
                        }
                    }
                );

                deleted++;

                console.log(
                    `🗑️ PDF excluído: ${file.name}`
                );
            }

            return res.json({
                success: true,
                ticketNumber,
                deleted
            });

        } catch (error) {

            const message =
                error.response?.data?.error?.message ||
                error.message;

            console.error(
                '❌ Erro ao excluir PDF:',
                message
            );

            return res
                .status(500)
                .json({
                    success: false,
                    error: message
                });
        }
    }
);

// ============================================================
// LIMPEZA DE PDFs DUPLICADOS
// ============================================================
//
// EXEMPLO:
//
// Laudo - SR-7382-17913948365.pdf
// Laudo - SR-7382-179134602920.pdf
//
// Os dois são:
// SR-7382
//
// A rotina:
//
// 1. Agrupa por ticket.
// 2. Ordena por lastModifiedDateTime.
// 3. Mantém o mais recente.
// 4. Exclui os PDFs mais antigos.
//
// NÃO APAGA ITENS DA LISTA SHAREPOINT.
//
// ============================================================

app.post(
    '/cleanup-duplicate-pdfs',
    async (req, res) => {

        try {

            console.log(
                '🧹 Iniciando limpeza de PDFs duplicados...'
            );

            const accessToken =
                await getAccessToken();

            const siteId =
                await getSiteId(
                    accessToken
                );

            const driveId =
                await getDriveId(
                    accessToken,
                    siteId
                );

            const files =
                await getAllFilesFromFolder(
                    accessToken,
                    driveId
                );

            console.log(
                `📂 ${files.length} arquivo(s) encontrado(s).`
            );

            const ticketGroups =
                new Map();

            // ------------------------------------------------
            // AGRUPAMENTO
            // ------------------------------------------------

            for (
                const file
                of files
            ) {

                // Ignora pastas.

                if (!file.file) {
                    continue;
                }

                // Ignora arquivos que não são PDF.

                if (
                    !file.name ||
                    !file.name
                        .toLowerCase()
                        .endsWith('.pdf')
                ) {

                    continue;
                }

                const ticketNumber =
                    extractTicketNumber(
                        file.name
                    );

                /*
                 * Se não reconheceu o padrão do laudo,
                 * não mexe no arquivo.
                 */

                if (!ticketNumber) {

                    console.log(
                        `ℹ️ Arquivo ignorado: ${file.name}`
                    );

                    continue;
                }

                if (
                    !ticketGroups.has(
                        ticketNumber
                    )
                ) {

                    ticketGroups.set(
                        ticketNumber,
                        []
                    );
                }

                ticketGroups
                    .get(ticketNumber)
                    .push(file);
            }

            const duplicates = [];
            const deletedFiles = [];
            const keptFiles = [];

            // ------------------------------------------------
            // PROCESSAMENTO POR TICKET
            // ------------------------------------------------

            for (
                const [
                    ticketNumber,
                    ticketFiles
                ]
                of ticketGroups.entries()
            ) {

                // Um único PDF: não há duplicidade.

                if (
                    ticketFiles.length <= 1
                ) {

                    continue;
                }

                console.log(
                    `⚠️ Ticket ${ticketNumber}: ` +
                    `${ticketFiles.length} PDFs encontrados.`
                );

                /*
                 * Mais recentemente modificado primeiro.
                 */

                ticketFiles.sort(
                    (a, b) => {

                        const dateA =
                            new Date(
                                a.lastModifiedDateTime ||
                                a.createdDateTime ||
                                0
                            ).getTime();

                        const dateB =
                            new Date(
                                b.lastModifiedDateTime ||
                                b.createdDateTime ||
                                0
                            ).getTime();

                        return (
                            dateB -
                            dateA
                        );
                    }
                );

                const keepFile =
                    ticketFiles[0];

                const filesToDelete =
                    ticketFiles.slice(1);

                console.log(
                    `✅ Ticket ${ticketNumber}: ` +
                    `mantendo "${keepFile.name}"`
                );

                keptFiles.push({
                    ticketNumber,

                    fileName:
                        keepFile.name,

                    modified:
                        keepFile.lastModifiedDateTime
                });

                duplicates.push({
                    ticketNumber,

                    total:
                        ticketFiles.length,

                    keep:
                        keepFile.name,

                    delete:
                        filesToDelete.map(
                            file =>
                                file.name
                        )
                });

                // --------------------------------------------
                // EXCLUSÃO DOS ANTIGOS
                // --------------------------------------------

                for (
                    const file
                    of filesToDelete
                ) {

                    const deleteUrl =
                        `https://graph.microsoft.com/v1.0/` +
                        `drives/${driveId}/items/` +
                        `${file.id}`;

                    await axios.delete(
                        deleteUrl,
                        {
                            headers: {
                                Authorization:
                                    `Bearer ${accessToken}`
                            }
                        }
                    );

                    console.log(
                        `🗑️ Duplicado excluído: ${file.name}`
                    );

                    deletedFiles.push({
                        ticketNumber,

                        fileName:
                            file.name
                    });
                }
            }

            console.log(
                '✅ Limpeza de PDFs concluída.'
            );

            console.log(
                `🗑️ ${deletedFiles.length} ` +
                `PDF(s) duplicado(s) removido(s).`
            );

            return res.json({

                success: true,

                message:
                    `${deletedFiles.length} ` +
                    `PDF(s) duplicado(s) removido(s).`,

                totalFilesChecked:
                    files.length,

                ticketsChecked:
                    ticketGroups.size,

                ticketsWithDuplicates:
                    duplicates.length,

                deletedCount:
                    deletedFiles.length,

                duplicates,

                keptFiles,

                deletedFiles
            });

        } catch (error) {

            const message =
                error.response?.data?.error?.message ||
                error.message;

            console.error(
                '❌ Erro ao limpar PDFs duplicados:',
                message
            );

            return res
                .status(500)
                .json({
                    success: false,
                    error: message
                });
        }
    }
);

// ============================================================
// LIMPAR TODA A LISTA SHAREPOINT
// ============================================================
//
// ATENÇÃO:
//
// Esta rota é separada.
//
// /cleanup-duplicate-pdfs
// NÃO chama esta rota.
//
// Portanto, limpar PDFs duplicados NÃO limpa a lista.
//
// ============================================================

app.delete(
    '/clear-list',
    async (req, res) => {

        try {

            console.log(
                '⚠️ Iniciando limpeza total da lista...'
            );

            const accessToken =
                await getAccessToken();

            const siteId =
                await getSiteId(
                    accessToken
                );

            const listId =
                await getListId(
                    accessToken,
                    siteId
                );

            let deleted = 0;

            /*
             * É melhor buscar novamente a primeira página
             * depois das exclusões, pois os itens da coleção
             * estão sendo modificados durante o processo.
             */

            while (true) {

                const url =
                    `https://graph.microsoft.com/v1.0/` +
                    `sites/${siteId}/lists/` +
                    `${listId}/items?$top=200`;

                const response =
                    await axios.get(
                        url,
                        {
                            headers: {
                                Authorization:
                                    `Bearer ${accessToken}`
                            }
                        }
                    );

                const items =
                    response.data.value || [];

                if (
                    items.length === 0
                ) {

                    break;
                }

                for (
                    const item
                    of items
                ) {

                    const deleteUrl =
                        `https://graph.microsoft.com/v1.0/` +
                        `sites/${siteId}/lists/` +
                        `${listId}/items/` +
                        `${item.id}`;

                    await axios.delete(
                        deleteUrl,
                        {
                            headers: {
                                Authorization:
                                    `Bearer ${accessToken}`
                            }
                        }
                    );

                    deleted++;
                }

                console.log(
                    `🗑️ ${deleted} registro(s) removido(s) até agora...`
                );
            }

            console.log(
                `✅ ${deleted} registro(s) removido(s) da lista.`
            );

            return res.json({
                success: true,

                deleted,

                message:
                    `${deleted} registro(s) ` +
                    `foram removidos da lista.`
            });

        } catch (error) {

            const message =
                error.response?.data?.error?.message ||
                error.message;

            console.error(
                '❌ Erro ao limpar lista:',
                message
            );

            return res
                .status(500)
                .json({
                    success: false,
                    error: message
                });
        }
    }
);

// ============================================================
// ROTA NÃO ENCONTRADA
// ============================================================

app.use(
    (req, res) => {

        res.status(404).json({
            success: false,

            error:
                `Rota não encontrada: ` +
                `${req.method} ${req.originalUrl}`
        });
    }
);

// ============================================================
// INICIA SERVIDOR
// ============================================================

app.listen(
    PORT,
    '0.0.0.0',
    () => {

        console.log(
            `🌐 API online na porta ${PORT}`
        );

        console.log(
            `📂 Pasta de laudos: ${FOLDER_PATH}`
        );

        console.log(
            `📚 Biblioteca: ${LIBRARY_NAME}`
        );

        console.log(
            `📋 Lista: ${LIST_NAME}`
        );
    }
);
