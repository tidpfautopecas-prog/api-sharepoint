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

const SITE_ID = process.env.SITE_ID;
const LIST_ID = process.env.LIST_ID;
const DRIVE_ID = process.env.DRIVE_ID;

const LIBRARY_NAME = process.env.LIBRARY_NAME;

const FOLDER_PATH =
    process.env.FOLDER_PATH || 'Laudos';

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
        'SITE_ID',
        'LIST_ID'
    ];

    const missing =
        requiredVariables.filter(
            variable => !process.env[variable]
        );

    if (missing.length > 0) {

        console.warn(
            '⚠️ Variáveis de ambiente não configuradas:',
            missing.join(', ')
        );
    }

    if (!DRIVE_ID && !LIBRARY_NAME) {

        console.warn(
            '⚠️ Configure DRIVE_ID ou LIBRARY_NAME no Render.'
        );
    }
}

validateEnvironment();

// ============================================================
// TRATAMENTO DE ERRO
// ============================================================

function getErrorMessage(error) {

    return (
        error?.response?.data?.error_description ||
        error?.response?.data?.error?.message ||
        error?.response?.data?.message ||
        error?.message ||
        'Erro desconhecido.'
    );
}

// ============================================================
// AUTENTICAÇÃO MICROSOFT
// ============================================================

async function getAccessToken() {

    try {

        if (!TENANT_ID) {

            throw new Error(
                'TENANT_ID não configurado no Render.'
            );
        }

        if (!CLIENT_ID) {

            throw new Error(
                'CLIENT_ID não configurado no Render.'
            );
        }

        if (!CLIENT_SECRET) {

            throw new Error(
                'CLIENT_SECRET não configurado no Render.'
            );
        }

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

        if (!response.data.access_token) {

            throw new Error(
                'Microsoft não retornou access_token.'
            );
        }

        return response.data.access_token;

    } catch (error) {

        const message =
            getErrorMessage(error);

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

async function getSiteId() {

    if (!SITE_ID) {

        throw new Error(
            'Variável SITE_ID não configurada no Render.'
        );
    }

    return SITE_ID;
}

// ============================================================
// LIST ID
// ============================================================

async function getListId() {

    if (!LIST_ID) {

        throw new Error(
            'Variável LIST_ID não configurada no Render.'
        );
    }

    return LIST_ID;
}

// ============================================================
// DRIVE / BIBLIOTECA
// ============================================================

async function getDriveId(
    accessToken,
    siteId = null
) {

    try {

        // Se já temos DRIVE_ID configurado,
        // não precisamos procurar a biblioteca.

        if (DRIVE_ID) {

            return DRIVE_ID;
        }

        if (!siteId) {

            siteId =
                await getSiteId();
        }

        if (!LIBRARY_NAME) {

            throw new Error(
                'DRIVE_ID não configurado e LIBRARY_NAME não informado.'
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
                    .map(
                        item => item.name
                    )
                    .join(', ');

            throw new Error(
                `Biblioteca "${LIBRARY_NAME}" não encontrada. ` +
                `Bibliotecas disponíveis: ${available}`
            );
        }

        return drive.id;

    } catch (error) {

        const message =
            getErrorMessage(error);

        throw new Error(
            `Erro ao localizar biblioteca: ${message}`
        );
    }
}

// ============================================================
// NORMALIZAÇÃO DO TICKET
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
// Exemplos:
//
// Laudo - SR-7382-17913948365.pdf
// Laudo - SR-7382-179134602920.pdf
// Laudo - SR-7382-20261007 1426.pdf
// Laudo - 12345-20261007 1426.pdf
//
// Resultado:
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

    // --------------------------------------------------------
    // SR-7382
    // --------------------------------------------------------

    let match =
        name.match(
            /^Laudo\s*-\s*(SR-\d+)-.+\.pdf$/i
        );

    if (match) {

        return normalizeTicket(
            match[1]
        );
    }

    // --------------------------------------------------------
    // Outros prefixos:
    //
    // OS-123
    // TK-123
    // ABC-123
    // --------------------------------------------------------

    match =
        name.match(
            /^Laudo\s*-\s*([A-Za-z]+-\d+)-.+\.pdf$/i
        );

    if (match) {

        return normalizeTicket(
            match[1]
        );
    }

    // --------------------------------------------------------
    // Ticket somente numérico
    // --------------------------------------------------------

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
// CAMINHO DA PASTA
// ============================================================

function getEncodedFolderPath() {

    const cleanFolderPath =
        String(FOLDER_PATH)
            .replace(/^\/+/, '')
            .replace(/\/+$/, '');

    return cleanFolderPath
        .split('/')
        .filter(Boolean)
        .map(
            part =>
                encodeURIComponent(part)
        )
        .join('/');
}

// ============================================================
// LISTA TODOS OS ARQUIVOS DA PASTA
// ============================================================

async function getAllFilesFromFolder(
    accessToken,
    driveId
) {

    const files = [];

    const encodedPath =
        getEncodedFolderPath();

    let url =
        `https://graph.microsoft.com/v1.0/` +
        `drives/${driveId}/root:/` +
        `${encodedPath}:/children` +
        `?$top=200&` +
        `$select=id,name,file,folder,` +
        `createdDateTime,lastModifiedDateTime,size,webUrl`;

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

        const currentFiles =
            response.data.value || [];

        files.push(
            ...currentFiles
        );

        url =
            response.data[
                '@odata.nextLink'
            ] || null;
    }

    return files;
}

// ============================================================
// VERIFICA SE PDF DO TICKET EXISTE
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
// LOCALIZA TICKET NOS CAMPOS DA LISTA
// ============================================================

function getTicketFromListFields(
    fields
) {

    if (!fields) {

        return '';
    }

    const possibleValues = [
        fields['N_x00b0_doticket'],
        fields['N_x00b0__x0020_do_x0020_ticket'],
        fields['N_x00b0_do_x0020_ticket'],
        fields['NumeroTicket'],
        fields['Ticket'],
        fields['Title']
    ];

    for (
        const value
        of possibleValues
    ) {

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
                    .toISOString(),

            configuration: {
                siteId:
                    Boolean(SITE_ID),

                listId:
                    Boolean(LIST_ID),

                driveId:
                    Boolean(DRIVE_ID),

                libraryName:
                    Boolean(LIBRARY_NAME),

                folderPath:
                    FOLDER_PATH
            }
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
                await getSiteId();

            const listId =
                await getListId();

            const driveId =
                await getDriveId(
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

            const message =
                getErrorMessage(error);

            console.error(
                '❌ Erro check-status:',
                message
            );

            return res
                .status(500)
                .json({
                    success: false,

                    error:
                        message
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

            const accessToken =
                await getAccessToken();

            const siteId =
                await getSiteId();

            const driveId =
                await getDriveId(
                    accessToken,
                    siteId
                );

            const encodedPath =
                getEncodedFolderPath();

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
                getErrorMessage(error);

            console.error(
                '❌ Erro PDF:',
                message
            );

            return res
                .status(500)
                .json({
                    success: false,

                    error:
                        message
                });
        }
    }
);

// ============================================================
// MAPEAMENTO DAS COLUNAS DA LISTA
// ============================================================
//
// Estes são os nomes internos utilizados no projeto.
//
// ============================================================

function buildSharePointFields(
    row,
    ticketNumber
) {

    const fields = {};

    const ticket =
        String(
            row['N° do ticket'] ||
            ticketNumber ||
            ''
        );

    // --------------------------------------------------------
    // Ticket
    // --------------------------------------------------------

    fields.Title =
        ticket;

    fields.N_x00b0_doticket =
        ticket;

    // --------------------------------------------------------
    // Cliente
    // --------------------------------------------------------

    fields.NomedoCliente =
        row['Nome do Cliente'] || '';

    // --------------------------------------------------------
    // Item
    // --------------------------------------------------------

    fields.Item =
        String(
            row.Item || ''
        );

    // --------------------------------------------------------
    // Quantidade
    // --------------------------------------------------------

    fields.Qtde =
        Number(
            row.Qtde || 0
        );

    // --------------------------------------------------------
    // Motivo
    // --------------------------------------------------------

    fields.Motivo =
        Array.isArray(
            row.Motivo
        )
            ? row.Motivo.join(', ')
            : (
                row.Motivo || ''
            );

    // --------------------------------------------------------
    // Origem do defeito
    // --------------------------------------------------------

    fields.Origemdodefeito =
        row['Origem do defeito'] || '';

    // --------------------------------------------------------
    // Disposição
    // --------------------------------------------------------

    fields[
        'Disposi_x00e7__x00e3_o'
    ] =
        row['Disposição'] || '';

    // --------------------------------------------------------
    // Disposição das peças
    // --------------------------------------------------------

    fields[
        'Disposi_x00e7__x00e3_odaspe_x00e'
    ] =
        row['Disposição das peças'] || '';

    // --------------------------------------------------------
    // Data
    // --------------------------------------------------------

    fields[
        'DatadeGera_x00e7__x00e3_o'
    ] =
        row['Data de Geração'] || '';

    // --------------------------------------------------------
    // Fotos
    // --------------------------------------------------------

    for (
        let i = 1;
        i <= 10;
        i++
    ) {

        const photo =
            row[`Foto ${i}`];

        if (photo) {

            fields[`Foto${i}`] =
                photo;
        }
    }

    return fields;
}

// ============================================================
// UPLOAD LIST DATA
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

            console.log(
                `📋 Enviando ticket ${ticketNumber} para lista...`
            );

            const accessToken =
                await getAccessToken();

            const siteId =
                await getSiteId();

            const listId =
                await getListId();

            let inserted = 0;

            for (
                const row
                of listData
            ) {

                const fields =
                    buildSharePointFields(
                        row,
                        ticketNumber
                    );

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

                ticketNumber,

                inserted
            });

        } catch (error) {

            const message =
                getErrorMessage(error);

            console.error(
                '❌ Erro lista:',
                message
            );

            return res
                .status(500)
                .json({
                    success: false,

                    error:
                        message
                });
        }
    }
);

// ============================================================
// EXCLUI PDFs DE UM TICKET
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
                await getSiteId();

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
                        ) === ticketNumber
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
                getErrorMessage(error);

            console.error(
                '❌ Erro ao excluir PDF:',
                message
            );

            return res
                .status(500)
                .json({
                    success: false,

                    error:
                        message
                });
        }
    }
);

// ============================================================
// LIMPEZA DE PDFs DUPLICADOS
// ============================================================
//
// IMPORTANTE:
//
// Esta rotina mexe SOMENTE nos PDFs.
//
// Ela NÃO exclui registros da Lista SharePoint.
//
// Para cada ticket:
//
// SR-7382
//
// se houver:
//
// Laudo - SR-7382-17913948365.pdf
// Laudo - SR-7382-179134602920.pdf
//
// mantém o arquivo mais recentemente modificado
// e exclui os demais.
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
                await getSiteId();

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

                // Não reconheceu como laudo.
                // Não toca no arquivo.

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
            // PROCESSAMENTO
            // ------------------------------------------------

            for (
                const [
                    ticketNumber,
                    ticketFiles
                ]
                of ticketGroups.entries()
            ) {

                if (
                    ticketFiles.length <= 1
                ) {

                    continue;
                }

                console.log(
                    `⚠️ Ticket ${ticketNumber}: ` +
                    `${ticketFiles.length} PDFs encontrados.`
                );

                // --------------------------------------------
                // Ordena pelo mais recente
                // --------------------------------------------

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

                // --------------------------------------------
                // Mantém o primeiro
                // --------------------------------------------

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
                // Exclui os antigos
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
                getErrorMessage(error);

            console.error(
                '❌ Erro ao limpar PDFs duplicados:',
                message
            );

            return res
                .status(500)
                .json({
                    success: false,

                    error:
                        message
                });
        }
    }
);

// ============================================================
// LIMPAR TODA A LISTA
// ============================================================
//
// ATENÇÃO:
//
// Essa rota apaga TODOS os itens da lista.
//
// Ela NÃO é executada pela rota:
//
// /cleanup-duplicate-pdfs
//
// ============================================================

app.delete(
    '/clear-list',
    async (req, res) => {

        try {

            console.log(
                '⚠️ Iniciando limpeza TOTAL da lista...'
            );

            const accessToken =
                await getAccessToken();

            const siteId =
                await getSiteId();

            const listId =
                await getListId();

            let deleted = 0;

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
                getErrorMessage(error);

            console.error(
                '❌ Erro ao limpar lista:',
                message
            );

            return res
                .status(500)
                .json({
                    success: false,

                    error:
                        message
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
            `🔗 SITE_ID configurado: ${Boolean(SITE_ID)}`
        );

        console.log(
            `📋 LIST_ID configurado: ${Boolean(LIST_ID)}`
        );

        console.log(
            `📚 DRIVE_ID configurado: ${Boolean(DRIVE_ID)}`
        );

        console.log(
            `📚 LIBRARY_NAME configurado: ${Boolean(LIBRARY_NAME)}`
        );
    }
);
