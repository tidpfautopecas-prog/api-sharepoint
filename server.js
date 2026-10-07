require('dotenv').config();

const express = require('express');
const cors = require('cors');
const axios = require('axios');

const app = express();

const PORT = process.env.PORT || 3000;

app.use(cors());

app.use(express.json({
    limit: '50mb'
}));

app.use(express.urlencoded({
    extended: true,
    limit: '50mb'
}));

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

        const url =
            `https://graph.microsoft.com/v1.0/sites/` +
            `${SHAREPOINT_HOSTNAME}:` +
            `${SHAREPOINT_SITE_PATH}`;

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
                        .trim()
                        .toLowerCase() ===
                    LIBRARY_NAME
                        .trim()
                        .toLowerCase()
            );

        if (!drive) {

            throw new Error(
                `Biblioteca "${LIBRARY_NAME}" ` +
                `não encontrada.`
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
                        .trim()
                        .toLowerCase() ===
                    LIST_NAME
                        .trim()
                        .toLowerCase()
            );

        if (!list) {

            throw new Error(
                `Lista "${LIST_NAME}" ` +
                `não encontrada.`
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
// IDENTIFICA TICKET PELO NOME DO PDF
// ============================================================
//
// Exemplos reconhecidos:
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

function extractTicketNumber(
    fileName
) {

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

    /*
     * Ticket padrão SR-9999
     */

    let match =
        name.match(
            /^Laudo\s*-\s*(SR-\d+)-.+\.pdf$/i
        );

    if (match) {

        return normalizeTicket(
            match[1]
        );
    }

    /*
     * Outros prefixos:
     *
     * OS-123
     * TK-123
     * ABC-123
     */

    match =
        name.match(
            /^Laudo\s*-\s*([A-Za-z]+-\d+)-.+\.pdf$/i
        );

    if (match) {

        return normalizeTicket(
            match[1]
        );
    }

    /*
     * Ticket somente numérico
     */

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
// LISTAR TODOS OS ARQUIVOS DA PASTA
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
// VERIFICA SE PDF EXISTE
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
// VERIFICA SE TICKET EXISTE NA LISTA
// ============================================================

async function ticketExistsInList(
    accessToken,
    siteId,
    listId,
    ticketNumber
) {

    const ticket =
        String(
            ticketNumber
        ).trim();

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
                item => {

                    const fields =
                        item.fields || {};

                    const possibleValues = [
                        fields['N_x00b0__x0020_do_x0020_ticket'],
                        fields['N_x00b0_do_x0020_ticket'],
                        fields['NumeroTicket'],
                        fields['Ticket'],
                        fields['Title']
                    ];

                    return possibleValues
                        .filter(
                            value =>
                                value !== undefined &&
                                value !== null
                        )
                        .some(
                            value =>
                                String(value)
                                    .trim()
                                    .toUpperCase() ===
                                ticket
                                    .toUpperCase()
                        );
                }
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
// IMPORTANTE:
//
// Mantém somente UM PDF por ticket.
//
// Exemplo:
//
// Laudo - SR-7382-17913948365.pdf
// Laudo - SR-7382-179134602920.pdf
//
// Ambos são identificados como:
//
// SR-7382
//
// O arquivo mais recentemente modificado no SharePoint é mantido.
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

            for (
                const file
                of files
            ) {

                /*
                 * Ignora pastas
                 */

                if (!file.file) {
                    continue;
                }

                /*
                 * Ignora arquivos não PDF
                 */

                if (
                    !file.name ||
                    !file.name
                        .toLowerCase()
                        .endsWith('.pdf')
                ) {
                    continue;
                }

                /*
                 * Identifica ticket
                 */

                const ticketNumber =
                    extractTicketNumber(
                        file.name
                    );

                /*
                 * Não reconheceu como laudo.
                 * Não mexe no arquivo.
                 */

                if (!ticketNumber) {

                    console.log(
                        `ℹ️ Ignorado: ${file.name}`
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

            /*
             * Percorre cada ticket
             */

            for (
                const [
                    ticketNumber,
                    ticketFiles
                ]
                of ticketGroups.entries()
            ) {

                /*
                 * Apenas um PDF.
                 * Não existe duplicidade.
                 */

                if (
                    ticketFiles.length <= 1
                ) {
                    continue;
                }

                console.log(
                    `⚠️ ${ticketNumber}: ` +
                    `${ticketFiles.length} PDFs encontrados.`
                );

                /*
                 * Ordena pelo lastModifiedDateTime.
                 *
                 * Mais recente fica na posição 0.
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

                /*
                 * Mantém o mais recente
                 */

                const keepFile =
                    ticketFiles[0];

                /*
                 * Todos os outros
                 * são duplicados
                 */

                const filesToDelete =
                    ticketFiles.slice(1);

                console.log(
                    `✅ ${ticketNumber}: ` +
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


                /*
                 * Exclui PDFs antigos
                 */

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
                        `🗑️ Excluído: ${file.name}`
                    );

                    deletedFiles.push({
                        ticketNumber,
                        fileName:
                            file.name
                    });
                }
            }


            console.log(
                '✅ Limpeza concluída.'
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
// LIMPAR LISTA SHAREPOINT
// ============================================================
//
// ATENÇÃO:
// Essa rota continua separada.
// Ela NÃO é chamada pela limpeza de PDFs duplicados.
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

            let url =
                `https://graph.microsoft.com/v1.0/` +
                `sites/${siteId}/lists/` +
                `${listId}/items?$top=200`;

            let deleted = 0;

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

                /*
                 * Guarda nextLink ANTES
                 * de começar a excluir.
                 */

                const nextLink =
                    response.data[
                        '@odata.nextLink'
                    ] || null;

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

                url =
                    nextLink;
            }

            console.log(
                `✅ ${deleted} registro(s) removido(s).`
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
// TRATAMENTO DE ROTA NÃO ENCONTRADA
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
