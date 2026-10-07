import 'dotenv/config';
import express from 'express';
import cors from 'cors';
import axios from 'axios';

const app = express();
const PORT = process.env.PORT || 3000;

// ============================================================
// CONFIGURAÇÃO
// ============================================================

const TENANT_ID = process.env.TENANT_ID;
const CLIENT_ID = process.env.CLIENT_ID;
const CLIENT_SECRET = process.env.CLIENT_SECRET;

const SITE_ID = process.env.SITE_ID;
const LIBRARY_NAME = process.env.LIBRARY_NAME;
const LIST_NAME = process.env.LIST_NAME;

const FOLDER_PATH =
    process.env.FOLDER_PATH || 'Laudos';

// ============================================================
// MIDDLEWARES
// ============================================================

app.use(cors());

app.use(
    express.json({
        limit: '100mb'
    })
);

app.use(
    express.urlencoded({
        extended: true,
        limit: '100mb'
    })
);

// ============================================================
// CACHE
// ============================================================

let cachedToken = null;
let cachedTokenExpiresAt = 0;

let cachedDriveId = null;
let cachedListId = null;
let cachedColumns = null;

// ============================================================
// INICIALIZAÇÃO
// ============================================================

console.log(
    '🚀 API SharePoint Global Plastic iniciando...'
);

// ============================================================
// VALIDAÇÃO DAS VARIÁVEIS
// ============================================================

function validateEnvironment() {

    const required = [
        'TENANT_ID',
        'CLIENT_ID',
        'CLIENT_SECRET',
        'SITE_ID',
        'LIBRARY_NAME',
        'LIST_NAME'
    ];

    const missing =
        required.filter(
            name =>
                !process.env[name]
        );

    if (missing.length > 0) {

        console.error(
            '❌ Variáveis de ambiente ausentes:',
            missing.join(', ')
        );

        return;
    }

    console.log(
        '✅ Variáveis de ambiente configuradas.'
    );
}

validateEnvironment();

// ============================================================
// ERROS
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
// NORMALIZA TEXTO
// ============================================================

function normalizeText(value) {

    return String(
        value ?? ''
    )
        .normalize('NFD')
        .replace(
            /[\u0300-\u036f]/g,
            ''
        )
        .trim()
        .toLowerCase();
}

// ============================================================
// NORMALIZA TICKET
// ============================================================

function normalizeTicket(value) {

    return String(
        value ?? ''
    )
        .trim()
        .toUpperCase()
        .replace(
            /^#/,
            ''
        )
        .replace(
            /[^A-Z0-9-]/g,
            ''
        );
}

// ============================================================
// TOKEN MICROSOFT
// ============================================================

async function getAccessToken() {

    if (
        cachedToken &&
        Date.now() <
            cachedTokenExpiresAt
    ) {

        return cachedToken;
    }

    if (!TENANT_ID) {

        throw new Error(
            'TENANT_ID não configurado.'
        );
    }

    if (!CLIENT_ID) {

        throw new Error(
            'CLIENT_ID não configurado.'
        );
    }

    if (!CLIENT_SECRET) {

        throw new Error(
            'CLIENT_SECRET não configurado.'
        );
    }

    const url =
        `https://login.microsoftonline.com/` +
        `${TENANT_ID}/oauth2/v2.0/token`;

    const body =
        new URLSearchParams();

    body.append(
        'client_id',
        CLIENT_ID
    );

    body.append(
        'client_secret',
        CLIENT_SECRET
    );

    body.append(
        'scope',
        'https://graph.microsoft.com/.default'
    );

    body.append(
        'grant_type',
        'client_credentials'
    );

    try {

        const response =
            await axios.post(
                url,
                body.toString(),
                {
                    headers: {
                        'Content-Type':
                            'application/x-www-form-urlencoded'
                    }
                }
            );

        if (
            !response.data.access_token
        ) {

            throw new Error(
                'Microsoft não retornou access_token.'
            );
        }

        cachedToken =
            response.data.access_token;

        const expires =
            Number(
                response.data.expires_in ||
                3600
            );

        cachedTokenExpiresAt =
            Date.now() +
            Math.max(
                expires - 300,
                60
            ) * 1000;

        return cachedToken;

    } catch (error) {

        throw new Error(
            `Erro na autenticação Microsoft: ` +
            `${getErrorMessage(error)}`
        );
    }
}

// ============================================================
// LOCALIZA BIBLIOTECA
// ============================================================

async function getDriveId(token) {

    if (cachedDriveId) {

        return cachedDriveId;
    }

    if (!SITE_ID) {

        throw new Error(
            'SITE_ID não configurado.'
        );
    }

    if (!LIBRARY_NAME) {

        throw new Error(
            'LIBRARY_NAME não configurado.'
        );
    }

    const url =
        `https://graph.microsoft.com/v1.0/` +
        `sites/${SITE_ID}/drives`;

    try {

        const response =
            await axios.get(
                url,
                {
                    headers: {
                        Authorization:
                            `Bearer ${token}`
                    }
                }
            );

        const drives =
            response.data.value || [];

        const drive =
            drives.find(
                item =>
                    normalizeText(
                        item.name
                    ) ===
                    normalizeText(
                        LIBRARY_NAME
                    )
            );

        if (!drive) {

            const available =
                drives
                    .map(
                        item =>
                            item.name
                    )
                    .filter(Boolean)
                    .join(', ');

            throw new Error(
                `Biblioteca "${LIBRARY_NAME}" não encontrada. ` +
                `Disponíveis: ${available}`
            );
        }

        cachedDriveId =
            drive.id;

        console.log(
            `✅ Biblioteca localizada: ${drive.name}`
        );

        return cachedDriveId;

    } catch (error) {

        throw new Error(
            `Erro ao localizar biblioteca: ` +
            `${getErrorMessage(error)}`
        );
    }
}

// ============================================================
// LOCALIZA LISTA
// ============================================================

async function getListId(token) {

    if (cachedListId) {

        return cachedListId;
    }

    if (!SITE_ID) {

        throw new Error(
            'SITE_ID não configurado.'
        );
    }

    if (!LIST_NAME) {

        throw new Error(
            'LIST_NAME não configurado.'
        );
    }

    let url =
        `https://graph.microsoft.com/v1.0/` +
        `sites/${SITE_ID}/lists` +
        `?$select=id,name,displayName&$top=200`;

    const lists = [];

    try {

        while (url) {

            const response =
                await axios.get(
                    url,
                    {
                        headers: {
                            Authorization:
                                `Bearer ${token}`
                        }
                    }
                );

            lists.push(
                ...(response.data.value || [])
            );

            url =
                response.data[
                    '@odata.nextLink'
                ] || null;
        }

        const list =
            lists.find(
                item => {

                    return (
                        normalizeText(
                            item.displayName
                        ) ===
                            normalizeText(
                                LIST_NAME
                            ) ||
                        normalizeText(
                            item.name
                        ) ===
                            normalizeText(
                                LIST_NAME
                            )
                    );
                }
            );

        if (!list) {

            const available =
                lists
                    .map(
                        item =>
                            item.displayName ||
                            item.name
                    )
                    .filter(Boolean)
                    .join(', ');

            throw new Error(
                `Lista "${LIST_NAME}" não encontrada. ` +
                `Disponíveis: ${available}`
            );
        }

        cachedListId =
            list.id;

        console.log(
            `✅ Lista localizada: ` +
            `${list.displayName || list.name}`
        );

        return cachedListId;

    } catch (error) {

        throw new Error(
            `Erro ao localizar lista: ` +
            `${getErrorMessage(error)}`
        );
    }
}

// ============================================================
// COLUNAS DA LISTA
// ============================================================

async function getListColumns(
    token,
    listId
) {

    if (cachedColumns) {

        return cachedColumns;
    }

    let url =
        `https://graph.microsoft.com/v1.0/` +
        `sites/${SITE_ID}/lists/` +
        `${listId}/columns` +
        `?$top=200`;

    const columns = [];

    while (url) {

        const response =
            await axios.get(
                url,
                {
                    headers: {
                        Authorization:
                            `Bearer ${token}`
                    }
                }
            );

        columns.push(
            ...(response.data.value || [])
        );

        url =
            response.data[
                '@odata.nextLink'
            ] || null;
    }

    cachedColumns =
        columns;

    console.log(
        `✅ ${columns.length} coluna(s) encontrada(s) na lista.`
    );

    return cachedColumns;
}

// ============================================================
// LOCALIZA COLUNA
// ============================================================

function findColumn(
    columns,
    aliases
) {

    const normalizedAliases =
        aliases.map(
            alias =>
                normalizeText(alias)
        );

    return columns.find(
        column => {

            const name =
                normalizeText(
                    column.name
                );

            const displayName =
                normalizeText(
                    column.displayName
                );

            return (
                normalizedAliases.includes(
                    name
                ) ||
                normalizedAliases.includes(
                    displayName
                )
            );
        }
    );
}

// ============================================================
// DATA PT-BR PARA ISO
// ============================================================

function convertBrazilianDateToIso(
    value
) {

    if (!value) {

        return null;
    }

    const text =
        String(value).trim();

    if (
        /^\d{4}-\d{2}-\d{2}/
            .test(text)
    ) {

        const direct =
            new Date(text);

        if (
            !Number.isNaN(
                direct.getTime()
            )
        ) {

            return direct.toISOString();
        }
    }

    // 07/10/2026 16:02:00

    const match =
        text.match(
            /^(\d{1,2})\/(\d{1,2})\/(\d{4})(?:,\s*|\s+)?(\d{1,2})?:?(\d{2})?:?(\d{2})?$/
        );

    if (!match) {

        return null;
    }

    const day =
        Number(match[1]);

    const month =
        Number(match[2]) - 1;

    const year =
        Number(match[3]);

    const hour =
        Number(match[4] || 0);

    const minute =
        Number(match[5] || 0);

    const second =
        Number(match[6] || 0);

    const date =
        new Date(
            year,
            month,
            day,
            hour,
            minute,
            second
        );

    if (
        Number.isNaN(
            date.getTime()
        )
    ) {

        return null;
    }

    return date.toISOString();
}

// ============================================================
// CONVERTE VALOR PELO TIPO DA COLUNA
// ============================================================

function convertValueForColumn(
    column,
    value
) {

    if (
        value === undefined ||
        value === null ||
        value === ''
    ) {

        return null;
    }

    // --------------------------------------------------------
    // NÚMERO
    // --------------------------------------------------------

    if (column.number) {

        const text =
            String(value)
                .trim()
                .replace(',', '.');

        const number =
            Number(text);

        if (
            Number.isNaN(number)
        ) {

            return null;
        }

        return number;
    }

    // --------------------------------------------------------
    // BOOLEAN
    // --------------------------------------------------------

    if (column.boolean) {

        if (
            typeof value === 'boolean'
        ) {

            return value;
        }

        const normalized =
            normalizeText(value);

        return (
            normalized === 'true' ||
            normalized === 'sim' ||
            normalized === '1'
        );
    }

    // --------------------------------------------------------
    // DATA
    // --------------------------------------------------------

    if (column.dateTime) {

        return convertBrazilianDateToIso(
            value
        );
    }

    // --------------------------------------------------------
    // LINK / IMAGEM
    // --------------------------------------------------------

    if (
        column.hyperlinkOrPicture
    ) {

        return String(value);
    }

    // --------------------------------------------------------
    // CHOICE
    // --------------------------------------------------------

    if (column.choice) {

        return String(value);
    }

    // --------------------------------------------------------
    // TEXTO
    // --------------------------------------------------------

    return String(value);
}

// ============================================================
// ADICIONA CAMPO
// ============================================================

function addField(
    fields,
    columns,
    aliases,
    value
) {

    if (
        value === undefined ||
        value === null ||
        value === ''
    ) {

        return;
    }

    const column =
        findColumn(
            columns,
            aliases
        );

    if (!column) {

        console.warn(
            `⚠️ Coluna não localizada: ${aliases[0]}`
        );

        return;
    }

    if (column.readOnly) {

        console.warn(
            `⚠️ Coluna somente leitura ignorada: ` +
            `${column.displayName}`
        );

        return;
    }

    const converted =
        convertValueForColumn(
            column,
            value
        );

    if (
        converted === null ||
        converted === undefined
    ) {

        console.warn(
            `⚠️ Valor ignorado para ` +
            `${column.displayName}: ${value}`
        );

        return;
    }

    fields[
        column.name
    ] =
        converted;
}

// ============================================================
// MONTA CAMPOS PARA LISTA
// ============================================================

function buildListFields(
    row,
    ticketNumber,
    columns
) {

    const fields = {};

    const ticket =
        normalizeTicket(
            row['N° do ticket'] ||
            ticketNumber
        );

    const title =
        `${ticket} - ` +
        `${row.Item || ''} - ` +
        `${row.Motivo || ''}`;

    addField(
        fields,
        columns,
        [
            'Title',
            'Título',
            'Titulo'
        ],
        title.substring(
            0,
            255
        )
    );

    addField(
        fields,
        columns,
        [
            'N_x00b0_doticket',
            'N° do ticket',
            'Nº do ticket',
            'Número do ticket',
            'Numero do ticket'
        ],
        ticket
    );

    addField(
        fields,
        columns,
        [
            'NomedoCliente',
            'Nome do Cliente'
        ],
        row['Nome do Cliente']
    );

    addField(
        fields,
        columns,
        [
            'Item'
        ],
        row.Item
    );

    addField(
        fields,
        columns,
        [
            'Qtde',
            'Quantidade'
        ],
        row.Qtde
    );

    addField(
        fields,
        columns,
        [
            'Motivo'
        ],
        row.Motivo
    );

    addField(
        fields,
        columns,
        [
            'Origemdodefeito',
            'Origem do defeito'
        ],
        row[
            'Origem do defeito'
        ]
    );

    addField(
        fields,
        columns,
        [
            'Disposi_x00e7__x00e3_o',
            'Disposição',
            'Disposicao'
        ],
        row['Disposição']
    );

    addField(
        fields,
        columns,
        [
            'Disposi_x00e7__x00e3_odaspe_x00e',
            'Disposição das peças',
            'Disposicao das pecas'
        ],
        row[
            'Disposição das peças'
        ]
    );

    addField(
        fields,
        columns,
        [
            'DatadeGera_x00e7__x00e3_o',
            'Data de Geração',
            'Data de Geracao'
        ],
        row[
            'Data de Geração'
        ]
    );

    for (
        let i = 1;
        i <= 10;
        i++
    ) {

        addField(
            fields,
            columns,
            [
                `Foto${i}`,
                `Foto ${i}`
            ],
            row[
                `Foto ${i}`
            ]
        );
    }

    return fields;
}

// ============================================================
// CAMINHO SHAREPOINT
// ============================================================

function getEncodedFolderPath() {

    return String(
        FOLDER_PATH || ''
    )
        .replace(
            /^\/+/,
            ''
        )
        .replace(
            /\/+$/,
            ''
        )
        .split('/')
        .filter(Boolean)
        .map(
            segment =>
                encodeURIComponent(
                    segment
                )
        )
        .join('/');
}

// ============================================================
// LISTA ARQUIVOS DA PASTA
// ============================================================

async function getAllFilesFromFolder(
    token,
    driveId
) {

    const files = [];

    const folder =
        getEncodedFolderPath();

    let url;

    if (folder) {

        url =
            `https://graph.microsoft.com/v1.0/` +
            `drives/${driveId}/root:/` +
            `${folder}:/children` +
            `?$select=id,name,createdDateTime,` +
            `lastModifiedDateTime,file,folder,webUrl,size` +
            `&$top=200`;

    } else {

        url =
            `https://graph.microsoft.com/v1.0/` +
            `drives/${driveId}/root/children` +
            `?$select=id,name,createdDateTime,` +
            `lastModifiedDateTime,file,folder,webUrl,size` +
            `&$top=200`;
    }

    while (url) {

        try {

            const response =
                await axios.get(
                    url,
                    {
                        headers: {
                            Authorization:
                                `Bearer ${token}`
                        }
                    }
                );

            files.push(
                ...(response.data.value || [])
            );

            url =
                response.data[
                    '@odata.nextLink'
                ] || null;

        } catch (error) {

            throw new Error(
                `Erro ao listar arquivos da pasta ` +
                `"${FOLDER_PATH}": ` +
                `${getErrorMessage(error)}`
            );
        }
    }

    return files;
}

// ============================================================
// EXTRAI TICKET DO NOME DO PDF
// ============================================================
//
// FUNCIONA COM:
//
// Laudo - SR-7726-20261007 1613.pdf
// Laudo - SR-7726-20261007 1602.pdf
// Laudo - SR-7726-1791394730858.pdf
// Laudo - SR-7726-1791394439951.pdf
//
// RETORNA:
//
// SR-7726
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

    // --------------------------------------------------------
    // Ticket com letras + número
    // Ex: SR-7726
    // --------------------------------------------------------

    let match =
        name.match(
            /^Laudo\s*-\s*([A-Za-z]+-\d+)(?:-.+)?\.pdf$/i
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
            /^Laudo\s*-\s*(\d+)(?:-.+)?\.pdf$/i
        );

    if (match) {

        return normalizeTicket(
            match[1]
        );
    }

    return null;
}

// ============================================================
// DATA CONTIDA NO NOVO NOME DO PDF
// ============================================================

function getTimestampFromPdfName(
    fileName
) {

    if (!fileName) {

        return 0;
    }

    // Ex:
    // Laudo - SR-7726-20261007 1613.pdf

    const match =
        String(fileName).match(
            /-(\d{4})(\d{2})(\d{2})[ _-](\d{2})(\d{2})(?:\s*\(\d+\))?\.pdf$/i
        );

    if (!match) {

        return 0;
    }

    const year =
        Number(match[1]);

    const month =
        Number(match[2]);

    const day =
        Number(match[3]);

    const hour =
        Number(match[4]);

    const minute =
        Number(match[5]);

    // Apenas fallback.
    // Para determinar o mais recente,
    // usamos primeiro lastModifiedDateTime do SharePoint.

    const timestamp =
        Date.UTC(
            year,
            month - 1,
            day,
            hour,
            minute,
            0
        );

    return Number.isFinite(
        timestamp
    )
        ? timestamp
        : 0;
}

// ============================================================
// DATA USADA PARA ORDENAR PDFs
// ============================================================

function getPdfSortTimestamp(
    file
) {

    // --------------------------------------------------------
    // 1. SHAREPOINT MODIFIED
    // --------------------------------------------------------

    const modified =
        Date.parse(
            file.lastModifiedDateTime ||
            ''
        );

    if (
        Number.isFinite(modified)
    ) {

        return modified;
    }

    // --------------------------------------------------------
    // 2. SHAREPOINT CREATED
    // --------------------------------------------------------

    const created =
        Date.parse(
            file.createdDateTime ||
            ''
        );

    if (
        Number.isFinite(created)
    ) {

        return created;
    }

    // --------------------------------------------------------
    // 3. DATA NO NOME
    // --------------------------------------------------------

    return getTimestampFromPdfName(
        file.name
    );
}

// ============================================================
// PDF DO TICKET EXISTE?
// ============================================================

async function ticketPdfExists(
    token,
    driveId,
    ticketNumber
) {

    const ticket =
        normalizeTicket(
            ticketNumber
        );

    const files =
        await getAllFilesFromFolder(
            token,
            driveId
        );

    return files.some(
        file => {

            if (!file.file) {

                return false;
            }

            return (
                extractTicketNumber(
                    file.name
                ) ===
                ticket
            );
        }
    );
}

// ============================================================
// TICKET EXISTE NA LISTA?
// ============================================================

async function ticketExistsInList(
    token,
    listId,
    columns,
    ticketNumber
) {

    const ticket =
        normalizeTicket(
            ticketNumber
        );

    const ticketColumn =
        findColumn(
            columns,
            [
                'N_x00b0_doticket',
                'N° do ticket',
                'Nº do ticket',
                'Número do ticket',
                'Numero do ticket'
            ]
        );

    if (!ticketColumn) {

        throw new Error(
            'Coluna do número do ticket não encontrada.'
        );
    }

    let url =
        `https://graph.microsoft.com/v1.0/` +
        `sites/${SITE_ID}/lists/` +
        `${listId}/items` +
        `?$expand=fields&$top=200`;

    while (url) {

        const response =
            await axios.get(
                url,
                {
                    headers: {
                        Authorization:
                            `Bearer ${token}`
                    }
                }
            );

        const items =
            response.data.value || [];

        const found =
            items.some(
                item => {

                    const value =
                        item.fields?.[
                            ticketColumn.name
                        ];

                    return (
                        normalizeTicket(
                            value
                        ) ===
                        ticket
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
            status:
                'online',

            timestamp:
                new Date()
                    .toISOString(),

            configuration: {
                tenantConfigured:
                    Boolean(TENANT_ID),

                clientConfigured:
                    Boolean(CLIENT_ID),

                secretConfigured:
                    Boolean(CLIENT_SECRET),

                siteConfigured:
                    Boolean(SITE_ID),

                listName:
                    LIST_NAME,

                libraryName:
                    LIBRARY_NAME,

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

            const ticket =
                normalizeTicket(
                    req.params.ticketNumber
                );

            if (!ticket) {

                return res
                    .status(400)
                    .json({
                        success:
                            false,

                        error:
                            'Ticket obrigatório.'
                    });
            }

            console.log(
                `🔎 Verificando ticket ${ticket}...`
            );

            const token =
                await getAccessToken();

            const [
                driveId,
                listId
            ] =
                await Promise.all([
                    getDriveId(token),
                    getListId(token)
                ]);

            const columns =
                await getListColumns(
                    token,
                    listId
                );

            const [
                existsInPdf,
                existsInList
            ] =
                await Promise.all([
                    ticketPdfExists(
                        token,
                        driveId,
                        ticket
                    ),

                    ticketExistsInList(
                        token,
                        listId,
                        columns,
                        ticket
                    )
                ]);

            console.log(
                `🔎 ${ticket} | ` +
                `PDF: ${existsInPdf} | ` +
                `Lista: ${existsInList}`
            );

            return res.json({
                success:
                    true,

                ticketNumber:
                    ticket,

                existsInPdf,

                existsInList
            });

        } catch (error) {

            const message =
                getErrorMessage(
                    error
                );

            console.error(
                '❌ Erro check-status:',
                message
            );

            return res
                .status(500)
                .json({
                    success:
                        false,

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
                ticketNumber
            } =
                req.body;

            if (
                !fileName ||
                !fileBase64
            ) {

                return res
                    .status(400)
                    .json({
                        success:
                            false,

                        error:
                            'fileName e fileBase64 são obrigatórios.'
                    });
            }

            console.log(
                `📄 Enviando PDF: ${fileName}`
            );

            console.log(
                `🎫 Ticket: ${ticketNumber || ''}`
            );

            const token =
                await getAccessToken();

            const driveId =
                await getDriveId(
                    token
                );

            const folder =
                getEncodedFolderPath();

            const encodedFile =
                encodeURIComponent(
                    fileName
                );

            let url;

            if (folder) {

                url =
                    `https://graph.microsoft.com/v1.0/` +
                    `drives/${driveId}/root:/` +
                    `${folder}/${encodedFile}:/content`;

            } else {

                url =
                    `https://graph.microsoft.com/v1.0/` +
                    `drives/${driveId}/root:/` +
                    `${encodedFile}:/content`;
            }

            const base64 =
                String(fileBase64)
                    .replace(
                        /^data:application\/pdf;base64,/i,
                        ''
                    );

            const buffer =
                Buffer.from(
                    base64,
                    'base64'
                );

            const response =
                await axios.put(
                    url,
                    buffer,
                    {
                        headers: {
                            Authorization:
                                `Bearer ${token}`,

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
                success:
                    true,

                sharePointUrl:
                    response.data.webUrl,

                fileName:
                    response.data.name,

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
                getErrorMessage(
                    error
                );

            console.error(
                '❌ Erro upload PDF:',
                message
            );

            return res
                .status(500)
                .json({
                    success:
                        false,

                    error:
                        message
                });
        }
    }
);

// ============================================================
// UPLOAD LISTA
// ============================================================

app.post(
    '/upload-list-data',
    async (req, res) => {

        try {

            const {
                ticketNumber,
                listData
            } =
                req.body;

            if (
                !Array.isArray(
                    listData
                ) ||
                listData.length === 0
            ) {

                return res
                    .status(400)
                    .json({
                        success:
                            false,

                        error:
                            'listData não informado.'
                    });
            }

            console.log(
                `📋 Enviando ticket ${ticketNumber} para a lista...`
            );

            const token =
                await getAccessToken();

            const listId =
                await getListId(
                    token
                );

            const columns =
                await getListColumns(
                    token,
                    listId
                );

            let inserted = 0;

            for (
                let index = 0;
                index < listData.length;
                index++
            ) {

                const row =
                    listData[index];

                const fields =
                    buildListFields(
                        row,
                        ticketNumber,
                        columns
                    );

                if (
                    Object.keys(
                        fields
                    ).length === 0
                ) {

                    throw new Error(
                        'Nenhuma coluna compatível foi encontrada.'
                    );
                }

                console.log(
                    `📋 Linha ${index + 1}/` +
                    `${listData.length}`
                );

                console.log(
                    '📋 Campos que serão enviados:',
                    Object.keys(
                        fields
                    ).join(', ')
                );

                const url =
                    `https://graph.microsoft.com/v1.0/` +
                    `sites/${SITE_ID}/lists/` +
                    `${listId}/items`;

                try {

                    const response =
                        await axios.post(
                            url,
                            {
                                fields
                            },
                            {
                                headers: {
                                    Authorization:
                                        `Bearer ${token}`,

                                    'Content-Type':
                                        'application/json',

                                    Accept:
                                        'application/json'
                                }
                            }
                        );

                    inserted++;

                    console.log(
                        `✅ Linha ${index + 1} criada. ` +
                        `ID ${response.data.id}`
                    );

                } catch (error) {

                    const graph =
                        error
                            ?.response
                            ?.data;

                    console.error(
                        '======================================'
                    );

                    console.error(
                        `❌ ERRO LINHA ${index + 1}`
                    );

                    console.error(
                        `❌ Ticket: ${ticketNumber}`
                    );

                    console.error(
                        `❌ HTTP: ` +
                        `${error?.response?.status || ''}`
                    );

                    console.error(
                        '❌ Campos enviados:'
                    );

                    console.error(
                        JSON.stringify(
                            fields,
                            null,
                            2
                        )
                    );

                    console.error(
                        '❌ Retorno Graph:'
                    );

                    console.error(
                        JSON.stringify(
                            graph,
                            null,
                            2
                        )
                    );

                    console.error(
                        '======================================'
                    );

                    return res
                        .status(500)
                        .json({
                            success:
                                false,

                            error:
                                graph?.error?.message ||
                                error.message,

                            row:
                                index + 1,

                            graphCode:
                                graph?.error?.code,

                            requestId:
                                graph
                                    ?.error
                                    ?.innerError
                                    ?.[
                                        'request-id'
                                    ]
                        });
                }
            }

            console.log(
                `✅ ${inserted} linha(s) inserida(s) ` +
                `para ${ticketNumber}.`
            );

            return res.json({
                success:
                    true,

                ticketNumber,

                inserted
            });

        } catch (error) {

            const message =
                getErrorMessage(
                    error
                );

            console.error(
                '❌ Erro upload lista:',
                message
            );

            return res
                .status(500)
                .json({
                    success:
                        false,

                    error:
                        message
                });
        }
    }
);

// ============================================================
// EXCLUI PDFs DO TICKET
// ============================================================

app.delete(
    '/delete-pdf-by-ticket-number/:ticketNumber',
    async (req, res) => {

        try {

            const ticket =
                normalizeTicket(
                    req.params.ticketNumber
                );

            if (!ticket) {

                return res
                    .status(400)
                    .json({
                        success:
                            false,

                        error:
                            'Ticket obrigatório.'
                    });
            }

            const token =
                await getAccessToken();

            const driveId =
                await getDriveId(
                    token
                );

            const files =
                await getAllFilesFromFolder(
                    token,
                    driveId
                );

            const matching =
                files.filter(
                    file =>
                        Boolean(
                            file.file
                        ) &&
                        extractTicketNumber(
                            file.name
                        ) ===
                            ticket
                );

            let deletedCount = 0;
            let alreadyMissingCount = 0;

            const deletedFiles = [];

            for (
                const file
                of matching
            ) {

                try {

                    await axios.delete(
                        `https://graph.microsoft.com/v1.0/` +
                        `drives/${driveId}/items/` +
                        `${file.id}`,
                        {
                            headers: {
                                Authorization:
                                    `Bearer ${token}`
                            }
                        }
                    );

                    deletedCount++;

                    deletedFiles.push(
                        file.name
                    );

                } catch (error) {

                    if (
                        error?.response?.status ===
                        404
                    ) {

                        alreadyMissingCount++;

                        console.warn(
                            `⚠️ PDF já não existe: ${file.name}`
                        );

                        continue;
                    }

                    throw error;
                }
            }

            return res.json({
                success:
                    true,

                ticketNumber:
                    ticket,

                deletedCount,

                alreadyMissingCount,

                deletedFiles
            });

        } catch (error) {

            const message =
                getErrorMessage(
                    error
                );

            console.error(
                '❌ Erro ao excluir PDFs:',
                message
            );

            return res
                .status(500)
                .json({
                    success:
                        false,

                    error:
                        message
                });
        }
    }
);

// ============================================================
// LIMPA PDFs DUPLICADOS
// ============================================================
//
// REGRAS:
//
// 1. Lista todos os PDFs da pasta Laudos.
//
// 2. Reconhece:
//
//    Laudo - SR-7726-20261007 1613.pdf
//    Laudo - SR-7726-1791394730858.pdf
//
// 3. Agrupa exatamente pelo ticket:
//
//    SR-7726
//
// 4. Ordena pelo lastModifiedDateTime do SharePoint.
//
// 5. Mantém somente o mais recente.
//
// 6. Exclui os demais.
//
// 7. Se um arquivo retornar 404, considera que ele
//    já foi removido e continua.
//
// 8. NÃO remove registros da Lista SharePoint.
//
// ============================================================

app.post(
    '/cleanup-duplicate-pdfs',
    async (req, res) => {

        try {

            console.log(
                '🧹 Iniciando limpeza de PDFs duplicados...'
            );

            const token =
                await getAccessToken();

            console.log(
                '✅ Token obtido.'
            );

            const driveId =
                await getDriveId(
                    token
                );

            console.log(
                `✅ Biblioteca localizada: ${LIBRARY_NAME}`
            );

            const files =
                await getAllFilesFromFolder(
                    token,
                    driveId
                );

            console.log(
                `📂 ${files.length} item(ns) localizado(s) ` +
                `na pasta ${FOLDER_PATH}.`
            );

            const groups =
                new Map();

            let recognizedPdfCount = 0;

            // ------------------------------------------------
            // AGRUPAMENTO
            // ------------------------------------------------

            for (
                const file
                of files
            ) {

                if (!file.file) {

                    continue;
                }

                if (
                    !file.name ||
                    !file.name
                        .toLowerCase()
                        .endsWith('.pdf')
                ) {

                    continue;
                }

                const ticket =
                    extractTicketNumber(
                        file.name
                    );

                if (!ticket) {

                    console.log(
                        `ℹ️ PDF ignorado: ${file.name}`
                    );

                    continue;
                }

                recognizedPdfCount++;

                if (
                    !groups.has(
                        ticket
                    )
                ) {

                    groups.set(
                        ticket,
                        []
                    );
                }

                groups
                    .get(ticket)
                    .push(file);
            }

            console.log(
                `🎫 ${groups.size} ticket(s) reconhecido(s).`
            );

            const duplicates = [];
            const keptFiles = [];
            const deletedFiles = [];
            const alreadyMissingFiles = [];
            const failedFiles = [];

            // ------------------------------------------------
            // PROCESSA TICKET POR TICKET
            // ------------------------------------------------

            for (
                const [
                    ticket,
                    ticketFiles
                ]
                of groups.entries()
            ) {

                if (
                    ticketFiles.length <= 1
                ) {

                    continue;
                }

                console.log(
                    `⚠️ ${ticket}: ` +
                    `${ticketFiles.length} PDF(s) encontrado(s).`
                );

                // --------------------------------------------
                // MAIS RECENTE PRIMEIRO
                // --------------------------------------------

                ticketFiles.sort(
                    (a, b) => {

                        return (
                            getPdfSortTimestamp(
                                b
                            ) -
                            getPdfSortTimestamp(
                                a
                            )
                        );
                    }
                );

                const keepFile =
                    ticketFiles[0];

                const filesToDelete =
                    ticketFiles.slice(1);

                console.log(
                    `✅ ${ticket}: mantendo ` +
                    `"${keepFile.name}"`
                );

                keptFiles.push({
                    ticketNumber:
                        ticket,

                    fileName:
                        keepFile.name,

                    modifiedDateTime:
                        keepFile
                            .lastModifiedDateTime
                });

                duplicates.push({
                    ticketNumber:
                        ticket,

                    totalFiles:
                        ticketFiles.length,

                    keptFile:
                        keepFile.name,

                    filesToDelete:
                        filesToDelete.map(
                            file =>
                                file.name
                        )
                });

                // --------------------------------------------
                // EXCLUSÃO
                // --------------------------------------------

                for (
                    const file
                    of filesToDelete
                ) {

                    console.log(
                        `🗑️ Tentando excluir: ${file.name}`
                    );

                    const deleteUrl =
                        `https://graph.microsoft.com/v1.0/` +
                        `drives/${driveId}/items/` +
                        `${file.id}`;

                    try {

                        await axios.delete(
                            deleteUrl,
                            {
                                headers: {
                                    Authorization:
                                        `Bearer ${token}`
                                }
                            }
                        );

                        deletedFiles.push({
                            ticketNumber:
                                ticket,

                            fileName:
                                file.name
                        });

                        console.log(
                            `✅ PDF duplicado excluído: ` +
                            `${file.name}`
                        );

                    } catch (error) {

                        const status =
                            error
                                ?.response
                                ?.status;

                        const message =
                            getErrorMessage(
                                error
                            );

                        // ------------------------------------
                        // 404
                        //
                        // O arquivo já não existe.
                        // Não aborta toda a limpeza.
                        // ------------------------------------

                        if (status === 404) {

                            alreadyMissingFiles.push({
                                ticketNumber:
                                    ticket,

                                fileName:
                                    file.name
                            });

                            console.warn(
                                `⚠️ Arquivo já não existe: ` +
                                `${file.name}`
                            );

                            continue;
                        }

                        // ------------------------------------
                        // OUTROS ERROS
                        //
                        // Registra e continua.
                        // ------------------------------------

                        failedFiles.push({
                            ticketNumber:
                                ticket,

                            fileName:
                                file.name,

                            status,

                            error:
                                message
                        });

                        console.error(
                            `❌ Falha ao excluir ` +
                            `"${file.name}": ` +
                            `${status || ''} ${message}`
                        );
                    }
                }
            }

            // ------------------------------------------------
            // RESULTADO
            // ------------------------------------------------

            console.log(
                '=========================================='
            );

            console.log(
                '🧹 LIMPEZA CONCLUÍDA'
            );

            console.log(
                `📂 Arquivos verificados: ${files.length}`
            );

            console.log(
                `📄 PDFs reconhecidos: ${recognizedPdfCount}`
            );

            console.log(
                `⚠️ Tickets com duplicidade: ${duplicates.length}`
            );

            console.log(
                `🗑️ PDFs excluídos: ${deletedFiles.length}`
            );

            console.log(
                `ℹ️ Arquivos já ausentes: ` +
                `${alreadyMissingFiles.length}`
            );

            console.log(
                `❌ Falhas: ${failedFiles.length}`
            );

            console.log(
                '=========================================='
            );

            const success =
                failedFiles.length === 0;

            return res.json({
                success,

                message:
                    `${deletedFiles.length} PDF(s) ` +
                    `duplicado(s) removido(s).`,

                totalFilesChecked:
                    files.length,

                recognizedPdfCount,

                ticketsChecked:
                    groups.size,

                ticketsWithDuplicates:
                    duplicates.length,

                deletedCount:
                    deletedFiles.length,

                alreadyMissingCount:
                    alreadyMissingFiles.length,

                failedCount:
                    failedFiles.length,

                duplicates,

                keptFiles,

                deletedFiles,

                alreadyMissingFiles,

                failedFiles,

                error:
                    success
                        ? undefined
                        : `${failedFiles.length} arquivo(s) ` +
                          `não puderam ser excluídos.`
            });

        } catch (error) {

            const message =
                getErrorMessage(
                    error
                );

            console.error(
                '❌ Erro geral na limpeza:',
                message
            );

            if (
                error?.response?.data
            ) {

                console.error(
                    '❌ Retorno Graph:'
                );

                console.error(
                    JSON.stringify(
                        error.response.data,
                        null,
                        2
                    )
                );
            }

            return res
                .status(500)
                .json({
                    success:
                        false,

                    error:
                        message
                });
        }
    }
);

// ============================================================
// APAGA TODOS OS ITENS DA LISTA
// ============================================================
//
// ATENÇÃO:
//
// Esta rota NÃO participa da limpeza de PDFs.
//
// Só será executada se o frontend chamar:
//
// DELETE /clear-list
//
// ============================================================

app.delete(
    '/clear-list',
    async (req, res) => {

        try {

            console.log(
                '⚠️ Iniciando exclusão TOTAL da lista...'
            );

            const token =
                await getAccessToken();

            const listId =
                await getListId(
                    token
                );

            let deletedCount = 0;

            while (true) {

                const url =
                    `https://graph.microsoft.com/v1.0/` +
                    `sites/${SITE_ID}/lists/` +
                    `${listId}/items?$top=200`;

                const response =
                    await axios.get(
                        url,
                        {
                            headers: {
                                Authorization:
                                    `Bearer ${token}`
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

                    await axios.delete(
                        `https://graph.microsoft.com/v1.0/` +
                        `sites/${SITE_ID}/lists/` +
                        `${listId}/items/` +
                        `${item.id}`,
                        {
                            headers: {
                                Authorization:
                                    `Bearer ${token}`
                            }
                        }
                    );

                    deletedCount++;
                }

                console.log(
                    `🗑️ ${deletedCount} registro(s) ` +
                    `removido(s) até agora...`
                );
            }

            return res.json({
                success:
                    true,

                deletedCount,

                message:
                    `${deletedCount} item(ns) ` +
                    `removido(s) da lista.`
            });

        } catch (error) {

            const message =
                getErrorMessage(
                    error
                );

            console.error(
                '❌ Erro clear-list:',
                message
            );

            return res
                .status(500)
                .json({
                    success:
                        false,

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
            success:
                false,

            error:
                `Rota não encontrada: ` +
                `${req.method} ${req.originalUrl}`
        });
    }
);

// ============================================================
// START
// ============================================================

app.listen(
    PORT,
    '0.0.0.0',
    () => {

        console.log(
            '🚀 API SharePoint Global Plastic'
        );

        console.log(
            `🌐 API online na porta ${PORT}`
        );

        console.log(
            `🔗 SITE_ID: ` +
            `${SITE_ID ? 'OK' : 'NÃO CONFIGURADO'}`
        );

        console.log(
            `📋 LIST_NAME: ${LIST_NAME}`
        );

        console.log(
            `📚 LIBRARY_NAME: ${LIBRARY_NAME}`
        );

        console.log(
            `📂 FOLDER_PATH: ${FOLDER_PATH}`
        );
    }
);
