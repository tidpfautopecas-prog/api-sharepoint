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
const FOLDER_PATH = process.env.FOLDER_PATH || 'Laudos';

app.use(cors());
app.use(express.json({ limit: '100mb' }));
app.use(express.urlencoded({ extended: true, limit: '100mb' }));

// ============================================================
// CACHE
// ============================================================

let cachedToken = null;
let cachedTokenExpiresAt = 0;
let cachedDriveId = null;
let cachedListId = null;
let cachedColumns = null;

// ============================================================
// AUXILIARES
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

function normalizeText(value) {
    return String(value ?? '')
        .normalize('NFD')
        .replace(/[\u0300-\u036f]/g, '')
        .trim()
        .toLowerCase();
}

function normalizeTicket(value) {
    return String(value ?? '')
        .trim()
        .toUpperCase()
        .replace(/^#/, '')
        .replace(/[^A-Z0-9-]/g, '');
}

function validateEnvironment() {
    const required = [
        'TENANT_ID',
        'CLIENT_ID',
        'CLIENT_SECRET',
        'SITE_ID',
        'LIBRARY_NAME',
        'LIST_NAME'
    ];

    const missing = required.filter(name => !process.env[name]);

    if (missing.length) {
        console.error(
            `❌ Variáveis ausentes: ${missing.join(', ')}`
        );
    } else {
        console.log('✅ Variáveis de ambiente configuradas.');
    }
}

validateEnvironment();

// ============================================================
// MICROSOFT TOKEN
// ============================================================

async function getAccessToken() {
    if (
        cachedToken &&
        Date.now() < cachedTokenExpiresAt
    ) {
        return cachedToken;
    }

    const url =
        `https://login.microsoftonline.com/` +
        `${TENANT_ID}/oauth2/v2.0/token`;

    const body = new URLSearchParams();

    body.append('client_id', CLIENT_ID);
    body.append('client_secret', CLIENT_SECRET);
    body.append(
        'scope',
        'https://graph.microsoft.com/.default'
    );
    body.append(
        'grant_type',
        'client_credentials'
    );

    try {
        const response = await axios.post(
            url,
            body.toString(),
            {
                headers: {
                    'Content-Type':
                        'application/x-www-form-urlencoded'
                }
            }
        );

        cachedToken = response.data.access_token;

        const expires =
            Number(response.data.expires_in || 3600);

        cachedTokenExpiresAt =
            Date.now() +
            Math.max(expires - 300, 60) * 1000;

        return cachedToken;

    } catch (error) {
        throw new Error(
            `Erro na autenticação Microsoft: ` +
            getErrorMessage(error)
        );
    }
}

// ============================================================
// DRIVE / BIBLIOTECA
// ============================================================

async function getDriveId(token) {
    if (cachedDriveId) {
        return cachedDriveId;
    }

    const url =
        `https://graph.microsoft.com/v1.0/` +
        `sites/${SITE_ID}/drives`;

    try {
        const response = await axios.get(url, {
            headers: {
                Authorization: `Bearer ${token}`
            }
        });

        const drives = response.data.value || [];

        const drive = drives.find(
            d =>
                normalizeText(d.name) ===
                normalizeText(LIBRARY_NAME)
        );

        if (!drive) {
            throw new Error(
                `Biblioteca "${LIBRARY_NAME}" não encontrada. ` +
                `Disponíveis: ` +
                drives.map(d => d.name).join(', ')
            );
        }

        cachedDriveId = drive.id;

        console.log(
            `✅ Biblioteca localizada: ${drive.name}`
        );

        return cachedDriveId;

    } catch (error) {
        throw new Error(
            `Erro ao localizar biblioteca: ` +
            getErrorMessage(error)
        );
    }
}

// ============================================================
// LISTA
// ============================================================

async function getListId(token) {
    if (cachedListId) {
        return cachedListId;
    }

    let url =
        `https://graph.microsoft.com/v1.0/` +
        `sites/${SITE_ID}/lists` +
        `?$select=id,name,displayName&$top=200`;

    const lists = [];

    try {
        while (url) {
            const response = await axios.get(url, {
                headers: {
                    Authorization: `Bearer ${token}`
                }
            });

            lists.push(
                ...(response.data.value || [])
            );

            url =
                response.data['@odata.nextLink'] ||
                null;
        }

        const list = lists.find(
            l =>
                normalizeText(l.displayName) ===
                    normalizeText(LIST_NAME) ||
                normalizeText(l.name) ===
                    normalizeText(LIST_NAME)
        );

        if (!list) {
            throw new Error(
                `Lista "${LIST_NAME}" não encontrada. ` +
                `Disponíveis: ` +
                lists
                    .map(
                        l =>
                            l.displayName ||
                            l.name
                    )
                    .join(', ')
            );
        }

        cachedListId = list.id;

        console.log(
            `✅ Lista localizada: ` +
            `${list.displayName || list.name}`
        );

        return cachedListId;

    } catch (error) {
        throw new Error(
            `Erro ao localizar lista: ` +
            getErrorMessage(error)
        );
    }
}

// ============================================================
// COLUNAS REAIS DA LISTA
// ============================================================

async function getListColumns(token, listId) {
    if (cachedColumns) {
        return cachedColumns;
    }

    let url =
        `https://graph.microsoft.com/v1.0/` +
        `sites/${SITE_ID}/lists/${listId}/columns` +
        `?$top=200`;

    const columns = [];

    while (url) {
        const response = await axios.get(url, {
            headers: {
                Authorization: `Bearer ${token}`
            }
        });

        columns.push(
            ...(response.data.value || [])
        );

        url =
            response.data['@odata.nextLink'] ||
            null;
    }

    cachedColumns = columns;

    console.log(
        `✅ ${columns.length} colunas encontradas na lista.`
    );

    return columns;
}

function findColumn(columns, aliases) {
    const normalized =
        aliases.map(normalizeText);

    return columns.find(column => {
        const name =
            normalizeText(column.name);

        const display =
            normalizeText(column.displayName);

        return (
            normalized.includes(name) ||
            normalized.includes(display)
        );
    });
}

// ============================================================
// CONVERSÃO DE DATA PT-BR -> ISO
// ============================================================

function convertBrazilianDateToIso(value) {
    if (!value) {
        return null;
    }

    const text = String(value).trim();

    // Já é uma data válida/ISO?
    const directDate = new Date(text);

    if (
        !Number.isNaN(directDate.getTime()) &&
        /^\d{4}-\d{2}-\d{2}/.test(text)
    ) {
        return directDate.toISOString();
    }

    // Ex:
    // 07/10/2026 16:02:00
    const match = text.match(
        /^(\d{1,2})\/(\d{1,2})\/(\d{4})(?:,\s*|\s+)?(\d{1,2})?:?(\d{2})?:?(\d{2})?$/
    );

    if (!match) {
        return null;
    }

    const day = Number(match[1]);
    const month = Number(match[2]) - 1;
    const year = Number(match[3]);
    const hour = Number(match[4] || 0);
    const minute = Number(match[5] || 0);
    const second = Number(match[6] || 0);

    const date = new Date(
        year,
        month,
        day,
        hour,
        minute,
        second
    );

    if (Number.isNaN(date.getTime())) {
        return null;
    }

    return date.toISOString();
}

// ============================================================
// VALOR COMPATÍVEL COM O TIPO DA COLUNA
// ============================================================

function convertValueForColumn(column, value) {
    if (
        value === undefined ||
        value === null ||
        value === ''
    ) {
        return null;
    }

    if (column.number) {
        const number = Number(
            String(value)
                .replace(/\./g, '')
                .replace(',', '.')
        );

        return Number.isNaN(number)
            ? null
            : number;
    }

    if (column.boolean) {
        return Boolean(value);
    }

    if (column.dateTime) {
        return convertBrazilianDateToIso(value);
    }

    // URL/Hyperlink no Graph é enviada como texto URL.
    if (column.hyperlinkOrPicture) {
        return String(value);
    }

    if (column.choice) {
        return String(value);
    }

    if (column.text) {
        return String(value);
    }

    return String(value);
}

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
        findColumn(columns, aliases);

    if (!column) {
        console.warn(
            `⚠️ Coluna não encontrada: ` +
            aliases[0]
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
            `⚠️ Valor inválido ignorado. ` +
            `Coluna: ${column.displayName} | ` +
            `Valor: ${value}`
        );

        return;
    }

    fields[column.name] =
        converted;
}

// ============================================================
// MONTA CAMPOS DA LISTA
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
        ['Title', 'Título', 'Titulo'],
        title.substring(0, 255)
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
        ['Item'],
        row.Item
    );

    addField(
        fields,
        columns,
        ['Qtde', 'Quantidade'],
        row.Qtde
    );

    addField(
        fields,
        columns,
        ['Motivo'],
        row.Motivo
    );

    addField(
        fields,
        columns,
        [
            'Origemdodefeito',
            'Origem do defeito'
        ],
        row['Origem do defeito']
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
        row['Disposição das peças']
    );

    addField(
        fields,
        columns,
        [
            'DatadeGera_x00e7__x00e3_o',
            'Data de Geração',
            'Data de Geracao'
        ],
        row['Data de Geração']
    );

    for (let i = 1; i <= 10; i++) {
        addField(
            fields,
            columns,
            [
                `Foto${i}`,
                `Foto ${i}`
            ],
            row[`Foto ${i}`]
        );
    }

    return fields;
}

// ============================================================
// CAMINHO SHAREPOINT
// ============================================================

function getEncodedFolderPath() {
    return String(FOLDER_PATH || '')
        .replace(/^\/+/, '')
        .replace(/\/+$/, '')
        .split('/')
        .filter(Boolean)
        .map(encodeURIComponent)
        .join('/');
}

// ============================================================
// ARQUIVOS DA PASTA
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
            `?$top=200`;
    } else {
        url =
            `https://graph.microsoft.com/v1.0/` +
            `drives/${driveId}/root/children` +
            `?$top=200`;
    }

    while (url) {
        const response = await axios.get(url, {
            headers: {
                Authorization: `Bearer ${token}`
            }
        });

        files.push(
            ...(response.data.value || [])
        );

        url =
            response.data['@odata.nextLink'] ||
            null;
    }

    return files;
}

// ============================================================
// TICKET PELO NOME DO PDF
// ============================================================

function extractTicketNumber(fileName) {
    if (!fileName) {
        return null;
    }

    const name =
        String(fileName).trim();

    let match = name.match(
        /^Laudo\s*-\s*([A-Za-z]+-\d+)(?:-.+)?\.pdf$/i
    );

    if (match) {
        return normalizeTicket(match[1]);
    }

    match = name.match(
        /^Laudo\s*-\s*(\d+)(?:-.+)?\.pdf$/i
    );

    if (match) {
        return normalizeTicket(match[1]);
    }

    return null;
}

// ============================================================
// PDF EXISTE?
// ============================================================

async function ticketPdfExists(
    token,
    driveId,
    ticketNumber
) {
    const target =
        normalizeTicket(ticketNumber);

    const files =
        await getAllFilesFromFolder(
            token,
            driveId
        );

    return files.some(
        file =>
            file.file &&
            extractTicketNumber(file.name) ===
                target
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
    const target =
        normalizeTicket(ticketNumber);

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
            'Coluna do número do ticket não encontrada na lista.'
        );
    }

    let url =
        `https://graph.microsoft.com/v1.0/` +
        `sites/${SITE_ID}/lists/` +
        `${listId}/items` +
        `?$expand=fields&$top=200`;

    while (url) {
        const response =
            await axios.get(url, {
                headers: {
                    Authorization: `Bearer ${token}`
                }
            });

        const found =
            (response.data.value || [])
                .some(item => {
                    const value =
                        item.fields?.[
                            ticketColumn.name
                        ];

                    return (
                        normalizeTicket(value) ===
                        target
                    );
                });

        if (found) {
            return true;
        }

        url =
            response.data['@odata.nextLink'] ||
            null;
    }

    return false;
}

// ============================================================
// HEALTH
// ============================================================

app.get('/', (req, res) => {
    res.json({
        status: 'online',
        timestamp:
            new Date().toISOString(),

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
});

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

            return res.json({
                success: true,
                ticketNumber: ticket,
                existsInPdf,
                existsInList
            });

        } catch (error) {
            console.error(
                '❌ check-status:',
                getErrorMessage(error)
            );

            return res.status(500).json({
                success: false,
                error:
                    getErrorMessage(error)
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
            } = req.body;

            if (!fileName || !fileBase64) {
                return res.status(400).json({
                    success: false,
                    error:
                        'fileName e fileBase64 são obrigatórios.'
                });
            }

            console.log(
                `📄 Enviando PDF: ${fileName}`
            );

            console.log(
                `🎫 Ticket: ${ticketNumber}`
            );

            const token =
                await getAccessToken();

            const driveId =
                await getDriveId(token);

            const folder =
                getEncodedFolderPath();

            const encodedFile =
                encodeURIComponent(fileName);

            const url =
                folder
                    ?
                    `https://graph.microsoft.com/v1.0/` +
                    `drives/${driveId}/root:/` +
                    `${folder}/${encodedFile}:/content`
                    :
                    `https://graph.microsoft.com/v1.0/` +
                    `drives/${driveId}/root:/` +
                    `${encodedFile}:/content`;

            const base64 =
                String(fileBase64).replace(
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
                success: true,

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
            console.error(
                '❌ Erro PDF:',
                getErrorMessage(error)
            );

            return res.status(500).json({
                success: false,
                error:
                    getErrorMessage(error)
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
            } = req.body;

            if (
                !Array.isArray(listData) ||
                listData.length === 0
            ) {
                return res.status(400).json({
                    success: false,
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
                await getListId(token);

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

                console.log(
                    `📋 Linha ${index + 1}/` +
                    `${listData.length}`
                );

                console.log(
                    '📋 Campos:',
                    Object.keys(fields)
                        .join(', ')
                );

                // Mostra tipo real das colunas.
                for (
                    const fieldName
                    of Object.keys(fields)
                ) {
                    const column =
                        columns.find(
                            c =>
                                c.name ===
                                fieldName
                        );

                    let type =
                        'desconhecido';

                    if (column?.text) {
                        type = 'text';
                    } else if (column?.number) {
                        type = 'number';
                    } else if (column?.dateTime) {
                        type = 'dateTime';
                    } else if (
                        column?.hyperlinkOrPicture
                    ) {
                        type =
                            'hyperlinkOrPicture';
                    } else if (column?.choice) {
                        type = 'choice';
                    }

                    console.log(
                        `   ${fieldName}: ${type}`
                    );
                }

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
                        `ID: ${response.data.id}`
                    );

                } catch (error) {
                    const graph =
                        error?.response?.data;

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
                        `${error?.response?.status}`
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
                        '❌ Graph:'
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

                    return res.status(500).json({
                        success: false,

                        error:
                            graph?.error?.message ||
                            error.message,

                        row:
                            index + 1,

                        graphCode:
                            graph?.error?.code,

                        requestId:
                            graph?.error
                                ?.innerError
                                ?.[
                                    'request-id'
                                ]
                    });
                }
            }

            console.log(
                `✅ ${inserted} linha(s) ` +
                `inserida(s) para ${ticketNumber}.`
            );

            return res.json({
                success: true,
                ticketNumber,
                inserted
            });

        } catch (error) {
            console.error(
                '❌ Erro upload lista:',
                getErrorMessage(error)
            );

            return res.status(500).json({
                success: false,
                error:
                    getErrorMessage(error)
            });
        }
    }
);

// ============================================================
// DELETE PDFs DO TICKET
// ============================================================

app.delete(
    '/delete-pdf-by-ticket-number/:ticketNumber',
    async (req, res) => {
        try {
            const ticket =
                normalizeTicket(
                    req.params.ticketNumber
                );

            const token =
                await getAccessToken();

            const driveId =
                await getDriveId(token);

            const files =
                await getAllFilesFromFolder(
                    token,
                    driveId
                );

            const matching =
                files.filter(
                    file =>
                        file.file &&
                        extractTicketNumber(
                            file.name
                        ) === ticket
                );

            let deleted = 0;

            for (const file of matching) {
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

                deleted++;
            }

            return res.json({
                success: true,
                ticketNumber: ticket,
                deleted
            });

        } catch (error) {
            return res.status(500).json({
                success: false,
                error:
                    getErrorMessage(error)
            });
        }
    }
);

// ============================================================
// LIMPAR PDFs DUPLICADOS
// ============================================================

app.post(
    '/cleanup-duplicate-pdfs',
    async (req, res) => {
        try {
            const token =
                await getAccessToken();

            const driveId =
                await getDriveId(token);

            const files =
                await getAllFilesFromFolder(
                    token,
                    driveId
                );

            const groups =
                new Map();

            for (const file of files) {
                if (!file.file) {
                    continue;
                }

                const ticket =
                    extractTicketNumber(
                        file.name
                    );

                if (!ticket) {
                    continue;
                }

                if (!groups.has(ticket)) {
                    groups.set(
                        ticket,
                        []
                    );
                }

                groups
                    .get(ticket)
                    .push(file);
            }

            const deletedFiles = [];
            const keptFiles = [];

            for (
                const [
                    ticket,
                    ticketFiles
                ]
                of groups
            ) {
                if (
                    ticketFiles.length <= 1
                ) {
                    continue;
                }

                ticketFiles.sort(
                    (a, b) =>
                        new Date(
                            b.lastModifiedDateTime ||
                            b.createdDateTime
                        ).getTime() -
                        new Date(
                            a.lastModifiedDateTime ||
                            a.createdDateTime
                        ).getTime()
                );

                const keep =
                    ticketFiles[0];

                keptFiles.push({
                    ticket,
                    fileName:
                        keep.name
                });

                for (
                    const file
                    of ticketFiles.slice(1)
                ) {
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

                    deletedFiles.push({
                        ticket,
                        fileName:
                            file.name
                    });

                    console.log(
                        `🗑️ Duplicado removido: ` +
                        `${file.name}`
                    );
                }
            }

            return res.json({
                success: true,

                totalFilesChecked:
                    files.length,

                ticketsChecked:
                    groups.size,

                deletedCount:
                    deletedFiles.length,

                keptFiles,
                deletedFiles
            });

        } catch (error) {
            console.error(
                '❌ Limpeza:',
                getErrorMessage(error)
            );

            return res.status(500).json({
                success: false,
                error:
                    getErrorMessage(error)
            });
        }
    }
);

// ============================================================
// LIMPAR LISTA COMPLETA
// ============================================================
//
// CUIDADO:
// esta rota continua existindo para compatibilidade.
// Ela apaga TODOS os registros da lista.
// Não é usada pela limpeza de PDFs duplicados.
//
// ============================================================

app.delete(
    '/clear-list',
    async (req, res) => {
        try {
            const token =
                await getAccessToken();

            const listId =
                await getListId(token);

            let deleted = 0;

            while (true) {
                const response =
                    await axios.get(
                        `https://graph.microsoft.com/v1.0/` +
                        `sites/${SITE_ID}/lists/` +
                        `${listId}/items?$top=200`,
                        {
                            headers: {
                                Authorization:
                                    `Bearer ${token}`
                            }
                        }
                    );

                const items =
                    response.data.value || [];

                if (!items.length) {
                    break;
                }

                for (const item of items) {
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

                    deleted++;
                }
            }

            return res.json({
                success: true,
                deleted
            });

        } catch (error) {
            return res.status(500).json({
                success: false,
                error:
                    getErrorMessage(error)
            });
        }
    }
);

// ============================================================
// 404
// ============================================================

app.use((req, res) => {
    res.status(404).json({
        success: false,
        error:
            `Rota não encontrada: ` +
            `${req.method} ${req.originalUrl}`
    });
});

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
