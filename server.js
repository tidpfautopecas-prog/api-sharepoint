import 'dotenv/config';
import express from 'express';
import cors from 'cors';
import axios from 'axios';

// ============================================================
// APLICAÇÃO
// ============================================================

const app = express();

const PORT =
    process.env.PORT || 3000;

// ============================================================
// VARIÁVEIS DO RENDER
// ============================================================

const TENANT_ID =
    process.env.TENANT_ID;

const CLIENT_ID =
    process.env.CLIENT_ID;

const CLIENT_SECRET =
    process.env.CLIENT_SECRET;

const SITE_ID =
    process.env.SITE_ID;

const LIBRARY_NAME =
    process.env.LIBRARY_NAME;

const LIST_NAME =
    process.env.LIST_NAME;

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

let cachedAccessToken = null;
let cachedAccessTokenExpiresAt = 0;

let cachedDriveId = null;
let cachedListId = null;
let cachedListColumns = null;

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

    const requiredVariables = [
        'TENANT_ID',
        'CLIENT_ID',
        'CLIENT_SECRET',
        'SITE_ID',
        'LIBRARY_NAME',
        'LIST_NAME'
    ];

    const missing =
        requiredVariables.filter(
            variable =>
                !process.env[variable]
        );

    if (missing.length > 0) {

        console.warn(
            '⚠️ Variáveis de ambiente não configuradas:',
            missing.join(', ')
        );

        return;
    }

    console.log(
        '✅ Variáveis de ambiente principais configuradas.'
    );
}

validateEnvironment();

// ============================================================
// TRATAMENTO DE ERROS
// ============================================================

function getErrorMessage(error) {

    return (
        error?.response?.data?.error_description ||
        error?.response?.data?.error?.message ||
        error?.response?.data?.message ||
        error?.response?.data?.error ||
        error?.message ||
        'Erro desconhecido.'
    );
}

// ============================================================
// NORMALIZA TEXTO
// ============================================================

function normalizeText(value) {

    if (
        value === undefined ||
        value === null
    ) {

        return '';
    }

    return String(value)
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
// AUTENTICAÇÃO MICROSOFT
// ============================================================

async function getAccessToken() {

    try {

        // ----------------------------------------------------
        // Reutiliza token ainda válido
        // ----------------------------------------------------

        if (
            cachedAccessToken &&
            Date.now() <
                cachedAccessTokenExpiresAt
        ) {

            return cachedAccessToken;
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

        if (
            !response.data.access_token
        ) {

            throw new Error(
                'Microsoft não retornou access_token.'
            );
        }

        cachedAccessToken =
            response.data.access_token;

        const expiresIn =
            Number(
                response.data.expires_in || 3600
            );

        // Renova cinco minutos antes
        // do vencimento real.

        cachedAccessTokenExpiresAt =
            Date.now() +
            Math.max(
                expiresIn - 300,
                60
            ) * 1000;

        console.log(
            '✅ Autenticação Microsoft realizada.'
        );

        return cachedAccessToken;

    } catch (error) {

        const message =
            getErrorMessage(error);

        console.error(
            '❌ Erro de autenticação:',
            message
        );

        throw new Error(
            `Erro na autenticação Microsoft: ${message}`
        );
    }
}

// ============================================================
// SITE
// ============================================================

function getSiteId() {

    if (!SITE_ID) {

        throw new Error(
            'Variável SITE_ID não configurada no Render.'
        );
    }

    return SITE_ID;
}

// ============================================================
// LOCALIZA BIBLIOTECA PELO LIBRARY_NAME
// ============================================================

async function getDriveId(
    accessToken
) {

    try {

        if (cachedDriveId) {

            return cachedDriveId;
        }

        if (!LIBRARY_NAME) {

            throw new Error(
                'Variável LIBRARY_NAME não configurada.'
            );
        }

        const siteId =
            getSiteId();

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

        const targetName =
            normalizeText(
                LIBRARY_NAME
            );

        const drive =
            drives.find(
                item =>
                    normalizeText(
                        item.name
                    ) === targetName
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
                `Bibliotecas disponíveis: ${available}`
            );
        }

        cachedDriveId =
            drive.id;

        console.log(
            `✅ Biblioteca localizada: ${drive.name}`
        );

        console.log(
            `📚 Drive ID localizado com sucesso.`
        );

        return cachedDriveId;

    } catch (error) {

        const message =
            getErrorMessage(error);

        console.error(
            '❌ Erro ao localizar biblioteca:',
            message
        );

        throw new Error(
            `Erro ao localizar biblioteca: ${message}`
        );
    }
}

// ============================================================
// LOCALIZA LISTA PELO LIST_NAME
// ============================================================

async function getListId(
    accessToken
) {

    try {

        if (cachedListId) {

            return cachedListId;
        }

        if (!LIST_NAME) {

            throw new Error(
                'Variável LIST_NAME não configurada.'
            );
        }

        const siteId =
            getSiteId();

        let url =
            `https://graph.microsoft.com/v1.0/` +
            `sites/${siteId}/lists` +
            `?$select=id,name,displayName&$top=200`;

        const lists = [];

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

            lists.push(
                ...(response.data.value || [])
            );

            url =
                response.data[
                    '@odata.nextLink'
                ] || null;
        }

        const targetName =
            normalizeText(
                LIST_NAME
            );

        const list =
            lists.find(
                item => {

                    const displayName =
                        normalizeText(
                            item.displayName
                        );

                    const name =
                        normalizeText(
                            item.name
                        );

                    return (
                        displayName ===
                            targetName ||
                        name ===
                            targetName
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
                `Listas disponíveis: ${available}`
            );
        }

        cachedListId =
            list.id;

        console.log(
            `✅ Lista localizada: ` +
            `${list.displayName || list.name}`
        );

        console.log(
            '📋 List ID localizado com sucesso.'
        );

        return cachedListId;

    } catch (error) {

        const message =
            getErrorMessage(error);

        console.error(
            '❌ Erro ao localizar lista:',
            message
        );

        throw new Error(
            `Erro ao localizar lista: ${message}`
        );
    }
}

// ============================================================
// COLUNAS DA LISTA
// ============================================================

async function getListColumns(
    accessToken,
    listId
) {

    try {

        if (cachedListColumns) {

            return cachedListColumns;
        }

        const siteId =
            getSiteId();

        let url =
            `https://graph.microsoft.com/v1.0/` +
            `sites/${siteId}/lists/` +
            `${listId}/columns` +
            `?$select=id,name,displayName,hidden,readOnly&$top=200`;

        const columns = [];

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

            columns.push(
                ...(response.data.value || [])
            );

            url =
                response.data[
                    '@odata.nextLink'
                ] || null;
        }

        cachedListColumns =
            columns;

        console.log(
            `✅ ${columns.length} coluna(s) da lista carregada(s).`
        );

        return columns;

    } catch (error) {

        const message =
            getErrorMessage(error);

        throw new Error(
            `Erro ao consultar colunas da lista: ${message}`
        );
    }
}

// ============================================================
// LOCALIZA UMA COLUNA
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

            if (column.readOnly) {

                return false;
            }

            const displayName =
                normalizeText(
                    column.displayName
                );

            const internalName =
                normalizeText(
                    column.name
                );

            return (
                normalizedAliases.includes(
                    displayName
                ) ||
                normalizedAliases.includes(
                    internalName
                )
            );
        }
    );
}

// ============================================================
// ATRIBUI CAMPO SE A COLUNA EXISTIR
// ============================================================

function setSharePointField(
    fields,
    columns,
    aliases,
    value,
    allowEmpty = false
) {

    if (
        !allowEmpty &&
        (
            value === undefined ||
            value === null ||
            String(value).trim() === ''
        )
    ) {

        return false;
    }

    const column =
        findColumn(
            columns,
            aliases
        );

    if (!column) {

        return false;
    }

    fields[column.name] =
        value;

    return true;
}

// ============================================================
// CAMINHO DA PASTA
// ============================================================

function getEncodedFolderPath() {

    return String(
        FOLDER_PATH || ''
    )
        .replace(/^\/+/, '')
        .replace(/\/+$/, '')
        .split('/')
        .filter(Boolean)
        .map(
            part =>
                encodeURIComponent(part)
        )
        .join('/');
}

// ============================================================
// ARQUIVOS DA PASTA
// ============================================================

async function getAllFilesFromFolder(
    accessToken,
    driveId
) {

    const files = [];

    const encodedFolder =
        getEncodedFolderPath();

    let url;

    if (encodedFolder) {

        url =
            `https://graph.microsoft.com/v1.0/` +
            `drives/${driveId}/root:/` +
            `${encodedFolder}:/children` +
            `?$top=200&` +
            `$select=id,name,file,folder,size,` +
            `createdDateTime,lastModifiedDateTime,webUrl`;

    } else {

        url =
            `https://graph.microsoft.com/v1.0/` +
            `drives/${driveId}/root/children` +
            `?$top=200&` +
            `$select=id,name,file,folder,size,` +
            `createdDateTime,lastModifiedDateTime,webUrl`;
    }

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

        files.push(
            ...(response.data.value || [])
        );

        url =
            response.data[
                '@odata.nextLink'
            ] || null;
    }

    return files;
}

// ============================================================
// EXTRAI TICKET DO NOME DO PDF
// ============================================================
//
// Exemplos:
//
// Laudo - SR-7382-17913948365.pdf
// Laudo - SR-7382-20261007 1426.pdf
// Laudo - SR-7382.pdf
//
// Retorna:
//
// SR-7382
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

    // SR-1234

    let match =
        name.match(
            /^Laudo\s*-\s*(SR-\d+)(?:-.+)?\.pdf$/i
        );

    if (match) {

        return normalizeTicket(
            match[1]
        );
    }

    // Outros prefixos:
    // OS-123, TK-123 etc.

    match =
        name.match(
            /^Laudo\s*-\s*([A-Za-z]+-\d+)(?:-.+)?\.pdf$/i
        );

    if (match) {

        return normalizeTicket(
            match[1]
        );
    }

    // Ticket apenas numérico

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
// VERIFICA PDF
// ============================================================

async function ticketPdfExists(
    accessToken,
    driveId,
    ticketNumber
) {

    const targetTicket =
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
                fileTicket ===
                targetTicket
            );
        }
    );
}

// ============================================================
// LOCALIZA COLUNA DO TICKET
// ============================================================

function getTicketColumn(
    columns
) {

    return findColumn(
        columns,
        [
            'N° do ticket',
            'Nº do ticket',
            'Número do ticket',
            'Numero do ticket',
            'N_x00b0_doticket',
            'N_x00b0__x0020_do_x0020_ticket',
            'N_x00b0_do_x0020_ticket',
            'NumeroTicket',
            'Ticket',
            'Title'
        ]
    );
}

// ============================================================
// TICKET DE UM ITEM DA LISTA
// ============================================================

function getTicketFromListFields(
    fields,
    ticketColumn
) {

    if (!fields) {

        return '';
    }

    if (
        ticketColumn &&
        fields[ticketColumn.name] !==
            undefined
    ) {

        return normalizeTicket(
            fields[ticketColumn.name]
        );
    }

    const fallbackNames = [
        'N_x00b0_doticket',
        'N_x00b0__x0020_do_x0020_ticket',
        'N_x00b0_do_x0020_ticket',
        'NumeroTicket',
        'Ticket',
        'Title'
    ];

    for (
        const fieldName
        of fallbackNames
    ) {

        if (
            fields[fieldName] !==
                undefined &&
            fields[fieldName] !==
                null
        ) {

            const value =
                normalizeTicket(
                    fields[fieldName]
                );

            if (value) {

                return value;
            }
        }
    }

    return '';
}

// ============================================================
// VERIFICA TICKET NA LISTA
// ============================================================

async function ticketExistsInList(
    accessToken,
    listId,
    columns,
    ticketNumber
) {

    const siteId =
        getSiteId();

    const targetTicket =
        normalizeTicket(
            ticketNumber
        );

    const ticketColumn =
        getTicketColumn(
            columns
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
                        item.fields,
                        ticketColumn
                    ) ===
                    targetTicket
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
// BUSCA VALOR EM OBJETO
// ============================================================

function getRowValue(
    row,
    possibleNames
) {

    for (
        const name
        of possibleNames
    ) {

        if (
            row[name] !== undefined &&
            row[name] !== null
        ) {

            return row[name];
        }
    }

    return '';
}

// ============================================================
// MONTA CAMPOS DA LISTA
// ============================================================

function buildSharePointFields(
    row,
    ticketNumber,
    columns
) {

    const fields = {};

    // --------------------------------------------------------
    // TICKET
    // --------------------------------------------------------

    const ticket =
        getRowValue(
            row,
            [
                'N° do ticket',
                'Nº do ticket',
                'Número do ticket',
                'Numero do ticket',
                'NumeroTicket',
                'Ticket',
                'ticketNumber'
            ]
        ) ||
        ticketNumber ||
        '';

    setSharePointField(
        fields,
        columns,
        [
            'N° do ticket',
            'Nº do ticket',
            'Número do ticket',
            'Numero do ticket',
            'N_x00b0_doticket',
            'N_x00b0__x0020_do_x0020_ticket',
            'N_x00b0_do_x0020_ticket',
            'NumeroTicket',
            'Ticket'
        ],
        String(ticket)
    );

    // --------------------------------------------------------
    // TITLE
    // --------------------------------------------------------

    setSharePointField(
        fields,
        columns,
        [
            'Title',
            'Título',
            'Titulo'
        ],
        String(ticket)
    );

    // --------------------------------------------------------
    // CLIENTE
    // --------------------------------------------------------

    const cliente =
        getRowValue(
            row,
            [
                'Nome do Cliente',
                'Nome do cliente',
                'NomeCliente',
                'Cliente'
            ]
        );

    setSharePointField(
        fields,
        columns,
        [
            'Nome do Cliente',
            'Nome do cliente',
            'NomedoCliente',
            'NomeCliente',
            'Cliente'
        ],
        cliente
    );

    // --------------------------------------------------------
    // ITEM
    // --------------------------------------------------------

    const item =
        getRowValue(
            row,
            [
                'Item',
                'item'
            ]
        );

    setSharePointField(
        fields,
        columns,
        [
            'Item'
        ],
        item
    );

    // --------------------------------------------------------
    // QUANTIDADE
    // --------------------------------------------------------

    const quantidadeRaw =
        getRowValue(
            row,
            [
                'Qtde',
                'Quantidade',
                'quantidade'
            ]
        );

    if (
        quantidadeRaw !== '' &&
        quantidadeRaw !== undefined &&
        quantidadeRaw !== null
    ) {

        const quantidade =
            Number(
                String(
                    quantidadeRaw
                )
                    .replace(
                        ',',
                        '.'
                    )
            );

        setSharePointField(
            fields,
            columns,
            [
                'Qtde',
                'Quantidade'
            ],
            Number.isNaN(
                quantidade
            )
                ? quantidadeRaw
                : quantidade,
            true
        );
    }

    // --------------------------------------------------------
    // MOTIVO
    // --------------------------------------------------------

    let motivo =
        getRowValue(
            row,
            [
                'Motivo',
                'motivo'
            ]
        );

    if (
        Array.isArray(
            motivo
        )
    ) {

        motivo =
            motivo.join(', ');
    }

    setSharePointField(
        fields,
        columns,
        [
            'Motivo'
        ],
        motivo
    );

    // --------------------------------------------------------
    // ORIGEM DO DEFEITO
    // --------------------------------------------------------

    const origemDefeito =
        getRowValue(
            row,
            [
                'Origem do defeito',
                'Origem do Defeito',
                'OrigemDefeito'
            ]
        );

    setSharePointField(
        fields,
        columns,
        [
            'Origem do defeito',
            'Origem do Defeito',
            'Origemdodefeito',
            'OrigemDefeito'
        ],
        origemDefeito
    );

    // --------------------------------------------------------
    // DISPOSIÇÃO
    // --------------------------------------------------------

    const disposicao =
        getRowValue(
            row,
            [
                'Disposição',
                'Disposicao'
            ]
        );

    setSharePointField(
        fields,
        columns,
        [
            'Disposição',
            'Disposicao',
            'Disposi_x00e7__x00e3_o'
        ],
        disposicao
    );

    // --------------------------------------------------------
    // DISPOSIÇÃO DAS PEÇAS
    // --------------------------------------------------------

    const disposicaoPecas =
        getRowValue(
            row,
            [
                'Disposição das peças',
                'Disposição das Peças',
                'Disposicao das pecas',
                'DisposicaoPecas'
            ]
        );

    setSharePointField(
        fields,
        columns,
        [
            'Disposição das peças',
            'Disposição das Peças',
            'Disposicao das pecas',
            'DisposicaoPecas',
            'Disposi_x00e7__x00e3_odaspe_x00e'
        ],
        disposicaoPecas
    );

    // --------------------------------------------------------
    // DATA DE GERAÇÃO
    // --------------------------------------------------------

    const dataGeracao =
        getRowValue(
            row,
            [
                'Data de Geração',
                'Data de geração',
                'Data de Geracao',
                'DataGeracao'
            ]
        );

    setSharePointField(
        fields,
        columns,
        [
            'Data de Geração',
            'Data de geração',
            'Data de Geracao',
            'DataGeracao',
            'DatadeGera_x00e7__x00e3_o'
        ],
        dataGeracao
    );

    // --------------------------------------------------------
    // FOTOS 1 A 10
    // --------------------------------------------------------

    for (
        let i = 1;
        i <= 10;
        i++
    ) {

        const foto =
            getRowValue(
                row,
                [
                    `Foto${i}`,
                    `Foto ${i}`,
                    `foto${i}`
                ]
            );

        setSharePointField(
            fields,
            columns,
            [
                `Foto${i}`,
                `Foto ${i}`
            ],
            foto
        );
    }

    return fields;
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
                tenantConfigured:
                    Boolean(TENANT_ID),

                clientConfigured:
                    Boolean(CLIENT_ID),

                secretConfigured:
                    Boolean(CLIENT_SECRET),

                siteConfigured:
                    Boolean(SITE_ID),

                listName:
                    LIST_NAME || null,

                libraryName:
                    LIBRARY_NAME || null,

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
                `🔎 Verificando ${ticketNumber}...`
            );

            const accessToken =
                await getAccessToken();

            const [
                driveId,
                listId
            ] =
                await Promise.all([
                    getDriveId(
                        accessToken
                    ),
                    getListId(
                        accessToken
                    )
                ]);

            const columns =
                await getListColumns(
                    accessToken,
                    listId
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
                        listId,
                        columns,
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
// UPLOAD DO PDF
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

            if (!fileName) {

                return res
                    .status(400)
                    .json({
                        success: false,
                        error:
                            'fileName não informado.'
                    });
            }

            if (!fileBase64) {

                return res
                    .status(400)
                    .json({
                        success: false,
                        error:
                            'fileBase64 não informado.'
                    });
            }

            console.log(
                `📄 Enviando PDF: ${fileName}`
            );

            if (ticketNumber) {

                console.log(
                    `🎫 Ticket: ${ticketNumber}`
                );
            }

            const accessToken =
                await getAccessToken();

            const driveId =
                await getDriveId(
                    accessToken
                );

            const encodedFolder =
                getEncodedFolderPath();

            const encodedFileName =
                encodeURIComponent(
                    fileName
                );

            let uploadUrl;

            if (encodedFolder) {

                uploadUrl =
                    `https://graph.microsoft.com/v1.0/` +
                    `drives/${driveId}/root:/` +
                    `${encodedFolder}/` +
                    `${encodedFileName}:/content`;

            } else {

                uploadUrl =
                    `https://graph.microsoft.com/v1.0/` +
                    `drives/${driveId}/root:/` +
                    `${encodedFileName}:/content`;
            }

            // Aceita Base64 puro ou
            // data:application/pdf;base64,...

            const cleanBase64 =
                String(fileBase64)
                    .replace(
                        /^data:application\/pdf;base64,/i,
                        ''
                    );

            const buffer =
                Buffer.from(
                    cleanBase64,
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
                '❌ Erro upload PDF:',
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
// UPLOAD DOS DADOS PARA A LISTA
// ============================================================

app.post(
    '/upload-list-data',
    async (req, res) => {

        try {

            const ticketNumber =
                req.body.ticketNumber ||
                req.body.ticket ||
                '';

            const listData =
                req.body.listData ||
                req.body.rows ||
                req.body.data;

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
                            'listData não informado ou vazio.'
                    });
            }

            console.log(
                `📋 Enviando ticket ${ticketNumber} ` +
                `para a lista...`
            );

            const accessToken =
                await getAccessToken();

            const listId =
                await getListId(
                    accessToken
                );

            const columns =
                await getListColumns(
                    accessToken,
                    listId
                );

            const siteId =
                getSiteId();

            let inserted = 0;

            const insertedItems = [];

            for (
                const row
                of listData
            ) {

                const fields =
                    buildSharePointFields(
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
                        'Nenhuma coluna compatível foi encontrada na lista SharePoint.'
                    );
                }

                console.log(
                    '📋 Campos que serão enviados:',
                    Object.keys(fields)
                        .join(', ')
                );

                const url =
                    `https://graph.microsoft.com/v1.0/` +
                    `sites/${siteId}/lists/` +
                    `${listId}/items`;

                const response =
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

                insertedItems.push(
                    response.data.id
                );
            }

            console.log(
                `✅ Lista atualizada. ` +
                `${inserted} linha(s) inserida(s).`
            );

            return res.json({
                success: true,

                ticketNumber,

                inserted,

                itemIds:
                    insertedItems
            });

        } catch (error) {

            const message =
                getErrorMessage(error);

            console.error(
                '❌ Erro upload lista:',
                message
            );

            if (
                error?.response?.data
            ) {

                console.error(
                    '❌ Retorno Graph:',
                    JSON.stringify(
                        error.response.data
                    )
                );
            }

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

            const driveId =
                await getDriveId(
                    accessToken
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

            const deletedFiles = [];

            for (
                const file
                of matchingFiles
            ) {

                const url =
                    `https://graph.microsoft.com/v1.0/` +
                    `drives/${driveId}/items/` +
                    `${file.id}`;

                await axios.delete(
                    url,
                    {
                        headers: {
                            Authorization:
                                `Bearer ${accessToken}`
                        }
                    }
                );

                deleted++;

                deletedFiles.push(
                    file.name
                );

                console.log(
                    `🗑️ PDF excluído: ${file.name}`
                );
            }

            return res.json({
                success: true,
                ticketNumber,
                deleted,
                deletedFiles
            });

        } catch (error) {

            const message =
                getErrorMessage(error);

            console.error(
                '❌ Erro exclusão PDF:',
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
// LIMPA PDFs DUPLICADOS
// ============================================================
//
// IMPORTANTE:
//
// Esta função NÃO exclui registros da lista.
//
// Ela somente verifica PDFs na biblioteca.
//
// Para cada ticket mantém o arquivo
// mais recentemente modificado.
//
// ============================================================

app.post(
    '/cleanup-duplicate-pdfs',
    async (req, res) => {

        try {

            console.log(
                '🧹 Iniciando verificação de PDFs duplicados...'
            );

            const accessToken =
                await getAccessToken();

            const driveId =
                await getDriveId(
                    accessToken
                );

            const files =
                await getAllFilesFromFolder(
                    accessToken,
                    driveId
                );

            console.log(
                `📂 ${files.length} arquivo(s) localizado(s).`
            );

            const ticketGroups =
                new Map();

            for (
                const file
                of files
            ) {

                // Ignora pastas

                if (!file.file) {

                    continue;
                }

                // Ignora arquivos não PDF

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

                // Arquivo não reconhecido.
                // Não toca nele.

                if (!ticketNumber) {

                    console.log(
                        `ℹ️ PDF ignorado: ${file.name}`
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
            const keptFiles = [];
            const deletedFiles = [];

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
                    `⚠️ ${ticketNumber}: ` +
                    `${ticketFiles.length} PDFs encontrados.`
                );

                // Ordena do mais novo
                // para o mais antigo.

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

                keptFiles.push({
                    ticketNumber,
                    fileName:
                        keepFile.name,

                    modified:
                        keepFile
                            .lastModifiedDateTime
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

                console.log(
                    `✅ ${ticketNumber}: mantendo ${keepFile.name}`
                );

                for (
                    const file
                    of filesToDelete
                ) {

                    const url =
                        `https://graph.microsoft.com/v1.0/` +
                        `drives/${driveId}/items/` +
                        `${file.id}`;

                    await axios.delete(
                        url,
                        {
                            headers: {
                                Authorization:
                                    `Bearer ${accessToken}`
                            }
                        }
                    );

                    deletedFiles.push({
                        ticketNumber,
                        fileName:
                            file.name
                    });

                    console.log(
                        `🗑️ Duplicado removido: ${file.name}`
                    );
                }
            }

            console.log(
                `✅ Limpeza concluída. ` +
                `${deletedFiles.length} arquivo(s) removido(s).`
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
                '❌ Erro limpeza de duplicados:',
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
// Esta rota APAGA TODOS OS ITENS
// da lista SharePoint.
//
// Ela NÃO é chamada pela limpeza
// de PDFs duplicados.
//
// ============================================================

app.delete(
    '/clear-list',
    async (req, res) => {

        try {

            console.log(
                '⚠️ Iniciando exclusão TOTAL da lista...'
            );

            const accessToken =
                await getAccessToken();

            const listId =
                await getListId(
                    accessToken
                );

            const siteId =
                getSiteId();

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
                    `🗑️ ${deleted} registro(s) removido(s)...`
                );
            }

            return res.json({
                success: true,
                deleted,

                message:
                    `${deleted} registro(s) removido(s).`
            });

        } catch (error) {

            const message =
                getErrorMessage(error);

            console.error(
                '❌ Erro clear-list:',
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
            `🔗 SITE_ID configurado: ${Boolean(SITE_ID)}`
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
