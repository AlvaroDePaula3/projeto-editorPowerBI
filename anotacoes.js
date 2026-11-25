/* =========================================================
   Anotações Power BI
   Cabeçalho e corpo são gerados pela MESMA lista (COLUNAS),
   então nunca desalinham.
   ========================================================= */

const corpoTabela      = document.getElementById('corpoTabela');
const cabecalhoTabela  = document.getElementById('cabecalhoTabela');
const inputArquivo     = document.getElementById('botaoPowerBI');
const statusImportacao = document.getElementById('statusImportacao');
const divisorio        = document.getElementById('divisorioTabela');
const barraFixa        = document.getElementById('barraRolagemFixa');
const barraFixaConteudo = document.getElementById('barraRolagemFixaConteudo');

/* ---------------------------------------------------------
   1) Definição única das colunas
   tipo: 'texto'  -> valor vindo da planilha
         'input'  -> input de texto editável
         'data'   -> input type="date" editável
         'checks' -> 3 checkboxes
   --------------------------------------------------------- */
const COLUNAS = [
    { chave: 'Sequencia',            tipo: 'texto' },
    { chave: 'SequenciaRaiz',        tipo: 'texto' },
    { chave: 'RAIZ',                 tipo: 'texto' },
    { chave: 'Prioridade',           tipo: 'texto' },
    { chave: 'Aging',                tipo: 'texto' },
    { chave: 'Categoria',            tipo: 'texto' },
    { chave: 'Classificação',        tipo: 'texto' },
    { chave: 'Grupo Responsável',    tipo: 'texto' },
    { chave: 'GRUPO_PLANTA_INTERNA', tipo: 'texto' },
    { chave: 'Localidade',           tipo: 'texto' },
    { chave: 'AfetFTTX',             tipo: 'texto' },
    { chave: 'Gestor Atual',         tipo: 'texto' },
    { chave: 'UF',                   tipo: 'texto' },
    { chave: 'SiglaLocalidade',      tipo: 'texto' },
    { chave: 'Sigla Site',           tipo: 'texto' },
    { chave: 'AfetIPTV',             tipo: 'texto' },
    { chave: 'AfetVoz',              tipo: 'texto' },

    /* ---- editáveis, inseridas ANTES de DataCriacao ---- */
    { chave: 'STATUS',               tipo: 'input',  campo: 'status',   placeholder: 'Digite o status' },
    { chave: 'Previsão',             tipo: 'input',  campo: 'previsao', placeholder: 'Digite a previsão' },
    { chave: 'Área_Responsável',     tipo: 'checks', campo: 'area' },
    /* --------------------------------------------------- */

    { chave: 'DataCriacao',          tipo: 'data',   campo: 'dataCriacao' },
    { chave: 'DataApresentacao',     tipo: 'texto' },
    { chave: 'TipoRede',             tipo: 'texto' },
    { chave: 'TpRedeRAIZ',           tipo: 'texto' },
    { chave: 'TipoPlanta',           tipo: 'texto' },
    { chave: 'TipoFalha',            tipo: 'texto' },
    { chave: 'TipoAlarme',           tipo: 'texto' },
    { chave: 'TipodeBilhete',        tipo: 'texto' },
    { chave: 'TipoAfetacao',         tipo: 'texto' },
    { chave: 'LocalidadeRAIZ',       tipo: 'texto' },
    { chave: 'CodLocalidadeRAIZ',    tipo: 'texto' },
    { chave: 'SEGS_DURACAO',         tipo: 'texto' },
    { chave: 'QTD_DIAS',             tipo: 'texto' },
    { chave: 'REGIONAL',             tipo: 'texto' },
    { chave: 'Faixa_dias',           tipo: 'texto' },
    { chave: 'Alarme',               tipo: 'texto' },
    { chave: 'B2B',                  tipo: 'texto' }
];


//   2) Normalização de nomes de coluna

function normalizar(texto) {
    return String(texto ?? '')
        .normalize('NFD')
        .replace(/[\u0300-\u036f]/g, '')
        .replace(/\s+/g, '')
        .toLowerCase();
}

function pegarValor(linha, colunaEsperada) {
    const alvo = normalizar(colunaEsperada);
    const chaveReal = Object.keys(linha).find(k => normalizar(k) === alvo);
    return chaveReal !== undefined ? (linha[chaveReal] ?? '') : '';
}

/* ---------------------------------------------------------
   3) Datas do Excel -> yyyy-mm-dd (formato do input date)
   --------------------------------------------------------- */
function serialParaISO(valor) {
    if (valor === '' || valor === null || valor === undefined) return '';

    if (typeof valor === 'number') {
        const d = XLSX.SSF.parse_date_code(valor);
        if (!d) return '';
        return `${d.y}-${String(d.m).padStart(2, '0')}-${String(d.d).padStart(2, '0')}`;
    }

    const texto = String(valor).trim();

    // dd/mm/aaaa
    let m = texto.match(/^(\d{1,2})\/(\d{1,2})\/(\d{4})$/);
    if (m) {
        return `${m[3]}-${m[2].padStart(2, '0')}-${m[1].padStart(2, '0')}`;
    }

    // aaaa-mm-dd (já no formato)
    m = texto.match(/^(\d{4})-(\d{1,2})-(\d{1,2})/);
    if (m) {
        return `${m[1]}-${m[2].padStart(2, '0')}-${m[3].padStart(2, '0')}`;
    }

    return '';
}

/* ---------------------------------------------------------
   4) Montagem do cabeçalho e das células
   --------------------------------------------------------- */
function montarCabecalho() {
    cabecalhoTabela.innerHTML = '';
    const tr = document.createElement('tr');

    COLUNAS.forEach(coluna => {
        const th = document.createElement('th');
        th.textContent = coluna.chave;
        if (coluna.tipo !== 'texto') th.classList.add('col-edicao');
        tr.appendChild(th);
    });

    cabecalhoTabela.appendChild(tr);
}

function criarCelula(coluna, dados) {
    const td = document.createElement('td');

    switch (coluna.tipo) {

        case 'input': {
            const input = document.createElement('input');
            input.type = 'text';
            input.className = 'campo-texto';
            input.dataset.campo = coluna.campo;
            input.placeholder = coluna.placeholder || '';
            td.appendChild(input);
            break;
        }

        case 'data': {
            const input = document.createElement('input');
            input.type = 'date';
            input.className = 'campo-texto';
            input.dataset.campo = coluna.campo;
            input.value = serialParaISO(pegarValor(dados, coluna.chave));
            td.appendChild(input);
            break;
        }

        case 'checks': {
            const div = document.createElement('div');
            div.className = 'grupo-checkbox';

            [
                ['falha_interna', 'Falha interna'],
                ['falha_externa', 'Falha externa'],
                ['outros',        'Outros']
            ].forEach(([valor, rotulo]) => {
                const label = document.createElement('label');

                const cb = document.createElement('input');
                cb.type = 'checkbox';
                cb.dataset.campo = coluna.campo;
                cb.value = valor;

                label.append(cb, document.createTextNode(' ' + rotulo));
                div.appendChild(label);
            });

            td.appendChild(div);
            break;
        }

        default: {
            const valor = pegarValor(dados, coluna.chave);
            td.textContent = valor === 0 || valor ? String(valor) : '';
        }
    }

    return td;
}

function criarLinha(dados) {
    const tr = document.createElement('tr');
    tr.dataset.id = String(pegarValor(dados, 'Sequencia') || crypto.randomUUID());

    COLUNAS.forEach(coluna => tr.appendChild(criarCelula(coluna, dados)));

    return tr;
}

function renderizarTabela(linhas) {
    corpoTabela.innerHTML = '';

    const fragmento = document.createDocumentFragment();
    linhas.forEach(linha => fragmento.appendChild(criarLinha(linha)));
    corpoTabela.appendChild(fragmento);

    sincronizarLargura();
}

/* ---------------------------------------------------------
   5) Barra de rolagem fixa (espelho da tabela)
   --------------------------------------------------------- */
let sincronizandoScroll = false;

function sincronizarScroll(origem, destino) {
    if (sincronizandoScroll) return;
    sincronizandoScroll = true;
    destino.scrollLeft = origem.scrollLeft;
    requestAnimationFrame(() => { sincronizandoScroll = false; });
}

function sincronizarLargura() {
    const tabela = divisorio.querySelector('.tabelaExportada');
    if (!tabela) return;
    barraFixaConteudo.style.width = tabela.scrollWidth + 'px';
}

barraFixa.addEventListener('scroll', () => sincronizarScroll(barraFixa, divisorio));
divisorio.addEventListener('scroll', () => sincronizarScroll(divisorio, barraFixa));
window.addEventListener('resize', sincronizarLargura);

/* ---------------------------------------------------------
   6) Leitura do arquivo .xlsx
   --------------------------------------------------------- */
inputArquivo.addEventListener('change', async (evento) => {
    const arquivo = evento.target.files[0];
    if (!arquivo) return;

    try {
        statusImportacao.textContent = 'Lendo arquivo...';
        statusImportacao.className = 'status-importacao';

        const buffer = await arquivo.arrayBuffer();
        const workbook = XLSX.read(buffer, { type: 'array' });

        if (!workbook.SheetNames.length) {
            throw new Error('O arquivo não tem nenhuma aba.');
        }

        const nomeAba = workbook.SheetNames[0];
        const linhas = XLSX.utils.sheet_to_json(workbook.Sheets[nomeAba], { defval: '' });

        if (linhas.length === 0) {
            statusImportacao.textContent = `A aba "${nomeAba}" está vazia.`;
            statusImportacao.className = 'status-importacao erro';
            return;
        }

        renderizarTabela(linhas);

        const colunasArquivo = Object.keys(linhas[0]);
        const esperadas = COLUNAS.map(c => normalizar(c.chave));
        const naoMapeadas = colunasArquivo.filter(k => !esperadas.includes(normalizar(k)));

        statusImportacao.textContent =
            `${linhas.length} linha(s) carregada(s) da aba "${nomeAba}" — ${colunasArquivo.length} colunas no arquivo.`;
        statusImportacao.className = 'status-importacao ok';

        if (naoMapeadas.length > 0) {
            console.info('Colunas do arquivo não exibidas na tabela:', naoMapeadas);
        }

    } catch (erro) {
        console.error('Falha ao ler o arquivo:', erro);
        statusImportacao.textContent = `Não foi possível ler o arquivo: ${erro.message}`;
        statusImportacao.className = 'status-importacao erro';
    }
});

/* ---------------------------------------------------------
   7) Inicialização
   --------------------------------------------------------- */
document.addEventListener('DOMContentLoaded', () => {
    montarCabecalho();
    sincronizarLargura();
});
