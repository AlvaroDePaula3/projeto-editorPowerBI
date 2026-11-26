/* =========================================================
   Anotações Power BI
   Uma única lista COLUNAS governa cabeçalho, células,
   coleta e exportação. Sem desalinhamento.
   ========================================================= */

const corpoTabela       = document.getElementById('corpoTabela');
const cabecalhoTabela   = document.getElementById('cabecalhoTabela');
const inputArquivo      = document.getElementById('botaoPowerBI');
const statusImportacao  = document.getElementById('statusImportacao');
const divisorio         = document.getElementById('divisorioTabela');
const barraFixa         = document.getElementById('barraRolagemFixa');
const barraFixaConteudo = document.getElementById('barraRolagemFixaConteudo');
const modalAlarme       = document.getElementById('modalAlarme');
const modalCorpo        = document.getElementById('modalCorpo');
const modalFechar       = document.getElementById('modalFechar');
const botaoExportar     = document.getElementById('botaoExportar');

/* ---------------------------------------------------------
   1) Definição das colunas

   tipo: 'texto'  -> somente leitura, vem do arquivo
         'input'  -> editável (texto livre)
         'checks' -> 3 checkboxes
         'data'   -> data, somente leitura
         'modal'  -> botão que abre o texto completo
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

    { chave: 'STATUS',               tipo: 'input',  campo: 'status',   placeholder: 'Digite o status' },
    { chave: 'Previsão',             tipo: 'input',  campo: 'previsao', placeholder: 'Digite a previsão' },
    { chave: 'Área_Responsável',     tipo: 'checks', campo: 'area' },

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
    { chave: 'Alarme',               tipo: 'modal' },
    { chave: 'B2B',                  tipo: 'texto' }
];

/* Linhas originais do arquivo, na ordem. */
let dadosOriginais = [];
/* Cabeçalhos completos do arquivo (inclui colunas não exibidas). */
let cabecalhosArquivo = [];

/* ---------------------------------------------------------
   2) Normalização de nomes de coluna
   --------------------------------------------------------- */
function normalizar(texto) {
    return String(texto ?? '')
        .normalize('NFD')
        .replace(/[\u0300-\u036f]/g, '')
        .replace(/\s+/g, '')
        .toLowerCase();
}

function pegarValor(linha, colunaEsperada) {
    if (!linha) return '';
    const alvo = normalizar(colunaEsperada);
    const chaveReal = Object.keys(linha).find(k => normalizar(k) === alvo);
    return chaveReal !== undefined ? (linha[chaveReal] ?? '') : '';
}

function escapar(texto) {
    return String(texto ?? '')
        .replace(/&/g, '&amp;')
        .replace(/</g, '&lt;')
        .replace(/>/g, '&gt;')
        .replace(/"/g, '&quot;')
        .replace(/'/g, '&#39;');
}

/* ---------------------------------------------------------
   3) Datas -> yyyy-mm-dd (formato do input date)
   --------------------------------------------------------- */
function serialParaISO(valor) {
    if (valor === '' || valor === null || valor === undefined) return '';

    if (valor instanceof Date && !isNaN(valor)) {
        const y = valor.getFullYear();
        const m = String(valor.getMonth() + 1).padStart(2, '0');
        const d = String(valor.getDate()).padStart(2, '0');
        return `${y}-${m}-${d}`;
    }

    if (typeof valor === 'number') {
        const d = XLSX.SSF.parse_date_code(valor);
        if (!d) return '';
        return `${d.y}-${String(d.m).padStart(2, '0')}-${String(d.d).padStart(2, '0')}`;
    }

    const texto = String(valor).trim();

    let m = texto.match(/^(\d{1,2})\/(\d{1,2})\/(\d{4})/);
    if (m) return `${m[3]}-${m[2].padStart(2, '0')}-${m[1].padStart(2, '0')}`;

    m = texto.match(/^(\d{4})-(\d{1,2})-(\d{1,2})/);
    if (m) return `${m[1]}-${m[2].padStart(2, '0')}-${m[3].padStart(2, '0')}`;

    return '';
}

function isoParaBR(iso) {
    if (!iso) return '';
    const [y, m, d] = iso.split('-');
    return `${d}/${m}/${y}`;
}

/* ---------------------------------------------------------
   4) Cabeçalho
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

/* ---------------------------------------------------------
   5) Células
   --------------------------------------------------------- */
function criarCelula(coluna, dados, indiceLinha) {
    const td = document.createElement('td');

    switch (coluna.tipo) {

        case 'input': {
            const input = document.createElement('input');
            input.type = 'text';
            input.className = 'campo-texto';
            input.dataset.campo = coluna.campo;
            input.dataset.linha = indiceLinha;
            input.placeholder = coluna.placeholder || '';
            input.value = pegarValor(dados, coluna.chave) ?? '';
            td.appendChild(input);
            break;
        }

        case 'data': {
            const input = document.createElement('input');
            input.type = 'date';
            input.className = 'campo-texto';
            input.dataset.campo = coluna.campo;
            input.dataset.linha = indiceLinha;
            input.value = serialParaISO(pegarValor(dados, coluna.chave));
            input.readOnly = true;
            input.tabIndex = -1;
            td.appendChild(input);
            break;
        }

        case 'checks': {
            const div = document.createElement('div');
            div.className = 'grupo-checkbox';

            const marcados = String(pegarValor(dados, coluna.chave) ?? '')
                .split(',')
                .map(s => normalizar(s));

            [
                ['falha_interna', 'Falha interna'],
                ['falha_externa', 'Falha externa'],
                ['outros',        'Outros']
            ].forEach(([valor, rotulo]) => {
                const label = document.createElement('label');

                const cb = document.createElement('input');
                cb.type = 'checkbox';
                cb.dataset.campo = coluna.campo;
                cb.dataset.linha = indiceLinha;
                cb.value = valor;
                cb.checked = marcados.some(m => m.includes(normalizar(valor).replace('_', '')));

                label.append(cb, document.createTextNode(' ' + rotulo));
                div.appendChild(label);
            });

            td.appendChild(div);
            break;
        }

        case 'modal': {
            const valor = String(pegarValor(dados, coluna.chave) ?? '').trim();

            if (!valor) {
                td.textContent = '-';
                td.classList.add('celula-vazia');
                break;
            }

            const botao = document.createElement('button');
            botao.type = 'button';
            botao.className = 'botao-modal';
            botao.textContent = valor.length > 22 ? valor.slice(0, 22) + '...' : valor;
            botao.title = 'Clique para ver o alarme completo';
            botao.addEventListener('click', () => abrirModalAlarme(valor, dados));
            td.appendChild(botao);
            break;
        }

        default: {
            const valor = pegarValor(dados, coluna.chave);
            td.textContent = (valor === 0 || valor) ? String(valor) : '';
            td.title = td.textContent;
        }
    }

    return td;
}

function criarLinha(dados, indiceLinha) {
    const tr = document.createElement('tr');
    tr.dataset.linha = String(indiceLinha);

    COLUNAS.forEach(coluna => tr.appendChild(criarCelula(coluna, dados, indiceLinha)));

    return tr;
}

function renderizarTabela(linhas) {
    corpoTabela.innerHTML = '';

    const fragmento = document.createDocumentFragment();
    linhas.forEach((linha, i) => fragmento.appendChild(criarLinha(linha, i)));
    corpoTabela.appendChild(fragmento);

    sincronizarLargura();
    botaoExportar.disabled = linhas.length === 0;
}

/* ---------------------------------------------------------
   6) Modal do alarme
   --------------------------------------------------------- */
function abrirModalAlarme(texto, dados) {
    const seq  = pegarValor(dados, 'Sequencia');
    const raiz = pegarValor(dados, 'RAIZ');
    const loc  = pegarValor(dados, 'Localidade');

    modalCorpo.innerHTML = `
        <dl class="modal-meta">
            <div><dt>Sequência</dt><dd>${escapar(seq)}</dd></div>
            <div><dt>RAIZ</dt><dd>${escapar(raiz)}</dd></div>
            <div><dt>Localidade</dt><dd>${escapar(loc)}</dd></div>
        </dl>
        <pre class="modal-texto">${escapar(texto)}</pre>
    `;

    modalAlarme.hidden = false;
    document.body.style.overflow = 'hidden';
    modalFechar.focus();
}

function fecharModalAlarme() {
    modalAlarme.hidden = true;
    document.body.style.overflow = '';
}

modalFechar.addEventListener('click', fecharModalAlarme);

modalAlarme.addEventListener('click', (evento) => {
    if (evento.target === modalAlarme) fecharModalAlarme();
});

document.addEventListener('keydown', (evento) => {
    if (evento.key === 'Escape' && !modalAlarme.hidden) fecharModalAlarme();
});

/* ---------------------------------------------------------
   7) Barra de rolagem fixa
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
   8) Leitura do arquivo
   --------------------------------------------------------- */
inputArquivo.addEventListener('change', async (evento) => {
    const arquivo = evento.target.files[0];
    if (!arquivo) return;

    const extensao = arquivo.name.split('.').pop().toLowerCase();

    if (extensao === 'pbix') {
        statusImportacao.textContent =
            'Arquivo .pbix não é uma planilha. Exporte os dados pelo Power BI Desktop (Exportar → Dados) como .xlsx ou .csv.';
        statusImportacao.className = 'status-importacao erro';
        inputArquivo.value = '';
        return;
    }

    try {
        statusImportacao.textContent = 'Lendo arquivo...';
        statusImportacao.className = 'status-importacao';

        const buffer = await arquivo.arrayBuffer();
        const workbook = XLSX.read(buffer, { type: 'array', cellDates: true });

        if (!workbook.SheetNames.length) throw new Error('O arquivo não tem nenhuma aba.');

        const nomeAba = workbook.SheetNames[0];
        const planilha = workbook.Sheets[nomeAba];

        /* header:1 devolve matriz; usamos para pegar os cabeçalhos reais */
        const matriz = XLSX.utils.sheet_to_json(planilha, { header: 1, defval: '', blankrows: false });
        if (matriz.length < 2) {
            statusImportacao.textContent = `A aba "${nomeAba}" está vazia.`;
            statusImportacao.className = 'status-importacao erro';
            return;
        }

        cabecalhosArquivo = matriz[0].map(h => String(h).trim()).filter(h => h !== '');

        dadosOriginais = XLSX.utils.sheet_to_json(planilha, { defval: '', blankrows: false });

        if (dadosOriginais.length === 0) {
            statusImportacao.textContent = `A aba "${nomeAba}" está vazia.`;
            statusImportacao.className = 'status-importacao erro';
            return;
        }

        renderizarTabela(dadosOriginais);

        const esperadas = COLUNAS.map(c => normalizar(c.chave));
        const naoMapeadas = cabecalhosArquivo.filter(k => !esperadas.includes(normalizar(k)));

        statusImportacao.textContent =
            `${dadosOriginais.length} linha(s) da aba "${nomeAba}" — ${cabecalhosArquivo.length} colunas no arquivo.`;
        statusImportacao.className = 'status-importacao ok';

        if (naoMapeadas.length > 0) {
            console.info('Colunas do arquivo não exibidas na tabela (serão exportadas mesmo assim):', naoMapeadas);
        }

    } catch (erro) {
        console.error('Falha ao ler o arquivo:', erro);
        statusImportacao.textContent = `Não foi possível ler o arquivo: ${erro.message}`;
        statusImportacao.className = 'status-importacao erro';
    } finally {
        inputArquivo.value = '';
    }
});

/* ---------------------------------------------------------
   9) Coleta do que o usuário editou
   Retorna um Map indiceLinha -> { status, previsao, area }
   --------------------------------------------------------- */
function coletarEdicoes() {
    const edicoes = new Map();

    corpoTabela.querySelectorAll('tr').forEach(tr => {
        const indice = Number(tr.dataset.linha);
        const registro = { status: '', previsao: '', area: [] };

        tr.querySelectorAll('input[data-campo="status"]').forEach(i => registro.status = i.value);
        tr.querySelectorAll('input[data-campo="previsao"]').forEach(i => registro.previsao = i.value);

        registro.area = [...tr.querySelectorAll('input[type="checkbox"]:checked')].map(cb => cb.value);

        edicoes.set(indice, registro);
    });

    return edicoes;
}

/* ---------------------------------------------------------
   10) Exportação: planilha idêntica à do site
   Parte das linhas originais e sobrescreve as 3 colunas.
   --------------------------------------------------------- */
function baixarTabela() {
    if (!dadosOriginais.length) {
        statusImportacao.textContent = 'Nada para exportar — carregue um arquivo primeiro.';
        statusImportacao.className = 'status-importacao erro';
        return;
    }

    const edicoes = coletarEdicoes();

    const rotulos = {
        falha_interna: 'Falha interna',
        falha_externa: 'Falha externa',
        outros: 'Outros'
    };

    const linhasExport = dadosOriginais.map((original, i) => {
        const edicao = edicoes.get(i) || { status: '', previsao: '', area: [] };

        const linha = {};
        cabecalhosArquivo.forEach(h => { linha[h] = original[h] ?? ''; });

        linha['STATUS'] = edicao.status;
        linha['Previsão'] = edicao.previsao;
        linha['Área_Responsável'] = edicao.area.map(v => rotulos[v] || v).join(', ');

        /* DataCriacao volta como data legível, não como serial */
        const chaveData = cabecalhosArquivo.find(h => normalizar(h) === normalizar('DataCriacao'));
        if (chaveData) {
            const iso = serialParaISO(original[chaveData]);
            linha[chaveData] = iso ? isoParaBR(iso) : (original[chaveData] ?? '');
        }

        return linha;
    });

        const planilha = XLSX.utils.json_to_sheet(linhasExport, {
        header: cabecalhosArquivo.concat(['STATUS', 'Previsão', 'Área_Responsável'])
            .filter((v, i, a) => a.indexOf(v) === i),
        skipHeader: false
    });

    /* Larguras de coluna para o arquivo não sair espremido */
    planilha['!cols'] = planilha['!ref']
        ? Array.from({ length: XLSX.utils.decode_range(planilha['!ref']).e.c + 1 }, () => ({ wch: 18 }))
        : [];

    const pasta = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(pasta, planilha, 'Acompanhamento');

    const agora = new Date();
    const carimbo =
        `${agora.getFullYear()}${String(agora.getMonth() + 1).padStart(2, '0')}${String(agora.getDate()).padStart(2, '0')}` +
        `_${String(agora.getHours()).padStart(2, '0')}${String(agora.getMinutes()).padStart(2, '0')}`;

    XLSX.writeFile(pasta, `acompanhamento_${carimbo}.xlsx`);
}

botaoExportar.addEventListener('click', baixarTabela);

/* ---------------------------------------------------------
   11) Inicialização
   --------------------------------------------------------- */
document.addEventListener('DOMContentLoaded', () => {
    montarCabecalho();
    sincronizarLargura();
});
