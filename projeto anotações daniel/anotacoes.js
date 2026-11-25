const corpoTabela = document.getElementById('corpoTabela');
const inputArquivo = document.getElementById('botaoPowerBI');   // ← era 'arquivoPowerBI'
const statusImportacao = document.getElementById('statusImportacao');

const Colunas_PowerBI = [
    'Sequência', 'SequenciaRaiz', 'RAIZ', 'Prioridade', 'Aging', 'Categoria',
    'Classificação', 'Grupo responsável', 'GRUPO_PLANTA_INTERNA', 'Localidade',
    'AfetFTTX', 'Gestor Atual', 'UF', 'SiglaLocalidade', 'Sigla Site', 'Data de criação'
];

// ← era getElementById('div'), que não existe. Agora cria um elemento solto.
function escapar(valor) {
    const div = document.createElement('div');
    div.textContent = valor ?? '';
    return div.innerHTML;
}

function criarLinha(dados) {                       // ← era criarLinhas (conflito de nome)
    const tr = document.createElement('tr');
    const id = dados['Sequência'] ?? crypto.randomUUID();

    Colunas_PowerBI.forEach(coluna => {
        const td = document.createElement('td');
        td.textContent = dados[coluna] ?? '';
        tr.appendChild(td);
    });

    const tdStatus = document.createElement('td');
    tdStatus.innerHTML = `<input type="text" class="campo-texto" data-campo="status" placeholder="Digite o status">`;
    tr.appendChild(tdStatus);

    const tdPrevisao = document.createElement('td');
    tdPrevisao.innerHTML = `<input type="date" class="campo-texto" data-campo="previsao">`;  // ← era type="data"
    tr.appendChild(tdPrevisao);

    const tdAreaResponsavel = document.createElement('td');
    tdAreaResponsavel.innerHTML = `
        <div class="grupo-checkbox">
            <label><input type="checkbox" data-campo="area" value="falha_interna"> Falha interna</label>
            <label><input type="checkbox" data-campo="area" value="falha_externa"> Falha externa</label>
            <label><input type="checkbox" data-campo="area" value="outros"> Outros</label>
        </div>`;
    tr.appendChild(tdAreaResponsavel);

    tr.dataset.id = id;
    return tr;
}

function renderizarTabela(linhas) {
    corpoTabela.innerHTML = '';
    linhas.forEach(linha => corpoTabela.appendChild(criarLinha(linha)));  // ← era criarLinhas
}

inputArquivo.addEventListener('change', async (evento) => {
    const arquivo = evento.target.files[0];
    if (!arquivo) return;

    try {
        statusImportacao.textContent = 'Lendo arquivo...';
        statusImportacao.className = 'status-importacao';

        const buffer = await arquivo.arrayBuffer();
        const workbook = XLSX.read(buffer, { type: 'array' });
        const primeiraAba = workbook.Sheets[workbook.SheetNames[0]];
        const linhas = XLSX.utils.sheet_to_json(primeiraAba, { defval: '' });

        renderizarTabela(linhas);
        statusImportacao.textContent = `${linhas.length} linha(s) carregada(s).`;
        statusImportacao.className = 'status-importacao ok';
    } catch (erro) {
        console.error(erro);
        statusImportacao.textContent = 'Não foi possível ler o arquivo.';
        statusImportacao.className = 'status-importacao erro';
    }
});