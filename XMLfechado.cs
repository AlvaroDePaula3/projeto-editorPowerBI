<!DOCTYPE html>
<html lang="pt-BR">
<head>
    <meta charset="UTF-8">
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    <title>Anotações Power BI</title>
    <link rel="stylesheet" href="anotacoes.css">
    <script src="https://cdn.sheetjs.com/xlsx-0.20.3/package/dist/xlsx.full.min.js"></script>
</head>
<body>

    <header class="topo">
        <div class="divisoriaBotaoPowerBI">
            <label class="botaoEnviarArquivo" for="botaoPowerBI">
                <span>Carregar arquivo</span>
            </label>
            <input type="file" id="botaoPowerBI" accept=".xlsx,.xls,.csv" hidden>

            <button type="button" class="botao-exportar" id="botaoExportar" disabled>
                Baixar Excel
            </button>

            <span id="statusImportacao" class="status-importacao"></span>
        </div>
    </header>

    <div class="divisorioTabela" id="divisorioTabela">
        <table class="tabelaExportada">
            <thead id="cabecalhoTabela"></thead>
            <tbody id="corpoTabela"></tbody>
        </table>
    </div>

    <div class="barraRolagemFixa" id="barraRolagemFixa">
        <div class="barraRolagemFixa-conteudo" id="barraRolagemFixaConteudo"></div>
    </div>

    <div class="modal-fundo" id="modalAlarme" hidden>
        <div class="modal-caixa" role="dialog" aria-modal="true" aria-labelledby="modalTitulo">
            <header class="modal-topo">
                <h2 id="modalTitulo">Alarme</h2>
                <button type="button" class="modal-fechar" id="modalFechar" aria-label="Fechar">×</button>
            </header>
            <div class="modal-corpo" id="modalCorpo"></div>
        </div>
    </div>

    <script src="anotacoes.js"></script>
</body>
</html>