using System;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;
using ClosedXML.Excel;

public static class ExportadorExcel
{
    private static readonly string[] ColunasFinais =
    {
        "Sequencia", "SequenciaRaiz", "RAIZ", "Prioridade", "Aging", "Categoria",
        "Classificação", "Grupo Responsável", "GRUPO_PLANTA_INTERNA", "Localidade",
        "AfetFTTX", "Gestor Atual", "UF", "SiglaLocalidade", "Sigla Site",
        "AfetIPTV", "AfetVoz",
        "STATUS", "Previsão", "Área_Responsável",
        "DataCriacao", "DataApresentacao", "TipoRede", "TpRedeRAIZ", "TipoPlanta",
        "TipoFalha", "TipoAlarme", "TipodeBilhete", "TipoAfetacao", "LocalidadeRAIZ",
        "CodLocalidadeRAIZ", "SEGS_DURACAO", "QTD_DIAS", "REGIONAL", "Faixa_dias",
        "Alarme", "B2B"
    };

    public static void Salvar(IEnumerable<OrdemPowerBI> ordens, string caminhoSaida)
    {
        using var wb = new XLWorkbook();
        var ws = wb.Worksheets.Add("Acompanhamento");

        for (int c = 0; c < ColunasFinais.Length; c++)
        {
            var celula = ws.Cell(1, c + 1);
            celula.Value = ColunasFinais[c];
            celula.Style.Font.Bold = true;
            celula.Style.Fill.BackgroundColor = XLColor.FromHtml("#660099");
            celula.Style.Font.FontColor = XLColor.White;
        }

        int linhaAtual = 2;
        foreach (var o in ordens)
        {
            ws.Cell(linhaAtual, 1).Value = o.Sequencia;
            ws.Cell(linhaAtual, 2).Value = o.SequenciaRaiz;
            ws.Cell(linhaAtual, 3).Value = o.Raiz;
            ws.Cell(linhaAtual, 4).Value = o.Prioridade;
            ws.Cell(linhaAtual, 5).Value = o.Aging;
            ws.Cell(linhaAtual, 6).Value = o.Categoria;
            ws.Cell(linhaAtual, 7).Value = o.Classificacao;
            ws.Cell(linhaAtual, 8).Value = o.GrupoResponsavel;
            ws.Cell(linhaAtual, 9).Value = o.GrupoPlantaInterna;
            ws.Cell(linhaAtual, 10).Value = o.Localidade;
            ws.Cell(linhaAtual, 11).Value = o.AfetFTTX;
            ws.Cell(linhaAtual, 12).Value = o.GestorAtual;
            ws.Cell(linhaAtual, 13).Value = o.UF;
            ws.Cell(linhaAtual, 14).Value = o.SiglaLocalidade;
            ws.Cell(linhaAtual, 15).Value = o.SiglaSite;
            ws.Cell(linhaAtual, 18).Value = o.Status;
            ws.Cell(linhaAtual, 19).Value = o.Previsao;
            ws.Cell(linhaAtual, 20).Value = o.AreaResponsavel;

            if (o.DataCriacao.HasValue)
            {
                ws.Cell(linhaAtual, 21).Value = o.DataCriacao.Value;
                ws.Cell(linhaAtual, 21).Style.DateFormat.Format = "dd/mm/yyyy";
            }

            linhaAtual++;
        }

        ws.Columns().AdjustToContents();
        ws.SheetView.FreezeRows(1);
        wb.SaveAs(caminhoSaida);
    }
}