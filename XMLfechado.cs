using System;
using System.Collections.Generic;
using System.Linq;
using ClosedXML.Excel;

public static class LeitorExcel
{
    public static List<OrdemPowerBI> LerExcel(string caminho)
    {
        var lista = new List<OrdemPowerBI>();
        using var wb = new XLWorkbook(caminho);
        var ws = wb.Worksheet(1);
        var cabecalho = ws.Row(1).CellsUsed()
                          .Select(c => c.GetString().Trim())
                          .ToList();

        foreach (var linha in ws.RowsUsed().Skip(1))
        {
            string Valor(string nomeColuna)
            {
                var indice = cabecalho.IndexOf(nomeColuna) + 1;
                return indice > 0 ? linha.Cell(indice).GetString().Trim() : string.Empty;
            }

            lista.Add(new OrdemPowerBI
            {
                Sequencia = Valor("Sequência"),
                Raiz = Valor("RAIZ"),
                Prioridade = Valor("Prioridade"),
                Aging = Valor("Aging"),
                Localidade = Valor("Localidade"),
                UF = Valor("UF"),
                GestorAtual = Valor("Gestor Atual"),
                DataCriacao = DateTime.TryParse(Valor("Data de criação"),
                                CultureInfo.GetCultureInfo("pt-BR"),
                                DateTimeStyles.None,
                                out var dataCriada) ? dataCriada : null
            });
        }
        return lista;
    }
}