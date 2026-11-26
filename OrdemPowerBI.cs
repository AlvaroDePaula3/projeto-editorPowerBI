using System;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
using System.Linq;
using CsvHelper;
using CsvHelper.Configuration;
using ClosedXML.Excel;

public class OrdemPowerBI
{
    public string Sequencia { get; set; }
    public string SequenciaRaiz { get; set; }
    public string Raiz { get; set; }
    public string Prioridade { get; set; }
    public string Aging { get; set; }
    public string Categoria { get; set; }
    public string Classificacao { get; set; }
    public string GrupoResponsavel { get; set; }
    public string GrupoPlantaInterna { get; set; }
    public string Localidade { get; set; }
    public string AfetFTTX { get; set; }
    public string GestorAtual { get; set; }
    public string UF { get; set; }
    public string SiglaLocalidade { get; set; }
    public string SiglaSite { get; set; }
    public DateTime? DataCriacao { get; set; }

    // Campos editáveis no site
    public string Status { get; set; }
    public string Previsao { get; set; }
    public string AreaResponsavel { get; set; }
}

public sealed class OrdemMap : ClassMap<OrdemPowerBI>
{
    public OrdemMap()
    {
        Map(m => m.Sequencia).Name("Sequencia");
        Map(m => m.SequenciaRaiz).Name("SequenciaRaiz");
        Map(m => m.Raiz).Name("RAIZ");
        Map(m => m.Prioridade).Name("Prioridade");
        Map(m => m.Aging).Name("Aging");
        Map(m => m.Categoria).Name("Categoria");
        Map(m => m.Classificacao).Name("Classificação");
        Map(m => m.GrupoResponsavel).Name("Grupo Responsável");
        Map(m => m.GrupoPlantaInterna).Name("GRUPO_PLANTA_INTERNA");
        Map(m => m.Localidade).Name("Localidade");
        Map(m => m.AfetFTTX).Name("AfetFTTX");
        Map(m => m.GestorAtual).Name("Gestor Atual");
        Map(m => m.UF).Name("UF");
        Map(m => m.SiglaLocalidade).Name("SiglaLocalidade");
        Map(m => m.SiglaSite).Name("Sigla Site");
        Map(m => m.DataCriacao).Name("DataCriacao", "DataCriação");
        Map(m => m.Status).Name("STATUS").Optional();
        Map(m => m.Previsao).Name("Previsão").Optional();
        Map(m => m.AreaResponsavel).Name("Área_Responsável").Optional();
    }
}

public static class LeitorPowerBI
{
    public static List<OrdemPowerBI> Ler(string caminho)
    {
        var ext = Path.GetExtension(caminho).ToLowerInvariant();
        return ext switch
        {
            ".csv" => LerCsv(caminho),
            ".xlsx" or ".xls" => LerExcel(caminho),
            ".pbix" => throw new NotSupportedException(
                ".pbix é um pacote binário do Power BI e não pode ser lido como planilha. " +
                "Exporte os dados pelo Power BI Desktop (Exportar → Dados) para .xlsx ou .csv."),
            _ => throw new NotSupportedException($"Extensão não suportada: {ext}")
        };
    }

    private static List<OrdemPowerBI> LerCsv(string caminho)
    {
        var config = new CsvConfiguration(CultureInfo.GetCultureInfo("pt-BR"))
        {
            Delimiter = ";",
            PrepareHeaderForMatch = args => args.Header.Trim(),
            MissingFieldFound = null,
            BadDataFound = null
        };

        using var reader = new StreamReader(caminho, System.Text.Encoding.UTF8);
        using var csv = new CsvReader(reader, config);
        csv.Context.RegisterClassMap<OrdemMap>();
        return csv.GetRecords<OrdemPowerBI>().ToList();
    }

    private static List<OrdemPowerBI> LerExcel(string caminho)
    {
        var lista = new List<OrdemPowerBI>();
        using var wb = new XLWorkbook(caminho);
        var ws = wb.Worksheet(1);

        var cabecalho = ws.Row(1).CellsUsed()
            .Select(c => c.GetString().Trim())
            .ToList();

        string Valor(IXLRow linha, params string[] nomes)
        {
            foreach (var nome in nomes)
            {
                var indice = cabecalho.FindIndex(h =>
                    string.Equals(h, nome, StringComparison.OrdinalIgnoreCase));
                if (indice >= 0) return linha.Cell(indice + 1).GetString().Trim();
            }
            return string.Empty;
        }

        foreach (var linha in ws.RowsUsed().Skip(1))
        {
            lista.Add(new OrdemPowerBI
            {
                Sequencia = Valor(linha, "Sequencia"),
                SequenciaRaiz = Valor(linha, "SequenciaRaiz"),
                Raiz = Valor(linha, "RAIZ"),
                Prioridade = Valor(linha, "Prioridade"),
                Aging = Valor(linha, "Aging"),
                Categoria = Valor(linha, "Categoria"),
                Classificacao = Valor(linha, "Classificação"),
                GrupoResponsavel = Valor(linha, "Grupo Responsável"),
                GrupoPlantaInterna = Valor(linha, "GRUPO_PLANTA_INTERNA"),
                Localidade = Valor(linha, "Localidade"),
                AfetFTTX = Valor(linha, "AfetFTTX"),
                GestorAtual = Valor(linha, "Gestor Atual"),
                UF = Valor(linha, "UF"),
                SiglaLocalidade = Valor(linha, "SiglaLocalidade"),
                SiglaSite = Valor(linha, "Sigla Site"),
                DataCriacao = DateTime.TryParse(
                    Valor(linha, "DataCriacao", "DataCriação"),
                    CultureInfo.GetCultureInfo("pt-BR"),
                    DateTimeStyles.None,
                    out var data) ? data : null,
                Status = Valor(linha, "STATUS"),
                Previsao = Valor(linha, "Previsão"),
                AreaResponsavel = Valor(linha, "Área_Responsável")
            });
        }

        return lista;
    }
}
