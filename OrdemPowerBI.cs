using System;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
using System.Linq;
using CsvHelper;
using CsvHelper.Configuration;

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

}

// Mapeia o cabeçalho da planilha para as propriedades
public sealed class OrdemMap : ClassMap<OrdemPowerBI>
{
    public OrdemMap()
    {
        Map(mapa => mapa.Sequencia).Name("Sequência");
        Map(mapa => mapa.SequenciaRaiz).Name("SequenciaRaiz");
        Map(mapa => mapa.Raiz).Name("RAIZ");
        Map(mapa => mapa.Prioridade).Name("Prioridade");
        Map(mapa => mapa.Aging).Name("Aging");
        Map(mapa => mapa.Categoria).Name("Categoria");
        Map(mapa => mapa.Classificacao).Name("Classificação");
        Map(mapa => mapa.GrupoResponsavel).Name("Grupo responsável");
        Map(mapa => mapa.GrupoPlantaInterna).Name("GRUPO_PLANTA_INTERNA");
        Map(mapa => mapa.Localidade).Name("Localidade");
        Map(mapa => mapa.AfetFTTX).Name("AfetFTTX");
        Map(mapa => mapa.GestorAtual).Name("Gestor Atual");
        Map(mapa => mapa.UF).Name("UF");
        Map(mapa => mapa.SiglaLocalidade).Name("SiglaLocalidade");
        Map(mapa => mapa.SiglaSite).Name("Sigla Site");
        Map(mapa => mapa.DataCriacao).Name("Data de criação");
    }
}

public static class LeitorPowerBI
{
    public static List<OrdemPowerBI> LerCsv(string caminho)
    {
        var configuracao = new CsvConfiguration(CultureInfo.GetCultureInfo("pt-BR"))
        {
            Delimiter = ";",
            PrepareHeaderForMatch = args => args.Header.Trim(),
            MissingFieldFound = null
        };

        using var reader = new StreamReader(caminho, System.Text.Encoding.UTF8);
        using var csv = new CsvReader(reader, configuracao);
        csv.Context.RegisterClassMap<OrdemMap>();
        return csv.GetRecords<OrdemPowerBI>().ToList();
    }
}
