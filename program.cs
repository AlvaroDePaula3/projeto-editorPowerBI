using System;
using System.IO;
using System.Linq;

class Program
{
    static int Main(string[] args)
    {
        if (args.Length < 2)
        {
            Console.WriteLine("Uso: OrdemPowerBI <entrada.xlsx|csv> <saida.xlsx>");
            return 1;
        }

        try
        {
            var ordens = LeitorPowerBI.Ler(args[0]);
            ExportadorExcel.Salvar(ordens, args[1]);
            Console.WriteLine($"{ordens.Count} linha(s) exportada(s) para {args[1]}");
            return 0;
        }
        catch (Exception ex)
        {
            Console.Error.WriteLine($"Erro: {ex.Message}");
            return 2;
        }
    }
}