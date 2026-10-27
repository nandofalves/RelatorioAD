using System;
using System.IO;
using System.Reflection;
using System.Text;

namespace RelatorioAD
{
    // Uma versão do Windows Server e o script embutido que ela usa
    class VersaoSO
    {
        public string Id;
        public string Rotulo;
        public int Build;
        public Version PsMinima;
        public string Recurso;

        public VersaoSO(string id, string rotulo, int build, string psMinima)
        {
            Id = id;
            Rotulo = rotulo;
            Build = build;
            PsMinima = new Version(psMinima);
            Recurso = "RelatorioAD.Scripts." + id + ".ps1";
        }

        public override string ToString()
        {
            return "Windows Server " + Rotulo;
        }
    }

    static class ScriptCatalog
    {
        // Em ordem crescente de build (OsDetector.Detectar depende disso)
        public static readonly VersaoSO[] Todas = new VersaoSO[]
        {
            new VersaoSO("2008R2", "2008 R2", 7600, "2.0"),
            new VersaoSO("2012", "2012", 9200, "3.0"),
            new VersaoSO("2012R2", "2012 R2", 9600, "4.0"),
            new VersaoSO("2016", "2016", 14393, "5.1"),
            new VersaoSO("2019", "2019", 17763, "5.1"),
            new VersaoSO("2022", "2022", 20348, "5.1"),
            new VersaoSO("2025", "2025", 26100, "5.1"),
        };

        public static VersaoSO PorId(string id)
        {
            foreach (VersaoSO v in Todas)
                if (v.Id == id) return v;
            return null;
        }

        public static string LerScript(VersaoSO v)
        {
            Stream s = Assembly.GetExecutingAssembly().GetManifestResourceStream(v.Recurso);
            if (s == null)
                throw new InvalidOperationException("O script da versão " + v + " (" + v.Recurso + ") não está embutido neste executável. Gere o .exe de novo com o build.ps1.");
            using (StreamReader r = new StreamReader(s, Encoding.UTF8, true))
                return r.ReadToEnd();
        }

        // Grava o script em UTF-8 com BOM (o PowerShell 2.0–5.1 lê sem BOM como ANSI)
        public static void Salvar(VersaoSO v, string caminho)
        {
            File.WriteAllText(caminho, LerScript(v), new UTF8Encoding(true));
        }

        public static string Extrair(VersaoSO v, string pastaSaida)
        {
            string pasta = Path.Combine(pastaSaida, "scripts");
            Directory.CreateDirectory(pasta);
            string caminho = Path.Combine(pasta, "Relatorio-" + v.Id + ".ps1");
            Salvar(v, caminho);
            return caminho;
        }
    }
}
