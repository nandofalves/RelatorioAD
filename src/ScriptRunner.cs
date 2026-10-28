using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Globalization;
using System.IO;
using System.Text;

namespace RelatorioAD
{
    class LinhaEventArgs : EventArgs
    {
        public string Texto { get; private set; }
        public bool Erro { get; private set; }

        public LinhaEventArgs(string texto, bool erro)
        {
            Texto = texto;
            Erro = erro;
        }
    }

    class FimEventArgs : EventArgs
    {
        public int ExitCode { get; private set; }
        public string CaminhoRelatorio { get; private set; }
        public string Resumo { get; private set; }
        public bool Cancelado { get; private set; }

        public FimEventArgs(int exitCode, string caminhoRelatorio, string resumo, bool cancelado)
        {
            ExitCode = exitCode;
            CaminhoRelatorio = caminhoRelatorio;
            Resumo = resumo;
            Cancelado = cancelado;
        }
    }

    // Executa o script com powershell.exe -File e repassa a saída linha a linha.
    // Os eventos chegam em outra thread: quem atualiza a UI usa BeginInvoke.
    class ScriptRunner
    {
        const string PrefixoRelatorio = "Relatório: ";
        const string PrefixoResumo = "Concluído.";

        public event EventHandler<LinhaEventArgs> Linha;
        public event EventHandler<FimEventArgs> Fim;

        Process processo;
        volatile string caminhoRelatorio;
        volatile string resumo;
        volatile bool cancelado;

        // Argumentos do powershell.exe da última execução (registrados no log)
        public string Argumentos { get; private set; }

        public void Iniciar(string scriptPath, IDictionary<string, string> parametros)
        {
            caminhoRelatorio = null;
            resumo = null;
            cancelado = false;

            StringBuilder args = new StringBuilder("-NoProfile -NonInteractive -ExecutionPolicy Bypass -File ");
            args.Append(Quote(scriptPath));
            foreach (KeyValuePair<string, string> p in parametros)
            {
                // valor null = switch
                args.Append(" -").Append(p.Key);
                if (p.Value != null) args.Append(' ').Append(Quote(p.Value));
            }
            Argumentos = args.ToString();

            Encoding oem = Encoding.GetEncoding(CultureInfo.CurrentCulture.TextInfo.OEMCodePage);
            string sys = Environment.GetEnvironmentVariable("SystemRoot") ?? @"C:\Windows";

            ProcessStartInfo psi = new ProcessStartInfo();
            psi.FileName = Path.Combine(sys, @"System32\WindowsPowerShell\v1.0\powershell.exe");
            psi.Arguments = Argumentos;
            psi.WorkingDirectory = Path.GetDirectoryName(scriptPath);
            psi.UseShellExecute = false;
            psi.RedirectStandardOutput = true;
            psi.RedirectStandardError = true;
            psi.CreateNoWindow = true;
            psi.StandardOutputEncoding = oem;
            psi.StandardErrorEncoding = oem;

            Process p2 = new Process();
            p2.StartInfo = psi;
            p2.EnableRaisingEvents = true;
            p2.OutputDataReceived += delegate(object s, DataReceivedEventArgs e) { Receber(e.Data, false); };
            p2.ErrorDataReceived += delegate(object s, DataReceivedEventArgs e) { Receber(e.Data, true); };
            p2.Exited += Terminou;

            processo = p2;
            p2.Start();
            p2.BeginOutputReadLine();
            p2.BeginErrorReadLine();
        }

        public void Cancelar()
        {
            cancelado = true;
            try
            {
                if (processo != null && !processo.HasExited) processo.Kill();
            }
            catch (Exception) { }
        }

        void Receber(string texto, bool erro)
        {
            if (texto == null) return;
            if (!erro)
            {
                if (texto.StartsWith(PrefixoRelatorio, StringComparison.Ordinal))
                    caminhoRelatorio = texto.Substring(PrefixoRelatorio.Length).Trim();
                else if (texto.StartsWith(PrefixoResumo, StringComparison.Ordinal))
                    resumo = texto;
            }
            EventHandler<LinhaEventArgs> h = Linha;
            if (h != null) h(this, new LinhaEventArgs(texto, erro));
        }

        void Terminou(object sender, EventArgs e)
        {
            Process p = (Process)sender;
            p.WaitForExit(); // esvazia os buffers de saída antes do Fim
            int codigo = p.ExitCode;
            p.Dispose();

            EventHandler<FimEventArgs> h = Fim;
            if (h != null) h(this, new FimEventArgs(codigo, caminhoRelatorio, resumo, cancelado));
        }

        // Aspas conforme as regras do CommandLineToArgvW
        public static string Quote(string s)
        {
            StringBuilder sb = new StringBuilder("\"");
            int barras = 0;
            foreach (char c in s)
            {
                if (c == '\\')
                {
                    barras++;
                    continue;
                }
                if (c == '"')
                    sb.Append('\\', barras * 2 + 1);
                else
                    sb.Append('\\', barras);
                sb.Append(c);
                barras = 0;
            }
            sb.Append('\\', barras * 2);
            sb.Append('"');
            return sb.ToString();
        }
    }
}
