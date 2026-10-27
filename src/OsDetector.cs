using System;
using System.IO;
using System.Security.Principal;
using Microsoft.Win32;

namespace RelatorioAD
{
    static class OsDetector
    {
        const string ChaveVersao = @"SOFTWARE\Microsoft\Windows NT\CurrentVersion";

        static string LerHklm(string chave, string valor)
        {
            try
            {
                using (RegistryKey k = Registry.LocalMachine.OpenSubKey(chave))
                {
                    if (k == null) return null;
                    object o = k.GetValue(valor);
                    return o == null ? null : o.ToString();
                }
            }
            catch (Exception)
            {
                return null;
            }
        }

        public static int BuildAtual()
        {
            int build;
            foreach (string nome in new string[] { "CurrentBuild", "CurrentBuildNumber" })
            {
                string s = LerHklm(ChaveVersao, nome);
                if (s != null && int.TryParse(s, out build)) return build;
            }
            return Environment.OSVersion.Version.Build;
        }

        public static bool EhServidor()
        {
            string tipo = LerHklm(ChaveVersao, "InstallationType");
            if (tipo == "Server" || tipo == "Server Core") return true;
            string produto = NomeProduto();
            return produto != null && produto.IndexOf("Server", StringComparison.OrdinalIgnoreCase) >= 0;
        }

        public static string NomeProduto()
        {
            return LerHklm(ChaveVersao, "ProductName");
        }

        public static VersaoSO Detectar()
        {
            if (!EhServidor()) return null;
            int build = BuildAtual();
            VersaoSO achada = null;
            foreach (VersaoSO v in ScriptCatalog.Todas)
                if (v.Build <= build) achada = v;
            return achada;
        }

        public static Version VersaoPowerShell()
        {
            foreach (string chave in new string[] {
                @"SOFTWARE\Microsoft\PowerShell\3\PowerShellEngine",
                @"SOFTWARE\Microsoft\PowerShell\1\PowerShellEngine" })
            {
                string s = LerHklm(chave, "PowerShellVersion");
                if (string.IsNullOrEmpty(s)) continue;
                try { return new Version(s); }
                catch (Exception) { }
            }
            return null;
        }

        public static bool ModuloADInstalado()
        {
            string sys = Environment.GetEnvironmentVariable("SystemRoot") ?? @"C:\Windows";
            return Directory.Exists(Path.Combine(sys, @"System32\WindowsPowerShell\v1.0\Modules\ActiveDirectory"));
        }

        public static bool EhAdministrador()
        {
            try
            {
                return new WindowsPrincipal(WindowsIdentity.GetCurrent()).IsInRole(WindowsBuiltInRole.Administrator);
            }
            catch (Exception)
            {
                return false;
            }
        }
    }
}
