using System;
using System.Threading;
using System.Windows.Forms;

namespace RelatorioAD
{
    static class Program
    {
        [STAThread]
        static void Main()
        {
            Application.EnableVisualStyles();
            Application.SetCompatibleTextRenderingDefault(false);
            Application.ThreadException += delegate(object s, ThreadExceptionEventArgs e) { Mostrar(e.Exception); };
            AppDomain.CurrentDomain.UnhandledException += delegate(object s, UnhandledExceptionEventArgs e)
            {
                Mostrar(e.ExceptionObject as Exception);
            };
            Application.Run(new MainForm());
        }

        static void Mostrar(Exception ex)
        {
            string msg = ex == null ? "Erro desconhecido." : ex.Message;
            MessageBox.Show(msg, "RelatorioAD — erro inesperado", MessageBoxButtons.OK, MessageBoxIcon.Error);
        }
    }
}
