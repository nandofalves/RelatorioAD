using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Drawing;
using System.IO;
using System.Text.RegularExpressions;
using System.Windows.Forms;

namespace RelatorioAD
{
    class MainForm : Form
    {
        // Mesmos padrões do script
        const string PadraoResponsavel = "Compumaq · Werner Informática";
        const string PadraoSaida = @"C:\Relatorio_AD";
        const string PadraoTerceiros = "Terceir|Tercer|Fornecedor|Externo";
        const string PadraoSensiveis = "^(rh|pessoal|dp|financeiro|contabil|contabilidade|diretoria|juridico|folha)";

        // Ambiente detectado na abertura
        readonly VersaoSO detectada = OsDetector.Detectar();
        readonly Version psInstalado = OsDetector.VersaoPowerShell();
        readonly bool moduloAD = OsDetector.ModuloADInstalado();
        readonly bool admin = OsDetector.EhAdministrador();

        Label lblDetectado, lblAviso;
        ComboBox cboVersao;
        TextBox txtCliente, txtResponsavel, txtSaida, txtCaminhos, txtTerceiros, txtSensiveis, txtLog;
        NumericUpDown numProfundidade, numDias;
        CheckBox chkAbrir;
        Button btnSaida, btnGerar, btnCancelar, btnAbrirRel, btnAbrirPasta, btnExportar;
        ToolStripStatusLabel lblStatus;
        Timer timer;
        TableLayoutPanel tblParametros;

        ScriptRunner runner;
        bool executando;
        DateTime inicio;
        string caminhoRelatorio;
        string erroAmigavel;

        public MainForm()
        {
            Text = "RelatorioAD 1.0 — Relatório de acessos do AD";
            Font = new Font("Segoe UI", 9F);
            AutoScaleDimensions = new SizeF(96F, 96F);
            AutoScaleMode = AutoScaleMode.Dpi;
            ClientSize = new Size(900, 680);
            MinimumSize = new Size(700, 560);
            StartPosition = FormStartPosition.CenterScreen;

            MontarTela();

            cboVersao.Items.AddRange(ScriptCatalog.Todas);
            cboVersao.SelectedItem = detectada ?? ScriptCatalog.PorId("2022");
            cboVersao.SelectedIndexChanged += delegate { AtualizarAviso(); };

            lblDetectado.Text = TextoDetectado();
            AtualizarAviso();
            Status("Pronto");
        }

        // ------------------------------------------------------------------
        // Montagem da interface
        // ------------------------------------------------------------------
        void MontarTela()
        {
            TableLayoutPanel principal = new TableLayoutPanel();
            principal.Dock = DockStyle.Fill;
            principal.Padding = new Padding(10);
            principal.ColumnCount = 1;
            principal.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 100F));
            principal.RowCount = 6;
            for (int i = 0; i < 5; i++) principal.RowStyles.Add(new RowStyle(SizeType.AutoSize));
            principal.RowStyles.Add(new RowStyle(SizeType.Percent, 100F));

            // Topo: ambiente detectado + versão
            lblDetectado = new Label();
            lblDetectado.AutoSize = true;
            lblDetectado.Margin = new Padding(3, 0, 3, 6);
            principal.Controls.Add(lblDetectado, 0, 0);

            FlowLayoutPanel linhaVersao = NovoFluxo();
            Label lblVersao = new Label();
            lblVersao.Text = "Versão do Windows Server:";
            lblVersao.AutoSize = true;
            lblVersao.Margin = new Padding(3, 7, 3, 3);
            cboVersao = new ComboBox();
            cboVersao.DropDownStyle = ComboBoxStyle.DropDownList;
            cboVersao.Width = 220;
            linhaVersao.Controls.Add(lblVersao);
            linhaVersao.Controls.Add(cboVersao);
            principal.Controls.Add(linhaVersao, 0, 1);

            lblAviso = new Label();
            lblAviso.AutoSize = true;
            lblAviso.BackColor = Color.FromArgb(255, 244, 179);
            lblAviso.ForeColor = Color.FromArgb(92, 64, 0);
            lblAviso.Padding = new Padding(6, 4, 6, 4);
            lblAviso.Margin = new Padding(3, 4, 3, 6);
            principal.Controls.Add(lblAviso, 0, 2);

            // Parâmetros
            GroupBox grp = new GroupBox();
            grp.Text = "Parâmetros";
            grp.Dock = DockStyle.Fill;
            grp.AutoSize = true;
            grp.Padding = new Padding(8);
            tblParametros = new TableLayoutPanel();
            tblParametros.Dock = DockStyle.Fill;
            tblParametros.AutoSize = true;
            // 4 colunas: rótulo | campo | rótulo | campo (campos longos ocupam as 3 últimas)
            tblParametros.ColumnCount = 4;
            tblParametros.ColumnStyles.Add(new ColumnStyle(SizeType.AutoSize));
            tblParametros.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 50F));
            tblParametros.ColumnStyles.Add(new ColumnStyle(SizeType.AutoSize));
            tblParametros.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 50F));

            txtCliente = NovoTexto("");
            txtResponsavel = NovoTexto(PadraoResponsavel);
            AddPar("Cliente:", txtCliente, "Responsável:", txtResponsavel);
            AddDica("Cliente vazio = nome NetBIOS do domínio");

            txtSaida = NovoTexto(PadraoSaida);
            btnSaida = new Button();
            btnSaida.Text = "...";
            btnSaida.Width = 32;
            btnSaida.Anchor = AnchorStyles.Right | AnchorStyles.Top;
            btnSaida.Margin = new Padding(3, 2, 3, 2);
            btnSaida.Click += delegate { EscolherPasta(); };
            TableLayoutPanel linhaSaida = new TableLayoutPanel();
            linhaSaida.Dock = DockStyle.Fill;
            linhaSaida.AutoSize = true;
            linhaSaida.Margin = new Padding(0);
            linhaSaida.ColumnCount = 2;
            linhaSaida.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 100F));
            linhaSaida.ColumnStyles.Add(new ColumnStyle(SizeType.AutoSize));
            linhaSaida.Controls.Add(txtSaida, 0, 0);
            linhaSaida.Controls.Add(btnSaida, 1, 0);
            AddLinha("Pasta de saída:", linhaSaida);

            txtCaminhos = NovoTexto("");
            txtCaminhos.Multiline = true;
            txtCaminhos.ScrollBars = ScrollBars.Vertical;
            txtCaminhos.AcceptsReturn = true;
            txtCaminhos.Height = 44;
            AddLinha("Caminhos:", txtCaminhos);
            AddDica("Opcional, um por linha. Vazio = todos os compartilhamentos do servidor");

            numProfundidade = NovoNumero(0, 10, 2);
            numDias = NovoNumero(1, 3650, 90);
            AddPar("Profundidade:", numProfundidade, "Dias inativo:", numDias);
            txtTerceiros = NovoTexto(PadraoTerceiros);
            AddLinha("Padrão de terceiros:", txtTerceiros);
            txtSensiveis = NovoTexto(PadraoSensiveis);
            AddLinha("Pastas sensíveis:", txtSensiveis);

            chkAbrir = new CheckBox();
            chkAbrir.Text = "Abrir relatório ao terminar";
            chkAbrir.Checked = true;
            chkAbrir.AutoSize = true;
            AddLinha("", chkAbrir);

            grp.Controls.Add(tblParametros);
            principal.Controls.Add(grp, 0, 3);

            // Botões
            FlowLayoutPanel botoes = NovoFluxo();
            botoes.Margin = new Padding(0, 6, 0, 6);
            btnGerar = NovoBotao("Gerar relatório", botoes);
            btnGerar.Font = new Font(Font, FontStyle.Bold);
            btnGerar.Click += delegate { Gerar(); };
            btnCancelar = NovoBotao("Cancelar", botoes);
            btnCancelar.Enabled = false;
            btnCancelar.Click += delegate { CancelarExecucao(); };
            btnAbrirRel = NovoBotao("Abrir relatório", botoes);
            btnAbrirRel.Enabled = false;
            btnAbrirRel.Click += delegate { AbrirRelatorio(); };
            btnAbrirPasta = NovoBotao("Abrir pasta", botoes);
            btnAbrirPasta.Enabled = false;
            btnAbrirPasta.Click += delegate { AbrirPasta(); };
            btnExportar = NovoBotao("Exportar script...", botoes);
            btnExportar.Click += delegate { ExportarScript(); };
            AcceptButton = btnGerar;
            principal.Controls.Add(botoes, 0, 4);

            // Log
            txtLog = new TextBox();
            txtLog.Multiline = true;
            txtLog.ReadOnly = true;
            txtLog.ScrollBars = ScrollBars.Both;
            txtLog.WordWrap = false;
            txtLog.MaxLength = 0;
            txtLog.BackColor = SystemColors.Window;
            txtLog.Font = new Font("Consolas", 9F);
            txtLog.Dock = DockStyle.Fill;
            principal.Controls.Add(txtLog, 0, 5);

            // Barra de status
            StatusStrip status = new StatusStrip();
            lblStatus = new ToolStripStatusLabel();
            lblStatus.Spring = true;
            lblStatus.TextAlign = ContentAlignment.MiddleLeft;
            status.Items.Add(lblStatus);

            timer = new Timer();
            timer.Interval = 1000;
            timer.Tick += delegate { Status("Executando... (" + Decorrido() + ")"); };

            Controls.Add(principal);
            Controls.Add(status);
        }

        static FlowLayoutPanel NovoFluxo()
        {
            FlowLayoutPanel f = new FlowLayoutPanel();
            f.AutoSize = true;
            f.Dock = DockStyle.Fill;
            f.WrapContents = true;
            f.Margin = new Padding(0);
            return f;
        }

        static TextBox NovoTexto(string valor)
        {
            TextBox t = new TextBox();
            t.Text = valor;
            t.Dock = DockStyle.Fill;
            return t;
        }

        static NumericUpDown NovoNumero(int min, int max, int valor)
        {
            NumericUpDown n = new NumericUpDown();
            n.Minimum = min;
            n.Maximum = max;
            n.Value = valor;
            n.Width = 80;
            return n;
        }

        static Button NovoBotao(string texto, Control pai)
        {
            Button b = new Button();
            b.Text = texto;
            b.AutoSize = true;
            b.Padding = new Padding(6, 2, 6, 2);
            pai.Controls.Add(b);
            return b;
        }

        int NovaLinha()
        {
            int linha = tblParametros.RowCount;
            tblParametros.RowCount = linha + 1;
            tblParametros.RowStyles.Add(new RowStyle(SizeType.AutoSize));
            return linha;
        }

        static Label Rotulo(string texto, int margemEsquerda)
        {
            Label l = new Label();
            l.Text = texto;
            l.AutoSize = true;
            l.Anchor = AnchorStyles.Left | AnchorStyles.Top;
            l.Margin = new Padding(margemEsquerda, 6, 8, 3);
            return l;
        }

        // Campo que ocupa a largura toda
        void AddLinha(string rotulo, Control controle)
        {
            int linha = NovaLinha();
            tblParametros.Controls.Add(Rotulo(rotulo, 3), 0, linha);
            tblParametros.Controls.Add(controle, 1, linha);
            tblParametros.SetColumnSpan(controle, 3);
        }

        // Dois campos lado a lado
        void AddPar(string rotulo1, Control controle1, string rotulo2, Control controle2)
        {
            int linha = NovaLinha();
            tblParametros.Controls.Add(Rotulo(rotulo1, 3), 0, linha);
            tblParametros.Controls.Add(controle1, 1, linha);
            tblParametros.Controls.Add(Rotulo(rotulo2, 16), 2, linha);
            tblParametros.Controls.Add(controle2, 3, linha);
        }

        void AddDica(string texto)
        {
            int linha = NovaLinha();
            Label l = new Label();
            l.Text = texto;
            l.AutoSize = true;
            l.ForeColor = SystemColors.GrayText;
            l.Margin = new Padding(3, 0, 3, 4);
            tblParametros.Controls.Add(l, 1, linha);
            tblParametros.SetColumnSpan(l, 3);
        }

        // ------------------------------------------------------------------
        // Ambiente e avisos
        // ------------------------------------------------------------------
        string TextoDetectado()
        {
            string produto = OsDetector.NomeProduto() ?? "Windows";
            string ps = psInstalado == null ? "não encontrado" : psInstalado.Major + "." + psInstalado.Minor;
            string txt = "Detectado: " + produto + " (build " + OsDetector.BuildAtual() + ")"
                + " · PowerShell " + ps
                + " · Módulo AD: " + SimNao(moduloAD)
                + " · Administrador: " + SimNao(admin);
            if (detectada == null) txt += " · Windows Server: não detectado";
            return txt;
        }

        static string SimNao(bool b)
        {
            return b ? "sim" : "não";
        }

        VersaoSO Escolhida
        {
            get { return (VersaoSO)cboVersao.SelectedItem; }
        }

        void AtualizarAviso()
        {
            List<string> avisos = new List<string>();
            VersaoSO v = Escolhida;
            if (detectada == null)
                avisos.Add("A versão do Windows Server não foi detectada; confira a versão escolhida.");
            else if (v != detectada)
                avisos.Add("A versão escolhida (" + v.Rotulo + ") é diferente da detectada (" + detectada.Rotulo + ").");
            if (psInstalado == null)
                avisos.Add("PowerShell não encontrado.");
            else if (psInstalado < v.PsMinima)
                avisos.Add("O PowerShell instalado (" + psInstalado.Major + "." + psInstalado.Minor + ") é anterior ao exigido por esta versão ("
                    + v.PsMinima.Major + "." + v.PsMinima.Minor + ").");
            if (!moduloAD)
                avisos.Add("O módulo ActiveDirectory do PowerShell não está instalado (RSAT / AD DS).");
            if (!admin)
                avisos.Add("O programa não está rodando como administrador.");

            lblAviso.Text = "⚠ " + string.Join(Environment.NewLine + "⚠ ", avisos.ToArray());
            lblAviso.Visible = avisos.Count > 0;
        }

        // ------------------------------------------------------------------
        // Execução
        // ------------------------------------------------------------------
        void Gerar()
        {
            if (executando) return;
            string saida = txtSaida.Text.Trim();
            if (saida.Length == 0)
            {
                Erro("Informe a pasta de saída.");
                txtSaida.Focus();
                return;
            }
            if (!RegexValida(txtTerceiros, "Padrão de terceiros") || !RegexValida(txtSensiveis, "Pastas sensíveis")) return;

            VersaoSO v = Escolhida;
            string script;
            try
            {
                script = ScriptCatalog.Extrair(v, saida);
            }
            catch (Exception ex)
            {
                Erro("Não foi possível gravar o script em " + saida + ":\n\n" + ex.Message);
                return;
            }

            Dictionary<string, string> p = new Dictionary<string, string>();
            AddParam(p, "Cliente", txtCliente.Text);
            AddParam(p, "Responsavel", txtResponsavel.Text);
            p["Profundidade"] = ((int)numProfundidade.Value).ToString();
            p["Saida"] = saida;
            List<string> caminhos = new List<string>();
            foreach (string linha in txtCaminhos.Lines)
                if (linha.Trim().Length > 0) caminhos.Add(linha.Trim());
            if (caminhos.Count > 0) p["Caminhos"] = string.Join("|", caminhos.ToArray());
            p["DiasInativo"] = ((int)numDias.Value).ToString();
            AddParam(p, "PadraoTerceiros", txtTerceiros.Text);
            AddParam(p, "PastasSensiveis", txtSensiveis.Text);
            // -Abrir nunca é passado: quem abre o relatório é o programa

            txtLog.Clear();
            Log("Versão: " + v + " · script: " + script);
            caminhoRelatorio = null;
            erroAmigavel = null;
            btnAbrirRel.Enabled = false;
            btnAbrirPasta.Enabled = false;

            runner = new ScriptRunner();
            runner.Linha += Runner_Linha;
            runner.Fim += Runner_Fim;
            try
            {
                runner.Iniciar(script, p);
            }
            catch (Exception ex)
            {
                Erro("Não foi possível iniciar o PowerShell:\n\n" + ex.Message);
                Status("Falhou ao iniciar o PowerShell");
                return;
            }

            executando = true;
            inicio = DateTime.Now;
            Bloquear(true);
            Status("Executando... (00:00)");
            timer.Start();
        }

        static void AddParam(Dictionary<string, string> p, string nome, string valor)
        {
            if (valor != null && valor.Trim().Length > 0) p[nome] = valor.Trim();
        }

        bool RegexValida(TextBox t, string nome)
        {
            if (t.Text.Trim().Length == 0) return true;
            try
            {
                new Regex(t.Text.Trim());
                return true;
            }
            catch (ArgumentException ex)
            {
                Erro("A expressão de \"" + nome + "\" é inválida:\n\n" + ex.Message);
                t.Focus();
                return false;
            }
        }

        void Bloquear(bool rodando)
        {
            cboVersao.Enabled = !rodando;
            tblParametros.Enabled = !rodando;
            btnGerar.Enabled = !rodando;
            btnExportar.Enabled = !rodando;
            btnCancelar.Enabled = rodando;
        }

        void CancelarExecucao()
        {
            if (runner != null && executando)
            {
                btnCancelar.Enabled = false;
                runner.Cancelar();
            }
        }

        void NaUI(MethodInvoker m)
        {
            try
            {
                if (!IsDisposed && IsHandleCreated) BeginInvoke(m);
            }
            catch (InvalidOperationException) { }
        }

        void Runner_Linha(object sender, LinhaEventArgs e)
        {
            NaUI(delegate
            {
                Log(e.Erro ? "[ERRO] " + e.Texto : e.Texto);
                if (e.Erro || e.Texto.IndexOf("Nenhum compartilhamento", StringComparison.OrdinalIgnoreCase) >= 0)
                    VerificarErroConhecido(e.Texto);
            });
        }

        void Runner_Fim(object sender, FimEventArgs e)
        {
            NaUI(delegate
            {
                timer.Stop();
                executando = false;
                Bloquear(false);
                Log("");

                if (e.Cancelado)
                {
                    Log("Execução cancelada pelo usuário.");
                    Status("Cancelado");
                    return;
                }

                if (e.CaminhoRelatorio != null && File.Exists(e.CaminhoRelatorio))
                {
                    caminhoRelatorio = e.CaminhoRelatorio;
                    btnAbrirRel.Enabled = true;
                    btnAbrirPasta.Enabled = true;
                }

                if (e.ExitCode == 0 && caminhoRelatorio != null)
                {
                    Status(TextoResumo(e.Resumo) + " (" + Decorrido() + ")");
                    if (chkAbrir.Checked) AbrirRelatorio();
                }
                else
                {
                    string s = "Falhou (código " + e.ExitCode + ")";
                    if (erroAmigavel != null) s += " — " + erroAmigavel;
                    Status(s);
                }
            });
        }

        static string TextoResumo(string linha)
        {
            // "Concluído. Usuários ativos: N | Pastas: N | Administradores: N | Sem uso: N"
            if (linha != null)
            {
                MatchCollection n = Regex.Matches(linha, @"\d+");
                if (n.Count >= 4)
                    return "Concluído: " + n[0].Value + " usuários, " + n[1].Value + " pastas, "
                        + n[2].Value + " administradores, " + n[3].Value + " sem uso";
            }
            return "Concluído";
        }

        void VerificarErroConhecido(string texto)
        {
            if (erroAmigavel != null) return;
            string curto = null, msg = null;
            if (texto.IndexOf("ActiveDirectory", StringComparison.OrdinalIgnoreCase) >= 0)
            {
                curto = "módulo ActiveDirectory ausente";
                msg = "O módulo ActiveDirectory do PowerShell não está disponível. Rode o programa num controlador de domínio ou instale o RSAT (Ferramentas do AD DS).";
            }
            else if (texto.IndexOf("Nenhum compartilhamento", StringComparison.OrdinalIgnoreCase) >= 0)
            {
                curto = "nenhum compartilhamento encontrado (preencha Caminhos)";
                msg = "Nenhum compartilhamento encontrado neste servidor. Preencha o campo Caminhos com as pastas a analisar.";
            }
            else if (texto.IndexOf("AuthorizationManager", StringComparison.OrdinalIgnoreCase) >= 0
                  || texto.IndexOf("execution policy", StringComparison.OrdinalIgnoreCase) >= 0)
            {
                curto = "bloqueado pela política de execução (GPO)";
                msg = "A política de execução do PowerShell bloqueou o script (provavelmente definida por GPO). Peça a liberação ou rode o script exportado manualmente.";
            }
            if (msg == null) return;

            erroAmigavel = curto;
            Status("Erro: " + curto);
            MessageBox.Show(this, msg, "RelatorioAD", MessageBoxButtons.OK, MessageBoxIcon.Warning);
        }

        // ------------------------------------------------------------------
        // Ações auxiliares
        // ------------------------------------------------------------------
        void EscolherPasta()
        {
            using (FolderBrowserDialog d = new FolderBrowserDialog())
            {
                d.Description = "Pasta onde o relatório será gravado";
                if (Directory.Exists(txtSaida.Text)) d.SelectedPath = txtSaida.Text;
                if (d.ShowDialog(this) == DialogResult.OK) txtSaida.Text = d.SelectedPath;
            }
        }

        void ExportarScript()
        {
            VersaoSO v = Escolhida;
            using (SaveFileDialog d = new SaveFileDialog())
            {
                d.Title = "Exportar script — " + v;
                d.Filter = "Script do PowerShell (*.ps1)|*.ps1";
                d.FileName = "Relatorio-" + v.Id + ".ps1";
                if (d.ShowDialog(this) != DialogResult.OK) return;
                try
                {
                    ScriptCatalog.Salvar(v, d.FileName);
                    Status("Script exportado: " + d.FileName);
                }
                catch (Exception ex)
                {
                    Erro("Não foi possível exportar o script:\n\n" + ex.Message);
                }
            }
        }

        void AbrirRelatorio()
        {
            if (caminhoRelatorio == null) return;
            try
            {
                Process.Start(caminhoRelatorio);
            }
            catch (Exception ex)
            {
                Erro("Não foi possível abrir o relatório:\n\n" + ex.Message);
            }
        }

        void AbrirPasta()
        {
            if (caminhoRelatorio == null) return;
            try
            {
                Process.Start("explorer.exe", "/select,\"" + caminhoRelatorio + "\"");
            }
            catch (Exception ex)
            {
                Erro("Não foi possível abrir a pasta:\n\n" + ex.Message);
            }
        }

        void Log(string linha)
        {
            txtLog.AppendText(linha + Environment.NewLine);
        }

        void Status(string texto)
        {
            lblStatus.Text = texto;
        }

        void Erro(string msg)
        {
            MessageBox.Show(this, msg, "RelatorioAD", MessageBoxButtons.OK, MessageBoxIcon.Error);
        }

        string Decorrido()
        {
            TimeSpan t = DateTime.Now - inicio;
            return ((int)t.TotalMinutes).ToString("00") + ":" + t.Seconds.ToString("00");
        }

        protected override void OnFormClosing(FormClosingEventArgs e)
        {
            if (executando)
            {
                DialogResult r = MessageBox.Show(this, "O relatório ainda está sendo gerado. Cancelar a execução e fechar?",
                    "RelatorioAD", MessageBoxButtons.YesNo, MessageBoxIcon.Question, MessageBoxDefaultButton.Button2);
                if (r != DialogResult.Yes)
                {
                    e.Cancel = true;
                    return;
                }
                runner.Cancelar();
            }
            base.OnFormClosing(e);
        }
    }
}
