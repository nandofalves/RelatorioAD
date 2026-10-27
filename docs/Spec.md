# Spec: RelatorioAD — Relatório de acessos multi-versão do Windows Server

> Baseada em `docs/PRD.md` (aprovado) · Data: 2026-10-27

## Overview
Programa WinForms em C# (CLR 2 + CLR 4) que detecta a versão do Windows Server, deixa o técnico escolher a versão e os parâmetros, extrai o script PowerShell embutido daquela versão para uma pasta visível e o executa com `powershell.exe -File`, mostrando o log ao vivo. Os 7 scripts saem de 3 bases de código (PS 5.1, PS 3/4 e PS 2), geradas pelo build com a versão gravada em cada um.

## Estrutura final
```
C:\RelatorioAD\
├── CLAUDE.md
├── build.ps1                       # gera scripts, valida, compila, (assina), escaneia
├── referencia\Relatorio-AD-Permissoes-2022.ps1   # fornecido pelo usuário (não alterar)
├── docs\PRD.md · Spec.md · progress.md
├── scripts\
│   ├── Relatorio-PS51.ps1          # base 2016 / 2019 / 2022 / 2025
│   ├── Relatorio-PS3.ps1           # base 2012 / 2012 R2
│   └── Relatorio-PS2.ps1           # base 2008 R2
├── tools\Test-ScriptCompat.ps1     # lint de compatibilidade por versão de PowerShell
├── src\
│   ├── Program.cs
│   ├── MainForm.cs
│   ├── OsDetector.cs
│   ├── ScriptCatalog.cs
│   ├── ScriptRunner.cs
│   ├── AssemblyInfo.cs
│   ├── app.manifest
│   └── RelatorioAD.exe.config
├── build\ (gerado: scripts\Relatorio-<versão>.ps1)
└── dist\  (gerado: RelatorioAD.exe + RelatorioAD.exe.config + RelatorioAD.zip)
```

> **Ajuste em relação ao PRD (RNF1):** um .exe compilado para o CLR 2 só roda no 2012+ (que não tem o .NET 3.5 ativado) se tiver ao lado o `RelatorioAD.exe.config` com `<supportedRuntime version="v4.0"/>`. Sem ele, o Windows pede para instalar o .NET 3.5. A entrega passa a ser **o .exe + o .config** (e um .zip com os dois). Continua sem instalação e sem dependências.

---

## Fase 1: Scripts base e lint de compatibilidade

Pré-requisito: `referencia\Relatorio-AD-Permissoes-2022.ps1` salvo pelo usuário. Os três scripts base partem dele e precisam gerar **o mesmo HTML e os mesmos CSVs**. Mudanças comuns às três bases:

- Novo parâmetro `[string]$VersaoSO = ''` (o build grava o valor padrão em cada cópia; ver Fase 2). Aparece no bloco "Como o levantamento foi feito": `... executado no servidor X (Windows Server <VersaoSO>) em ...`.
- `-Caminhos` aceita também um único texto separado por `|` (caractere proibido em caminhos do Windows), porque o `powershell.exe -File` não passa arrays. Logo após o `param()`: se `@($Caminhos).Count -eq 1` e o item contém `|`, dividir por `|`, aplicar `Trim()` e descartar vazios.
- A linha final `Relatório: <caminho>` e a linha `Concluído. Usuários ativos: N | Pastas: N | Administradores: N | Sem uso: N` **mantêm exatamente esse formato**: o C# lê as duas para preencher o resumo.
- Arquivos salvos em UTF-8 com BOM.

### `scripts/Relatorio-PS51.ps1` (CRIAR)
- Cópia da referência + as mudanças comuns acima. Nenhuma outra alteração.

### `scripts/Relatorio-PS3.ps1` (CRIAR)
- Parte do PS51. `#Requires -Version 3.0`.
- Troca a coleta de subpastas (`Get-ChildItem -Recurse -Depth`) por `function Get-Subpastas([string]$raiz, [int]$niveis, [ref]$erros)`: busca em largura, nível a nível, com `Get-ChildItem -LiteralPath $p -Directory -Force -ErrorAction SilentlyContinue -ErrorVariable ev`, juntando os erros na mesma lista `$erros` (mesmo formato `Pasta`/`Erro`). Retorna os `FullName` na mesma ordem lógica.
- Não seguir pontos de junção nem links simbólicos (`$_.Attributes -band [IO.FileAttributes]::ReparsePoint`), igual ao comportamento padrão do `-Recurse`.
- O resto fica igual (`Get-SmbShare`, `[pscustomobject]`, `[ordered]` e `-in` existem no PS 3).

### `scripts/Relatorio-PS2.ps1` (CRIAR)
- Parte do PS3. `#Requires -Version 2.0`. Substituições obrigatórias:

| Construção do PS 3+/CLR 4 | Equivalente no PS 2.0 |
|---|---|
| `[pscustomobject]@{...}` | `function NovoObj([hashtable]$h) { New-Object PSObject -Property $h }`. Onde a ordem das colunas importa (CSVs), exportar com `Select-Object` explícito na ordem do original |
| `[ordered]@{}` | `New-Object System.Collections.Specialized.OrderedDictionary` |
| `HashSet[string]` / `SortedSet[string]` | hashtable usada como conjunto (`$h[$x] = 1`; `@($h.Keys \| Sort-Object)` onde o original usa `SortedSet`) |
| `[System.Net.WebUtility]::HtmlEncode` | `HE` próprio: troca `&`, `<`, `>`, `"` e `'` por entidades, nessa ordem |
| `-in` / `-notin` | `-contains` / `-notcontains` com os operandos invertidos |
| `Where-Object Nome -eq $x` / `ForEach-Object Nome` | blocos `{ $_.Nome -eq $x }` / `{ $_.Nome }` |
| `Get-ChildItem -Directory` | `Get-ChildItem ... \| Where-Object { $_.PSIsContainer }` |
| `Get-Acl -LiteralPath` | `(Get-Item -LiteralPath $p -Force).GetAccessControl('Access')` |
| enumeração de membros (`$lista.Prop`) | `ForEach-Object { $_.Prop }` |
| `.Count` em valor que pode ser escalar ou nulo | sempre `@(...).Count` |
| `Import-Module SmbShare`, `Get-SmbShare`, `Get-SmbShareAccess` | WMI (abaixo) |

- **Compartilhamentos via WMI**, na função `Get-Compartilhamentos` → objetos `Nome`, `Path`, `Acl` (o mesmo formato de `$alvos`):
  - `Get-WmiObject Win32_Share -Filter 'Type = 0'` (somente discos, sem os administrativos), excluindo `NETLOGON`, `SYSVOL` e `print$`, e com a mesma ordenação do original (nomes terminados em `$` por último).
  - ACL: `Get-WmiObject Win32_LogicalShareSecuritySetting -Filter "Name='<nome com ' escapado>'"` → `.GetSecurityDescriptor().Descriptor.DACL`. Para cada ACE: `Sid = $ace.Trustee.SIDString`, `Conta = domínio\nome do Trustee`, `Tipo = 'Allow'` se `AceType -eq 0`, `'Deny'` se `-eq 1`, `Perm` pela máscara: `2032127` → `Full`, `1245631` → `Change`, `1179817` → `Read`, qualquer outra → `Custom`.
  - DACL nula ou ausente (compartilhamento sem descritor) = `Everyone` / `S-1-1-0` / `Allow` / `Full`, que é o comportamento do Windows.
- `Get-ADGroupMember -Recursive`, `Get-ADUser -LDAPFilter` e `LastLogonDate` existem no módulo AD do 2008 R2: mantidos.

### `tools/Test-ScriptCompat.ps1` (CRIAR)
- `param([string]$Path, [ValidateSet('2.0','3.0','5.1')][string]$Alvo)`.
- Faz o parse com `[System.Management.Automation.Language.Parser]::ParseFile($Path, [ref]$tokens, [ref]$erros)`. Qualquer erro de sintaxe → falha.
- Percorre o AST (`FindAll`) e os tokens e reprova o que não existe no alvo:
  - Alvo 3.0: parâmetro `-Depth` em `Get-ChildItem`.
  - Alvo 2.0: tudo do 3.0 + parâmetros `-Directory`/`-File`/`-LiteralPath` em `Get-Acl`; tipos `pscustomobject`, `ordered`, `HashSet`, `SortedSet`, `WebUtility`; operadores `-in`/`-notin`; comandos `Get-SmbShare`, `Get-SmbShareAccess`, `Get-CimInstance`; `Where-Object`/`ForEach-Object` sem bloco de script.
- Saída: lista `Linha: problema`, com exit code 1 se houver falhas.

### Verificação Fase 1
- [x] `tools\Test-ScriptCompat.ps1` aprova as 3 bases nos alvos 5.1 / 3.0 / 2.0.
- [x] Teste negativo do lint: um arquivo temporário com `[pscustomobject]@{}` e `-in` é reprovado no alvo 2.0.
- [x] `Compare-Object` entre a referência e `Relatorio-PS51.ps1` mostra só as mudanças comuns.
- [x] Os 3 arquivos estão em UTF-8 com BOM (bytes `EF BB BF`).

---

## Fase 2: Build (geração, compilação, assinatura e scan)

### `build.ps1` (CRIAR)
- `param([string]$CertThumbprint, [switch]$SemScan)`. Sempre usar `$PSScriptRoot` e caminhos absolutos.
- Etapas, com `throw` na primeira falha:
  1. **Gerar scripts por versão** em `build\scripts\`. Mapa: `2008R2`→PS2, `2012`/`2012R2`→PS3, `2016`/`2019`/`2022`/`2025`→PS51. Em cada cópia, trocar o padrão do parâmetro `$VersaoSO = ''` pelo rótulo (`'2008 R2'`, `'2012'`, ...). Salvar em UTF-8 com BOM.
  2. **Lint**: rodar `tools\Test-ScriptCompat.ps1` em cada um dos 7 scripts com o alvo correspondente.
  3. **Compilar** com `C:\Windows\Microsoft.NET\Framework64\v3.5\csc.exe`:
     `/target:winexe /platform:anycpu /optimize+ /warnaserror+ /win32manifest:src\app.manifest /out:dist\RelatorioAD.exe /r:System.dll /r:System.Drawing.dll /r:System.Windows.Forms.dll /resource:build\scripts\Relatorio-2008R2.ps1,RelatorioAD.Scripts.2008R2.ps1 ... src\*.cs`
     (um `/resource` por versão, com o nome lógico `RelatorioAD.Scripts.<id>.ps1`). Copiar `src\RelatorioAD.exe.config` para `dist\`.
  4. **Assinar** (só com `-CertThumbprint`): `Set-AuthenticodeSignature` com `-TimestampServer http://timestamp.digicert.com -HashAlgorithm SHA256` no .exe e nos 7 scripts de `build\scripts` **antes** da etapa 3 (os scripts embutidos já vão assinados). A ordem real fica 1 → 2 → assinar scripts → 3 → assinar exe.
  5. **Scan do Defender** (exceto com `-SemScan`): `& "$env:ProgramFiles\Windows Defender\MpCmdRun.exe" -Scan -ScanType 3 -File "<dist>" -DisableRemediation`. Exit code diferente de 0 → falha com a mensagem do MpCmdRun.
  6. **Empacotar**: `dist\RelatorioAD.zip` com o .exe e o .config (via `System.IO.Compression.ZipFile`, `Add-Type -AssemblyName System.IO.Compression.FileSystem`).
- Imprimir no final: tamanho do .exe, SHA256 do .exe e do .zip.

### Verificação Fase 2
- [x] `.\build.ps1 -SemScan` termina sem erros e sem avisos do csc (só depois de existirem os arquivos da Fase 3; antes disso, validar só as etapas 1 e 2). *Etapas 1 e 2 ok no projeto; etapas 3 e 6 validadas numa cópia com um `.cs` mínimo.*
- [x] `build\scripts` tem 7 arquivos, cada um com o `$VersaoSO` certo.

---

## Fase 3: Aplicativo C#

Todas as classes no namespace `RelatorioAD`, com sintaxe C# 3 (o csc v3.5 **não** aceita `async/await`, parâmetros opcionais, `dynamic` nem interpolação `$""`).

### `src/RelatorioAD.exe.config` (CRIAR)
```xml
<?xml version="1.0" encoding="utf-8"?>
<configuration>
  <startup>
    <supportedRuntime version="v4.0"/>
    <supportedRuntime version="v2.0.50727"/>
  </startup>
</configuration>
```

### `src/app.manifest` (CRIAR)
- `assemblyIdentity` com `name="Compumaq.RelatorioAD"` e `version` igual à do AssemblyInfo.
- `requestedExecutionLevel level="requireAdministrator" uiAccess="false"`.
- `compatibility/supportedOS` com os GUIDs do Windows 7/2008 R2, 8/2012, 8.1/2012 R2 e 10/11/2016+ (para o `Environment.OSVersion` não mentir no 2012 R2+).
- `dpiAware` = true.

### `src/AssemblyInfo.cs` (CRIAR)
- `AssemblyTitle("RelatorioAD")`, `AssemblyDescription("Relatório de usuários do AD e permissões de pastas")`, `AssemblyCompany("Compumaq · Werner Informática")`, `AssemblyProduct("RelatorioAD")`, `AssemblyCopyright("© 2026 Compumaq")`, `AssemblyVersion("1.0.0.0")`, `AssemblyFileVersion("1.0.0.0")`, `ComVisible(false)`.

### `src/ScriptCatalog.cs` (CRIAR)
- `class VersaoSO { string Id; string Rotulo; int Build; Version PsMinima; string Recurso; }` e `override ToString()` → `"Windows Server " + Rotulo`.
- `static class ScriptCatalog`:
  - `static readonly VersaoSO[] Todas`: 2008R2 (build 7600/7601, PS 2.0), 2012 (9200, PS 3.0), 2012R2 (9600, PS 4.0), 2016 (14393, PS 5.1), 2019 (17763, 5.1), 2022 (20348, 5.1), 2025 (26100, 5.1).
  - `static string LerScript(VersaoSO v)` → lê o recurso com `Assembly.GetManifestResourceStream` e devolve o texto. Se o recurso não existir → `InvalidOperationException` com mensagem clara.
  - `static string Extrair(VersaoSO v, string pastaSaida)` → grava `<pastaSaida>\scripts\Relatorio-<Id>.ps1` com `new UTF8Encoding(true)` e devolve o caminho.

### `src/OsDetector.cs` (CRIAR)
- `static class OsDetector`:
  - `static int BuildAtual()` → `CurrentBuild` (ou `CurrentBuildNumber`) de `HKLM\SOFTWARE\Microsoft\Windows NT\CurrentVersion`, com fallback para `Environment.OSVersion.Version.Build`.
  - `static bool EhServidor()` → `InstallationType` = `Server`/`Server Core`, ou `ProductName` contém `Server`.
  - `static string NomeProduto()` → `ProductName`.
  - `static VersaoSO Detectar()` → a versão de `ScriptCatalog.Todas` com o maior `Build` ≤ build atual, ou `null` se não for servidor ou se o build for menor que 7600.
  - `static Version VersaoPowerShell()` → `PowerShellVersion` de `HKLM\SOFTWARE\Microsoft\PowerShell\3\PowerShellEngine`, com fallback para `...\PowerShell\1\PowerShellEngine`. `null` se não encontrar.
  - `static bool ModuloADInstalado()` → existe `%SystemRoot%\System32\WindowsPowerShell\v1.0\Modules\ActiveDirectory`.
  - `static bool EhAdministrador()` → `WindowsPrincipal.IsInRole(WindowsBuiltInRole.Administrator)`.

### `src/ScriptRunner.cs` (CRIAR)
- `class ScriptRunner`, com os eventos `event EventHandler<LinhaEventArgs> Linha` (texto e `bool Erro`) e `event EventHandler<FimEventArgs> Fim` (`ExitCode`, `CaminhoRelatorio`, `Resumo`, `Cancelado`).
- `void Iniciar(string scriptPath, IDictionary<string,string> parametros)`:
  - `ProcessStartInfo`: `FileName = %SystemRoot%\System32\WindowsPowerShell\v1.0\powershell.exe`; `Arguments = -NoProfile -NonInteractive -ExecutionPolicy Bypass -File "<script>" -Nome "valor" ...`, com cada valor passando por `Quote()` (aspas duplas, `\` antes de aspas e barras finais conforme as regras do CommandLineToArgvW); `UseShellExecute=false`, `RedirectStandardOutput/Error=true`, `CreateNoWindow=true`; `StandardOutputEncoding` e `StandardErrorEncoding` = `Encoding.GetEncoding(CultureInfo.CurrentCulture.TextInfo.OEMCodePage)`.
  - `BeginOutputReadLine`/`BeginErrorReadLine`. Cada linha vira um evento `Linha`; o runner memoriza as linhas que começam com `Relatório: ` e `Concluído.`.
  - `EnableRaisingEvents = true` + `Exited` → dispara `Fim` (depois de `WaitForExit()` sem timeout, para esvaziar os buffers).
- `void Cancelar()` → `Process.Kill()` dentro de try/catch e `Cancelado = true`.
- Os eventos chegam em outra thread: quem atualiza a UI é o `MainForm`, via `BeginInvoke`.

### `src/MainForm.cs` (CRIAR)
- Interface montada em código (sem designer), com `TableLayoutPanel`, fonte Segoe UI 9 e janela de 900×680 redimensionável. Título: `RelatorioAD 1.0 — Relatório de acessos do AD`.
- **Topo**: o texto `Detectado: <ProductName> (build N) · PowerShell X.Y · Módulo AD: sim/não · Administrador: sim/não`, e um `ComboBox` (DropDownList) com `ScriptCatalog.Todas`, já com `OsDetector.Detectar()` selecionado (ou 2022 se não detectar).
- Rótulo de aviso amarelo, visível quando: a versão escolhida ≠ a detectada; o PowerShell instalado < `PsMinima` da escolhida; o módulo AD não existe; ou o programa não está rodando como administrador.
- **Parâmetros** (com os mesmos padrões do script): Cliente (vazio = NetBIOS do domínio), Responsável, Profundidade (`NumericUpDown` de 0 a 10, padrão 2), Pasta de saída (`TextBox` + botão `...` com `FolderBrowserDialog`, padrão `C:\Relatorio_AD`), Caminhos (`TextBox` multilinha, um por linha, opcional), Dias inativo (`NumericUpDown` de 1 a 3650, padrão 90), Padrão de terceiros, Pastas sensíveis, checkbox "Abrir relatório ao terminar" (marcado).
- **Botões**: `Gerar relatório`, `Cancelar` (habilitado só durante a execução), `Abrir relatório` e `Abrir pasta` (habilitados quando houver relatório), `Exportar script...` (`SaveFileDialog` → grava o script da versão escolhida em UTF-8 com BOM).
- **Log**: `TextBox` multilinha, somente leitura, fonte Consolas, rolagem automática, e linhas de erro prefixadas com `[ERRO]`.
- **Barra de status**: `Pronto` / `Executando... (mm:ss)` (com um `Timer` de 1 s) / `Concluído: N usuários, N pastas, N administradores, N sem uso` / `Falhou (código N)` / `Cancelado`.
- `Gerar relatório` → valida (pasta de saída não vazia e regex válidas, testadas com `new Regex(...)`) → bloqueia os campos → `ScriptCatalog.Extrair` → `ScriptRunner.Iniciar`. Caminhos: juntar as linhas não vazias com `|`. Parâmetros vazios não são passados. `-Abrir` nunca é passado, porque quem abre é o C#, com `Process.Start(caminho)` se o checkbox estiver marcado.
- Erros conhecidos no log → mensagem amigável na barra de status e um `MessageBox` uma única vez: `ActiveDirectory` (módulo AD ausente), `Nenhum compartilhamento` (sugerir preencher Caminhos), `AuthorizationManager`/`execution policy` (política de execução definida por GPO).
- Fechar a janela durante a execução → perguntar e cancelar.

### `src/Program.cs` (CRIAR)
- `[STAThread] static void Main()`: `Application.EnableVisualStyles()`, `SetCompatibleTextRenderingDefault(false)`, `Application.ThreadException` e `AppDomain.UnhandledException` → `MessageBox` com a mensagem, e por fim `Application.Run(new MainForm())`.

### Verificação Fase 3
- [x] `.\build.ps1 -SemScan` compila sem avisos.
- [x] O .exe abre nesta máquina (Windows 11), mostra "não detectado" sem quebrar e seleciona 2022 por padrão.
- [x] `Exportar script...` gera para cada versão um arquivo idêntico ao de `build\scripts` (mesmo SHA256). *Testado chamando `ScriptCatalog.Salvar` (o mesmo método do botão) por reflexão.*
- [x] Rodando o 2022 aqui (sem AD), o log mostra o erro do módulo `ActiveDirectory` e aparece a mensagem amigável, sem travar a janela.
- [x] O .exe também abre com o .config removido? **Não** deve ser exigido; só registrar o comportamento no progress.md.

---

## Fase 4: Antivírus e entrega

### `build.ps1` (MODIFICAR) — o scan passa a ser padrão
- Rodar `.\build.ps1` (com scan). Se o Defender acusar algo, **não** contornar com exclusões: investigar qual padrão disparou e corrigir no código.

### `docs/ENTREGA.md` (CRIAR)
- Como usar: copiar o `RelatorioAD.zip` para o servidor, extrair, rodar o .exe como administrador e escolher a versão.
- Antivírus: o que o programa faz e não faz (sem ofuscação, sem download, scripts legíveis em `<Saída>\scripts`), o SHA256 da versão e como enviar um falso positivo ao Microsoft WDSI (https://www.microsoft.com/wdsi/filesubmission) e a outros fabricantes.
- Certificado: como assinar quando houver um (`.\build.ps1 -CertThumbprint <x>`) e a alternativa interna com certificado autoassinado + GPO para "Editores Confiáveis".

### Verificação Fase 4
- [x] `.\build.ps1` (com scan) termina com o Defender limpo. *O scan passou a ser a última etapa (depois do zip) e cobre `dist\` e `build\scripts\`.*
- [x] `dist\RelatorioAD.zip` contém exatamente `RelatorioAD.exe` e `RelatorioAD.exe.config`.
- [x] `Get-AuthenticodeSignature dist\RelatorioAD.exe` → `NotSigned` (esperado, sem certificado) e a etapa de assinatura é pulada sem erro.

## Testes
- Pulados por decisão do usuário. Valem só os itens de verificação acima (lint/parser, build, abertura da janela, scan do Defender). A validação real fica a cargo do usuário nos servidores, começando pelo 2022 (referência) e depois pelo 2008 R2 (maior risco).

## Ordem de implementação
1. Fase 1 (scripts + lint) → 2. Fase 2 (build) → 3. Fase 3 (app) → 4. Fase 4 (antivírus + entrega).

## Notas / Não fazer
- Não usar PS2EXE, UPX, ofuscadores, `-EncodedCommand`, `-WindowStyle Hidden`, base64 nem download de nada pelo .exe.
- Não usar Visual Studio, NuGet nem libs externas. Só o `csc.exe` v3.5 e as DLLs do Framework.
- Não alterar `referencia\`: ela é a fonte da verdade do HTML e dos cálculos.
- Não executar o script dentro do processo (hospedar `System.Management.Automation`): a versão do assembly muda entre o PS 2 e o 5.1, e o `powershell.exe -File` é o caminho mais simples e transparente.
- Não reproduzir o script de referência inteiro pelo chat: trabalhar a partir do arquivo em disco.
- Decisões já tomadas: 2008 R2 sempre usa a base PS2 (mesmo com WMF 5.1); 2025 usa a base PS51.
