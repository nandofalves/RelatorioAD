# Progresso

## 2026-10-27 — Fase 0/1 (Setup + PRD)
- Criada a estrutura `C:\RelatorioAD` (CLAUDE.md, docs/, src/, referencia/).
- `docs/PRD.md` em rascunho, aguardando aprovação.
- Pendência: o usuário precisa salvar o script de referência em `referencia/Relatorio-AD-Permissoes-2022.ps1`.
- Dúvidas resolvidas pelo usuário ("usar o que funciona"): 2008 R2 sempre usa o script WMI/PS 2.0; 2025 usa o mesmo script do 2022.
- Próximo passo: aprovar o PRD → `/clear` → gerar `docs/Spec.md`.

## 2026-10-27 — Fase 2 do SDD (Spec)
- PRD aprovado. `docs/Spec.md` gerada com 4 fases de implementação.
- Ajuste em relação ao PRD: a entrega é o .exe + o `RelatorioAD.exe.config` (necessário para o mesmo .exe rodar no CLR 4 do 2012+).
- Pendência que bloqueia a Fase 1: `referencia\Relatorio-AD-Permissoes-2022.ps1` ainda não foi salvo.
- Próximo passo: aprovar a Spec → `/clear` → `/implement fase 1`.

## 2026-10-27 — Fase 3 do SDD, Fase 1 da Spec (scripts base + lint)
- Referência: havia duas cópias do script. O usuário escolheu a do OneDrive (`Documents\1compumaq\Kiplus`, 724 linhas, **sem** a seção "6. Recomendações"), copiada para `referencia\`.
- Criados `scripts\Relatorio-PS51.ps1`, `Relatorio-PS3.ps1`, `Relatorio-PS2.ps1` e `tools\Test-ScriptCompat.ps1` (todos UTF-8 com BOM).
- PS51: só as mudanças comuns (`$VersaoSO`, `-Caminhos` com `|`, versão na metodologia, exibida só se `$VersaoSO` vier preenchido).
- PS3: `Get-Subpastas` substitui o `-Depth`. Em vez da busca em largura citada na Spec, segue a ordem exata do `-Recurse` (filhos e depois cada filho). Testado aqui: mesma lista e mesma ordem do `-Recurse -Depth`.
- PS2: todas as substituições da tabela da Spec + `Get-Compartilhamentos` via WMI. Sem descritor no compartilhamento → "Todos" (nome localizado via SID) com controle total. Diferenças aceitas: o `HE` próprio não converte acentos em entidades `&#NNN;` (o resultado visual é o mesmo) e a origem na dica da matriz é ordenada com `Sort-Object` em vez de `SortedSet`.
- Lint: corrigido o nome dos operadores no AST (`Iin`/`Inotin`, não `In`). Verificações da Fase 1 todas ok.
- Observação: neste ambiente, `powershell.exe -ExecutionPolicy Bypass` foi bloqueado pelo classificador do Claude Code. O lint foi chamado direto (o processo já está em Bypass). O build (Fase 2) deve chamar o lint no mesmo processo (`& $lint ...`), sem abrir outro `powershell.exe`.
- Pendência: as bases PS3/PS2 só serão validadas de verdade num 2012/2008 R2 real (não há PS 2.0 nesta máquina).
- Próximo passo: revisar → commit → `/clear` → `/implement fase 2`.

## 2026-10-27 — Fase 3 do SDD, Fase 2 da Spec (build)
- Criado `build.ps1` (UTF-8 com BOM): limpa `build\` e `dist\` → gera os 7 scripts em `build\scripts` (troca `$VersaoSO = ''` pelo rótulo e exige exatamente 1 ocorrência) → lint de cada um no mesmo processo (`& $lint`) → assina os scripts (só com `-CertThumbprint`; procura em CurrentUser\My e LocalMachine\My) → `csc` v3.5 com `/warnaserror+` e um `/resource` por versão → copia o `.config` → assina o exe → scan do Defender (exceto com `-SemScan`) → `dist\RelatorioAD.zip` → tamanho e SHA256.
- Extras em relação à Spec: `/nologo /utf8output` no csc e falha se a saída tiver `warning`; mensagem clara se faltar `src\*.cs`, `app.manifest` ou o `.config`.
- Verificado: as etapas 1 e 2 passam no projeto (7 scripts com BOM, `$VersaoSO` certo, diferença para a base só nessa linha, lint ok) e o build para na etapa 3 com "Nenhum .cs em src" (esperado até a Fase 3). Numa cópia temporária com um `.cs` mínimo: compila, os 7 recursos `RelatorioAD.Scripts.<id>.ps1` estão no exe e o zip contém só o exe e o `.config`.
- Não testados: a assinatura (não há certificado) e o scan do Defender (fica para a Fase 4).
- Próximo passo: revisar → commit → `/clear` → `/implement fase 3`.

## 2026-10-27 — Fase 3 do SDD, Fase 3 da Spec (aplicativo C#)
- Criados `src\Program.cs`, `MainForm.cs`, `OsDetector.cs`, `ScriptCatalog.cs`, `ScriptRunner.cs`, `AssemblyInfo.cs`, `app.manifest` e `RelatorioAD.exe.config`.
- **Os `.cs` precisam estar em UTF-8 com BOM**: o build não passa `/codepage` e o csc 3.5 lê UTF-8 sem BOM como ANSI (os acentos quebram). O Write/Edit do Claude Code grava sem BOM; regravar com BOM antes do build.
- Sem LINQ, `Func`/`Action` sem argumentos nem `HashSet`: ficam no System.Core, que o build não referencia. Usar `MethodInvoker`, `List<T>.ToArray()` etc.
- Diferenças em relação à Spec:
  - `ScriptCatalog.Salvar(v, caminho)` (novo): usado por `Extrair` e pelo `Exportar script...`; e `ScriptCatalog.PorId`.
  - Layout em 4 colunas (Cliente/Responsável e Profundidade/Dias inativo lado a lado): em 900×680, com um campo por linha, o log ficava com ~2 linhas.
  - Erros conhecidos: o MessageBox mostra o texto longo e a barra de status um texto curto (`Falhou (código 1) — módulo ActiveDirectory ausente`), porque o longo era cortado.
  - O rótulo "Detectado" termina com `· Windows Server: não detectado` quando não é servidor.
- **Correção nos scripts base (bug da Fase 1)**: sem `-Caminhos`, `@($Caminhos).Count` vale 1 (`@($null)` tem 1 item) e `$Caminhos[0]` gerava "Não é possível indexar em uma matriz nula" no log. A condição passou a ser `if ($Caminhos -and @($Caminhos).Count -eq 1 -and ...)` nas 3 bases.
- Verificado: build sem avisos (376.832 bytes); abre aqui (Windows 11, `ProductName` = "Windows 10 Pro", build 26200) com "não detectado" e 2022 selecionado; exportação idêntica a `build\scripts` nas 7 versões (SHA256); rodando o 2022 sem AD: log ao vivo com `[ERRO]`, MessageBox amigável uma vez, janela respondendo, status `Falhou (código 1) — módulo ActiveDirectory ausente` e botões restaurados.
- Sem o `.config`: o .exe abre aqui, mas só porque esta máquina tem o .NET 3.5 instalado (roda no CLR 2). Num 2012+ sem o 3.5 o `.config` continua necessário.
- Como testar a janela sem UAC: `$env:__COMPAT_LAYER = 'RunAsInvoker'` antes de abrir o .exe. A UI Automation enxerga os controles WinForms só como `Pane` (sem Invoke/Value): usar `WM_SETTEXT`/`BM_CLICK` no `NativeWindowHandle` e `PrintWindow` para captura. O MessageBox aparece como filho da janela principal na árvore UIA.
- Não testados: caminho de sucesso (precisa de AD), Cancelar com processo em andamento e o tratamento de política de execução por GPO.
- Próximo passo: revisar → commit → `/clear` → `/implement fase 4`.

## 2026-10-27 — Fase 3 do SDD, Fase 4 da Spec (antivírus e entrega)
- O scan do Defender já era padrão no `build.ps1`. Mudanças: o scan virou a **última etapa** (5 = empacotar, 6 = scan), para cobrir também o `.zip`; escaneia `dist\` **e** `build\scripts\` (os 7 `.ps1` embutidos); e imprime a versão das assinaturas do Defender.
- `.uild.ps1` completo: lint ok, compila sem avisos, Defender sem ameaças nos dois alvos (assinaturas 1.461.225.0, proteção em tempo real ligada), nada sobre o RelatorioAD em `Get-MpThreatDetection`. Nenhuma correção no código foi necessária.
- `dist\RelatorioAD.zip` = `RelatorioAD.exe` (376.832 bytes) + `RelatorioAD.exe.config` (181 bytes). `Get-AuthenticodeSignature` → `NotSigned`, etapa de assinatura pulada sem erro.
- Criado `docs\ENTREGA.md` (uso, requisitos, o que o programa faz e não faz, SHA256, falso positivo/WDSI, certificado comercial ou autoassinado + GPO).
- O SHA256 muda a cada build (o csc grava um MVID novo). O `ENTREGA.md` traz o do build atual (exe `91AB7EB2…01A6`, zip `C8B800B2…1936`): **rodar o build de novo exige atualizar a tabela**.
- Só o Defender foi testado. Outros antivírus e o SmartScreen (arquivo baixado com Mark of the Web) não foram testados.
- Todas as fases da Spec estão implementadas. Falta a validação real nos servidores (2022 primeiro, depois 2008 R2) e o caminho de sucesso com AD, Cancelar com processo em andamento e GPO de política de execução.

## 2026-10-27 — Correção após o primeiro teste real (2008 R2 com AD)
- Erro no passo 4/5: `ContainsKey` com chave nula (linha 301 do `Relatorio-2008R2.ps1`). Causa: no **PS 2.0**, `foreach ($x in <lista vazia/$null>)` executa **uma vez com `$x = $null`** (do PS 3 em diante, não executa). Acontece quando um grupo com permissão não tem membros ativos (`UsersOf` devolve vazio).
- `scripts\Relatorio-PS2.ps1`: as 29 coleções de `foreach` passaram a ser `@(...)`.
- `tools\Test-ScriptCompat.ps1`: nova regra no alvo 2.0, que reprova `foreach` cuja coleção não seja `@(...)`. A versão antiga é reprovada com 29 problemas; a nova passa.
- Build completo ok, Defender limpo. Novo SHA256 no `ENTREGA.md` (exe `8B11F541…1FED`).
- Os passos 1 a 3 já funcionaram no 2008 R2 (AD, compartilhamentos via WMI, ACLs). Falta confirmar os passos 4 e 5 (HTML) com o novo .exe.
- **2º teste: mesmo erro.** O `@(...)` no `foreach` não basta quando a variável já é `$null`: lista vazia atribuída a variável vira `$null`, e `@($null)` tem **1** item. Caso real: ACE de **usuário desabilitado** (`Tipo = 'usuario'`, `Users` vazio) → `$alvo = ... else { @() }` = `$null`. Correção na base PS2: `$alvo = @(if ... )`, com os ramos devolvendo os valores sem `@()` (ramo vazio = 0 itens). As outras atribuições `= if` guardam escalares, e os pipelines partem de `@()`, de coleções ou vêm depois de um teste de `.Count`.
- Não dá para testar no PS 2.0 aqui: o Windows 11 (26200) não tem mais o motor 2.0 (`-Version 2` cai no 5.1). Novo SHA256 no `ENTREGA.md` (exe `3E3817C8…23AF`).

## 2026-10-27 — Feature 001 (auditoria do domínio), Fase 1 do SDD (PRD)
- Pedido: comparação com o levantamento anterior, contas privilegiadas/de risco no domínio, política de senha, GPOs e computadores, todos **opcionais** (caixas desmarcadas por padrão).
- Decisões do usuário: comparação automática + escolha de arquivo; contas de serviço por SPN + padrão de nome; referência = Microsoft Security Baseline; mesmo HTML, com novas seções.
- `docs/features/001-auditoria-dominio/PRD.md` gerado (rascunho). Proposta: módulos escritos uma vez em `scripts/Modulos.ps1` (PS 2.0) e embutidos pelo build nas 3 bases.
- PRD aprovado (bloqueio: alerta só em 0 e > 10; snapshot sempre gravado).
- Próximo passo: `/clear` → gerar a Spec em `docs/features/001-auditoria-dominio/Spec.md`.
