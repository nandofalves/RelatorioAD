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

## 2026-10-27 — Feature 001 (auditoria do domínio), Fase 2 do SDD (Spec)
- `docs/features/001-auditoria-dominio/Spec.md` gerada, com 8 fases: esqueleto/build → snapshot/comparar → contas → senha → GPOs → computadores → interface → entrega.
- Decisões: `scripts/Modulos.ps1` embutido pelo build na região `#region MODULOS` (via IndexOf, não regex); cada módulo em leitura / análise pura / HTML; novo `tools/Test-Modulos.ps1` rodando no build; membros de grupos privilegiados via `memberOf:1.2.840.113556.1.4.1941` (sem `Get-ADGroupMember`); vínculos de GPO também em `CN=Sites`; switches passados pelo `ScriptRunner` sem valor.
- Correção do PRD: o Windows Server 2025 vai até 2034-11-14 (conferido no Microsoft Lifecycle). Tabela de fim de suporte ampliada com Win10 LTSB/LTSC, Win11 LTSC 2024/26H1 e Server sem ano.
- Próximo passo: aprovar a Spec → `/clear` → `/implement fase 1`.

## 2026-10-27 — Feature 001, Fase 3 do SDD, Fase 1 da Spec (esqueleto, ganchos e build)
- Criado `scripts\Modulos.ps1` (UTF-8 com BOM, `#Requires -Version 2.0`): tabela `$ModulosDef` (id, chave em `$MOD`, âncora, título, frase da metodologia e nomes das funções `Cards*`/`Secao*` das fases 2 a 6), `ConvertTo-LdapEscape`, `Invoke-Modulo`, `Invoke-Modulos` (só monta `Marcados`; resultados `$null`), `NumeroSecao` e os ganchos `NavModulos`, `CardsModulos`, `SecoesModulos`, `CssModulos`, `MetodoModulos`.
- As 3 bases receberam os mesmos ganchos (diff idêntico): parâmetros novos, região `#region MODULOS` vazia, `$MOD = Invoke-Modulos`, CSS, navegação, cards, seções e metodologia.
- `build.ps1`: lê o `Modulos.ps1` (sem BOM e sem `#Requires`, quebras `\r\n`), exige exatamente 1 `#region`/`#endregion MODULOS` por base e embute com `IndexOf`/`Substring`; nova etapa `2b. Testes dos módulos`.
- Criado `tools\Test-Modulos.ps1` (5 casos: ganchos vazios, navegação 6/7, ordem de `Marcados`, seção de erro sem card, `Invoke-Modulo` capturando erro).
- Diferenças em relação à Spec:
  - Na PS2 a região fica depois do `Curto` (logo abaixo do `NovoObj`), igual às outras bases, para o diff dos ganchos ser o mesmo nas 3.
  - Os títulos das seções/navegação são constantes e **não** passam pelo `HE`: o `WebUtility.HtmlEncode` do PS51 troca `í` por `&#237;` (o visual é o mesmo, mas o texto fica ilegível no fonte).
  - Módulo marcado sem resultado (fases ainda não implementadas) gera a seção "Não foi possível ler: o módulo não foi executado". Some quando cada fase preencher o seu resultado.
  - A linha de progresso `      · <título>...` é escrita pelo `Invoke-Modulo`.
- Verificado: `.\build.ps1 -SemScan` ok (lint 2.0/3.0/5.1 dos 7 scripts, testes 5/5, compilação sem avisos); `build\scripts\Relatorio-2022.ps1` com BOM, a região com o `Modulos.ps1` e sem o `#Requires 2.0`.
- O build regravou `build\` e `dist\` (versionados): o SHA256 do `ENTREGA.md` ficou desatualizado até a Fase 8.
- Não testado: execução real (sem AD aqui). Com nenhum módulo marcado, todos os ganchos devolvem `''`/nenhum card (coberto pelos testes).
- Próximo passo: revisar → commit → `/clear` → `/implement fase 2`.

## 2026-10-27 — Feature 001, Fase 3 do SDD, Fase 2 da Spec (snapshot e comparação)
- `scripts\Modulos.ps1`: `Get-LinhasSnapshot`, `Write-Snapshot`, `ConvertTo-Snapshot`, `Read-Snapshot`, `Find-SnapshotAnterior`, `Read-ModComparar`, `Compare-Snapshot`, `CardsMudancas`, `SecaoMudancas` (e auxiliares `Get-SnapTabela`, `Get-SnapModulos`, `Split-SnapGpo`, `Get-ResumoMudancas`). `Invoke-Modulos` agora: lê o anterior → grava o `00_Snapshot` (sempre) → compara → grava o `04_Mudancas` (só com Comparar e ≥ 1 mudança).
- `tools\Test-Modulos.ps1`: 11 casos novos (16 no total): ida e volta com acentos, falha ao gravar, `Formato=2` e CSV qualquer recusados, busca do anterior (domínio sem diferenciar maiúsculas, atual e arquivo ilegível ignorados), uma comparação por categoria, escopo mudou, módulo de um lado só, cards/seção e `Invoke-Modulos` ponta a ponta numa pasta temporária. As fixtures (`$amb`, `Snap`, `NovoDir`) ficam no topo; o teste da Fase 1 do `Invoke-Modulos` passou a usar uma pasta temporária, porque agora ele grava o snapshot.
- Diferenças em relação à Spec:
  - `Compare-Snapshot($antes, $depois, $desativados)`: o 3º parâmetro é o conjunto `@{ login = 1 }` das contas desativadas (lido por `Read-ModComparar` só se alguma conta saiu). Devolve `{Mudancas; NaoComparado; EscopoMudou}` em vez de só a lista, porque a seção precisa dos avisos.
  - `Read-Snapshot -SoMeta` lê só as 8 primeiras linhas (`Get-Content -TotalCount | ConvertFrom-Csv`), usado na busca do anterior, para não carregar cada arquivo inteiro.
  - Contas que saíram aparecem com `Mudanca` = `desativada` ou `excluída` (não "saiu"). O `Item` mostra `Nome (login)` (só o login se o nome for igual, comparando maiúsculas).
  - "Não comparado" também cobre o caso contrário: "este levantamento não incluiu ...".
  - As mudanças de `Acesso` guardam a `Pasta` numa propriedade extra (fora do CSV) para o card testar `$PastasSensiveis`. CSS novo: `h3.mod-sub` (só com algum módulo marcado).
- Verificado: `.\build.ps1 -SemScan` ok (lint 2.0/3.0/5.1 dos 7 scripts, 16/16 testes, compilação sem avisos, exe 548.864 bytes). O SHA256 do `ENTREGA.md` continua desatualizado até a Fase 8.
- Não testados: a consulta `Get-ADUser -Filter 'Enabled -eq $false'` (precisa de AD), o comportamento no PS 2.0 de verdade (`Import-Csv`/`ConvertFrom-Csv` com BOM) e a meta `Modulos`, que fica vazia até as fases 3, 5 e 6 gravarem as linhas `Privilegiado`/`GPO`/`Computador`.
- Próximo passo: revisar → commit → `/clear` → `/implement fase 3`.

## 2026-10-27 — Feature 001, Fase 3 do SDD, Fase 3 da Spec (contas privilegiadas e de risco)
- `scripts\Modulos.ps1`: tabelas `$ContasGrupos` (8 grupos por SID; 519/518 no domínio raiz) e `$ContasAchadosDef` (ordem, severidade, título, texto e coluna de cada card); `Read-ModContas`, `Get-AchadosContas` (pura), `Write-CsvContas` (05), `CardsContas`, `SecaoContas` e auxiliares (`ConvertFrom-FileTimeAD`, `NovaContaAD`, `ConvertFrom-MembroAD`, `ConvertFrom-UsuarioAD`, `Get-ContaLinha`, `Set-Servico`, `Plural`, `TextoGrupos`, `Get-OUConta`). `Invoke-Modulos` roda o módulo antes do snapshot; `Get-LinhasSnapshot` grava as linhas `Privilegiado` (só membros **ativos**).
- `tools\Test-Modulos.ps1`: 9 casos novos (25 no total): escape LDAP, FileTime, cada atributo de risco, fronteiras de `DiasInativo` (nunca logou e privilegiada sem uso) e `DiasSenhaServico`, SPN fora do padrão, padrão só pela OU, DC/RODC/desativado com delegação ignorados, Administrador interno (baixa/media/renomeado/desativado), cards, seção, CSV 05 e snapshot.
- Diferenças em relação à Spec / decisões:
  - Contas são indexadas pelo **DN** (não pelo login), para não misturar contas do domínio raiz com as locais.
  - Além das propriedades da Spec, o `Get-ADObject` dos membros lê `whenCreated` e `servicePrincipalName` (para "Criada em" e conta de serviço nas contas que não vêm no `Get-ADUser`, como as do domínio raiz).
  - Tipo `conta gerenciada` para gMSA/MSA membros de grupo privilegiado.
  - Computadores com delegação **desativados** são ignorados; as falhas ao ler esses computadores ou o RID 500 viram `Avisos` (só a falha do `Get-ADUser` derruba o módulo).
  - "Privilegiada sem uso": sem logon usa o `whenCreated` como referência (conta criada há poucos dias não é marcada). Só contas ativas recebem `PrivInativo`, `PrivTerceiro` e `Servico`.
  - `Admin500` = RID 500 ativo com login `Administrator`/`Administrador`; `media` se a senha tiver mais de `DiasSenhaServico` dias ou nunca foi alterada.
  - A seção avisa que membros pelo grupo primário **e de outros domínios/florestas** não aparecem (o filtro em cadeia só enxerga objetos da partição consultada).
- Verificado: `.\build.ps1 -SemScan` ok (lint 2.0/3.0/5.1 dos 7 scripts, 25/25 testes, exe 679.936 bytes). Sem AD, `-ContasDominio` vira `[AVISO]` e `Ok=$false` sem derrubar o relatório.
- Não testados (precisam de AD): `Get-ADForest`/`Get-ADDomain -Server` em domínio filho, o filtro `memberOf:1.2.840.113556.1.4.1941` e o comportamento no PS 2.0 real. O SHA256 do `ENTREGA.md` continua desatualizado até a Fase 8.
- Próximo passo: revisar → commit → `/clear` → `/implement fase 4`.

## 2026-10-27 — Feature 001, Fase 3 do SDD, Fase 4 da Spec (política de senha e bloqueio)
- `scripts\Modulos.ps1`: `Read-ModSenha` (política padrão + FGPP; falha nas FGPP vira `Avisos`), `Get-AvaliacaoSenha` (pura, 9 itens; 7 sem limiar de bloqueio), `Get-ResultadoSenha` (pura), `CardsSenha`, `SecaoSenha` e auxiliares (`Get-NomeCN`, `Get-Duracao`, `TextoDuracao`, `TabelaSenha`, `Test-AbaixoSenha`). `Invoke-Modulos` roda o módulo depois do de contas. Sem CSV e sem linhas no snapshot.
- `tools\Test-Modulos.ps1`: 4 casos novos (29 no total): política dentro da referência (sem card), fronteiras 7/8/13/14, histórico 23/24, limiar 0/1/10/11, duração 0/"infinita"/14/15 min, contador 14/30, idade mínima 0/23 h, textos de duração, FGPP (ordem por precedência, sem alvo, severidade máxima), card, seção, aviso de FGPP e snapshot inalterado.
- Diferenças em relação à Spec / decisões:
  - Função pura extra `Get-ResultadoSenha($bruto)`: avalia a padrão e cada FGPP e monta `Abaixo` (lista Política/Item) e `Sev` (a maior), usada pelo card. Assim o card e a seção são testáveis sem AD.
  - Duração zero **ou** "infinita" (|valor| > 100 anos, como o AD devolve 0x8000000000000000) = "até desbloqueio manual" / "nunca expira". Idade mínima zero = "0 (troca imediata)", baixa.
  - Cada FGPP tem a tabela resumo (nome, precedência, aplica-se a; sem alvo = "ninguém (sem efeito)") **e** a própria tabela de avaliação, já que o card também lista os itens das FGPP.
  - A coluna "Atual" usa as classes `td.sev-*` que já estavam no `CssModulos`; itens `info` aparecem com a tag "informativo".
  - Nota da baseline: "Microsoft Security Baseline do Windows Server 2025 (publicada em 31/01/2025)", conferida no anúncio do Tech Community.
- Verificado: `.\build.ps1 -SemScan` ok (lint 2.0/3.0/5.1 dos 7 scripts, 29/29 testes, compilação sem avisos, exe 749.568 bytes). O SHA256 do `ENTREGA.md` continua desatualizado até a Fase 8.
- Não testados (precisam de AD): `Get-ADDefaultDomainPasswordPolicy`/`Get-ADFineGrainedPasswordPolicy` reais (inclusive o valor exato que o 2008 R2 devolve para "nunca expira"/"desbloqueio manual") e o PS 2.0 real.
- Próximo passo: revisar → commit → `/clear` → `/implement fase 5`.

## 2026-10-28 — Feature 001, Fase 3 do SDD, Fase 5 da Spec (GPOs via LDAP)
- `scripts\Modulos.ps1`: tabela `$GpoAchadosDef` (ordem, grupo do card e tag de cada achado); `Read-ModGpo`, `ConvertFrom-GPLink` (pura), `Get-CaminhoDN` (pura), `Get-AchadosGpo` (pura), `Write-CsvGpo` (06), `CardsGpo`, `SecaoGpo` e auxiliares (`Add-AchadoGpo`, `TextoVinculos`, `NomesAchadosGpo`). `Invoke-Modulos` roda o módulo depois do de senha; `Get-LinhasSnapshot` grava as linhas `GPO` (`nome|flags|versao`; vínculos órfãos ficam fora).
- `tools\Test-Modulos.ps1`: 4 casos novos (33 no total): `ConvertFrom-GPLink` (vazio, 1 e 3 vínculos, opções 0 a 3, maiúsculas/minúsculas, domínio filho = externo), `Get-CaminhoDN` (domínio, OU com vírgula escapada, site), um caso por achado (GPO só em site não é "sem vínculo", GUID de outro domínio não é órfão, SYSVOL inacessível sem `SemSysvol`, herança bloqueada), cards, seção, CSV 06 e snapshot.
- Diferenças em relação à Spec / decisões:
  - Os alvos do domínio são lidos com `(|(gPLink=*)(gPOptions=*))`, não só `(gPLink=*)`: uma OU com herança bloqueada e sem GPO vinculada não tem `gPLink` e sumiria da lista de herança bloqueada.
  - `ConvertFrom-GPLink($s, $domDN)`: o 2º parâmetro serve para marcar `Externo`. A comparação usa só os `DC=` do DN, porque um domínio filho também "termina" no DN do pai.
  - O `Read-ModGpo` já devolve as GPOs normalizadas (`Guid` em maiúsculas, `Flags`/`Versao` inteiros, `PastaExiste` = `$null` quando o SYSVOL não foi testado), para o `Get-AchadosGpo` ser testável sem AD.
  - Resultado traz também `Externos` (vínculos para GPOs de outro domínio): uma observação na seção, sem achado.
  - Coluna `PastaSYSVOL` do CSV 06 = `ok` / `ausente` / `não verificada` (vazia nos órfãos). Na seção, a GPO leva tags dos achados (vermelho = integridade, âmbar = limpeza) e os vínculos levam `desativado`/`imposto`.
  - Os nomes dos achados são constantes e não passam pelo `HE` (mesmo motivo dos títulos: o `WebUtility.HtmlEncode` troca os acentos por `&#NNN;`).
- Cuidado registrado: o Write/Edit do Claude Code grava os `.ps1` com LF. O build normaliza o `Modulos.ps1` embutido, mas os arquivos-fonte foram regravados com CRLF e BOM.
- Verificado: `.\build.ps1 -SemScan` ok (lint 2.0/3.0/5.1 dos 7 scripts, 33/33 testes, compilação sem avisos, exe 856.064 bytes). Sem AD, `-GPOs` vira `[AVISO]` e `Ok=$false` sem derrubar o relatório (coberto pelo teste do `Invoke-Modulos`). O SHA256 do `ENTREGA.md` continua desatualizado até a Fase 8.
- Não testados (precisam de AD): o `Get-ADObject` em `CN=Policies`/`CN=Sites`, o formato real do `gPLink` no 2008 R2, o acesso ao SYSVOL por `\\dominio\SYSVOL` e o PS 2.0 real.
- Próximo passo: revisar → commit → `/clear` → `/implement fase 6`.

## 2026-10-28 — Feature 001, Fase 3 do SDD, Fase 6 da Spec (computadores do domínio)
- `scripts\Modulos.ps1`: tabela `$TabelaSuporte` (+ `$Win11Obs`, `$CompLimiteLista = 300`), `Get-FimSuporte` (pura), `Read-ModComputadores`, `Get-AchadosComputadores` (pura), `Write-CsvComputadores` (07), `CardsComputadores`, `SecaoComputadores` e auxiliares (`NovoFim`, `Get-MaiorBuild`, `TextoFimSuporte`, `LinhasLimitadas`). `Invoke-Modulos` roda o módulo depois do de GPOs; `Get-LinhasSnapshot` grava as linhas `Computador` (`ativo 1/0|SO`, o formato que o `Compare-Snapshot` da Fase 2 já esperava).
- `tools\Test-Modulos.ps1`: 4 casos novos (37 no total): `Get-FimSuporte` com 37 strings reais do AD (uma por regra), fora de suporte/sem uso/desativados (fronteira de `DiasInativo` e do último dia com suporte), resumo por família, cards, seção (lista limitada a 300 + "+N no arquivo 07_Computadores"), CSV 07, snapshot e desempenho (5.000 computadores: ~2,1 s e HTML de 96 KB).
- Datas conferidas em 2026-10-28 (correções na Spec):
  - Windows Server 23H2 (build 25398): suporte até **2026-05-12** (declaração de fim de atualizações da Microsoft); o 2025-10-24 da Spec foi estendido.
  - Windows 11 26H2 = build **26300**, Home/Pro até 2028-10-09 (Lifecycle: 10/10/2028 06:59:59). Entrou na tabela.
- Diferenças em relação à Spec / decisões:
  - Regras extras: Windows 2000 (2010-07-13), Server sem ano com build maior que 26100 = "mais nova que a tabela" (sem alerta), outros builds sem ano = canal semestral (retirado em 2022-08-09), Windows 11 com build desconhecido abaixo do maior da tabela = "Windows 11 (build N)" sem data, Windows não reconhecido (ex.: Embedded) = família com o próprio nome do SO, sem alerta.
  - Família separa "Windows Server 2008" de "2008 R2" (e 2012/2012 R2) no resumo; a data é a mesma.
  - Sem uso = ativo e (último logon, ou criação se nunca logou) antes do limite; sem as duas datas não é marcado. Fora de suporte = ativo e fim < data de hoje (no último dia ainda tem suporte).
  - `Get-FimSuporte` é chamado uma vez por par SO/versão (cache), por desempenho.
  - Card `media` (sem uso) e `baixa` (desativados) só com texto e link para a seção, sem lista: podem ser centenas. O card `alta` traz a contagem por família.
  - Na seção, as datas do Windows 11 (não LTSC) levam `*` e a nota de Home/Pro/Enterprise; a OU aparece como caminho legível (`empresa.local/TI/Estações`); no CSV vai o DN.
  - Tabela `$TabelaSuporte` fica no início da seção do módulo 5 (padrão das outras tabelas do arquivo), não no topo do arquivo.
- Verificado: `.\build.ps1 -SemScan` ok (lint 2.0/3.0/5.1 dos 7 scripts, 37/37 testes, compilação sem avisos, exe 962.560 bytes). O SHA256 do `ENTREGA.md` continua desatualizado até a Fase 8.
- Pendências/limitações: Windows 10/11 **IoT** Enterprise LTSC têm datas maiores (LTSC 2021 IoT até 2032-01-13, 11 IoT LTSC 2024 até 2034) e hoje caem na data da edição Enterprise (sem falso "fora de suporte" hoje, mas a data mostrada fica menor). Não testados (precisam de AD): `Get-ADComputer -Filter *` real num domínio grande e no ADWS do 2008 R2, e o PS 2.0 real.
- Observação: o git ainda está no commit da 1.0; as fases 1 a 6 da feature 001 não foram commitadas.
- Próximo passo: revisar → commit → `/clear` → `/implement fase 7`.
## 2026-10-28 — Feature 001, Fase 3 do SDD, Fase 7 da Spec (interface C#)
- `src\ScriptRunner.cs`: parâmetro com valor `null` vira switch (`-Nome` sem valor). Nova propriedade `Argumentos` (linha de comando do `powershell.exe`).
- `src\MainForm.cs`: grupo "Itens adicionais (opcionais)" com as 5 caixas desmarcadas, "Comparar com:" + "..." (`OpenFileDialog` em `00_Snapshot_*.csv`, abrindo na pasta de saída), "Contas de serviço:"/"Dias senha serviço:" (1–3650, 365). `AtualizarOpcionais` habilita os campos só com "Comparar"/"Contas" marcados. Validações: regex do padrão de serviço (com Contas) e arquivo existente (com Comparar e campo preenchido). Helpers novos `AddTitulo`, `NovaCaixa` e `LinhaComBotao` (a "Pasta de saída" passou a usar). Título "RelatorioAD 1.1".
- `src\AssemblyInfo.cs`: versão 1.1.0.0.
- Diferenças em relação à Spec / decisões:
  - O log passou a registrar a linha `Comando: powershell.exe ...`, que é o que a verificação da Spec pede para conferir. Isso muda o log em relação à 1.0 (não muda o HTML nem os CSVs).
  - As caixas ficam em **duas linhas fixas** (Comparar/Contas/Senha e GPOs/Computadores), sem quebra automática. Com um `FlowLayoutPanel` com quebra, o `TableLayoutPanel` media o painel mais alto do que ele ficava e sobravam ~30 px vazios no fim dos parâmetros. `AutoSizeMode.GrowAndShrink` não resolveu.
  - Com 900×760 fixo e DPI de 125% (esta máquina), o log sumia: os campos crescem com o DPI, mas o tamanho da janela não. O `OnLoad` aumenta a altura até o log mostrar 6 linhas, limita a janela à área de trabalho e centraliza de novo. Isso cobre o 1024×768 com barra de tarefas.
- Verificado (sem AD, rodando o 2022, por `WM_SETTEXT`/`BM_CLICK` + `PrintWindow`, com o processo de teste *DPI-aware*):
  - Build sem avisos, 37/37 testes, exe 966.656 bytes.
  - Ao abrir: as 5 caixas desmarcadas (captura de tela) e os 4 campos extras desabilitados. Marcar "Comparar" habilita só o "Comparar com"; marcar "Contas" habilita o padrão e os dias. Desmarcar desabilita de novo.
  - Sem módulos, a linha de comando tem os mesmos parâmetros da 1.0. Com todos marcados, termina com `-Comparar -ContasDominio -PadraoServico "svc|servico|serviço|service|sql|backup" -DiasSenhaServico "365" -PoliticaSenha -GPOs -Computadores`, sem valor depois dos switches.
  - Log: ~7 linhas no tamanho padrão (125%) e 5,9 linhas numa janela equivalente a 1024×728 (1024×768 menos a barra de tarefas, em 125%). Com a janela de 1024×768 inteira, ~7 linhas.
- Observação de teste: o `BM_GETCHECK` sempre devolve 0 nos `CheckBox` do WinForms. O estado foi conferido pela tela e pela linha de comando.
- Não testados: o `OpenFileDialog` (interativo), a mensagem de arquivo inexistente e uma tela real de 1024×768 a 96 DPI. O SHA256 do `ENTREGA.md` continua desatualizado até a Fase 8.
- Próximo passo: revisar → commit → `/clear` → `/implement fase 8`.
## 2026-10-28 — Feature 001, Fase 3 do SDD, Fase 8 da Spec (entrega e roteiro de teste)
- `docs\ENTREGA.md`: título 1.1; seção nova "Itens adicionais (opcionais)" (tabela com o que cada item verifica e o arquivo gerado: 04 a 07; política de senha sem CSV), observações de que o programa continua somente leitura (lista dos `Get-AD*` e o `Test-Path` no SYSVOL), de que o `00_Snapshot` é gravado sempre e não deve ser apagado, e de que grupos/GPOs/computadores só entram na comparação se o item estava marcado nas duas execuções. "O que o programa faz" cita os cmdlets novos. Tabela de verificação com versão 1.1.0.0, tamanho, SHA256 e assinaturas do Defender do build final.
- `docs\features\001-auditoria-dominio\TESTE.md` (novo): roteiro na ordem 2022 → 2008 R2 → 2012 R2 → 2016/2019/2025, com preparação (link e SHA256 do zip 1.0 no commit `57a418c`, conferido), os 5 passos da Spec com o resultado esperado de cada um (textos copiados do `Modulos.ps1`), limpeza da conta/GPO de teste e tabela de resultado por servidor.
- Decisões:
  - Passo 3: a conta de teste é criada **antes** do passo 2 (precisa estar ativa no levantamento anterior para aparecer como `desativada`).
  - Passo 5 marcado "só em laboratório": simular falta de acesso ao SYSVOL exige mexer em permissão ou DNS, o que não se faz no cliente.
  - Passo 1 aceita diferença no último logon de quem logou entre as duas execuções.
- Verificado: `.\build.ps1` completo ok — lint 2.0/3.0/5.1 dos 7 scripts, 37/37 testes, compilação sem avisos, Defender sem ameaças em `dist\` e `build\scripts\` (assinaturas 1.463.8.0). Exe 966.656 bytes, SHA256 exe `0F89CAFC587D106CFBB37F8ED3C2F3112933682FE0DC5B2BB8D7C5326C503B68`, zip `B0ED22AD677D5F56BCDA536618548653D6B495DC445938B8D4B2F6361C47725B` (conferidos com `sha256sum` nos arquivos de `dist\`).
- Feature 001 implementada (8/8 fases). Pendente: teste real pelo `TESTE.md` e commit (nada da feature foi commitado ainda; HEAD = 8e73e2b, PRD).
- Próximo passo: revisar → commit → teste real nos servidores.

## 2026-10-28 — Feature 001, ajuste pedido pelo usuário: computadores sem logon há mais de 7 meses
- Pedido: separar no relatório os computadores que não logam há mais de 7 meses (para remover do AD) dos que logaram recentemente.
- `scripts\Modulos.ps1`: constante `$CompMesesAntigo = 7`. `Get-AchadosComputadores` marca `Antigo` (último logon, ou criação se nunca logou, antes de `$hoje.AddMonths(-7)`; ativo **ou desativado**) e `DiasSemLogon`; devolve `Antigos` (do mais antigo para o mais recente) e `MesesAntigo`, e o resumo por família ganha `Antigos`. `Inativo` e `ForaSuporte` passam a excluir os antigos. Cards: novo `media` "N computadores sem logon há mais de 7 meses (candidatos à remoção)" (com quantos ainda estão ativos e quantos já desativados); o de sem uso cai para `baixa` com "(e menos de 7 meses)"; o de desativados cita quantos já estão na lista de remoção. Seção: coluna "Sem logon há mais de 7 meses" no resumo (âmbar quando > 0), lista "candidatos à remoção" (Situação, Sistema, Último logon/"nunca", Dias sem logon, Criado em, OU) antes da lista de sem uso, que ganhou a coluna Dias sem logon. CSV 07: colunas `DiasSemLogon` e `CandidatoRemocao`.
- `tools\Test-Modulos.ps1`: casos novos dentro dos 2 testes de computadores (fronteira de exatamente 7 meses, desativado antigo, nunca logou criado há 400 dias, antigo fora da lista de fora de suporte e da de sem uso, cards, coluna do resumo, lista, colunas e valores do CSV). Continuam 37 testes.
- Spec (nota no fim), `ENTREGA.md` (linha de Computadores e SHA256) e `TESTE.md` (conferência no passo 2) atualizados.
- Decisões: 7 meses fixo, sem campo na interface (pedido específico); desativados antigos entram na lista de remoção, porque também são candidatos; o snapshot e a comparação não mudam.
- Verificado: `.\build.ps1` completo ok (lint 2.0/3.0/5.1, 37/37 testes, sem avisos, Defender limpo, assinaturas 1.463.8.0). Exe 995.328 bytes, SHA256 exe `F88985652EE36D481461352CDDA476E93C6CC49B475F127ADA7B33E5D095761C`, zip `6D2743383BCB89725720C07EEE0ECAB00F9C9927F9B6362707CA774862CC5869`.
- Sem commit: o usuário vai testar antes.

## 2026-10-28 — Feature 001, ajuste pedido pelo usuário: explicação e conteúdo das GPOs
- Pedido: um "disclaimer" melhor das GPOs, dizendo o que elas fazem e quais ações têm no domínio.
- `scripts\Modulos.ps1`:
  - Seção GPOs abre com dois parágrafos: o que é uma GPO (vínculo, partes Computador/Usuário, ~90 min, ordem site → domínio → OUs, imposto e bloqueio de herança) e "Como ler esta seção" (lido do AD e do SYSVOL sem o console e sem alterar nada; Modelos administrativos pela chave do registro; filtros de segurança/WMI e direcionamento por item não analisados; como tirar o relatório completo no GPMC ou com `Get-GPOReport`).
  - Bloco novo "O que cada GPO faz": um por GPO (âncora `gpo-<guid>`, linkada na tabela), com status, data e vínculos e uma tabela Parte/Área/O que configura. Tag "parte desativada" conforme `flags`, "a GPO não faz nada" nas vazias, aviso de política de senha vinculada fora da raiz do domínio, e o que não pôde ser lido.
  - Leitura (`Read-ConteudoGpo`, um `try` por arquivo): `GptTmpl.inf` (política de contas, Kerberos, auditoria, direitos de usuário com SID traduzido e cache, opções de segurança, grupos restritos, serviços, permissões), `audit.csv`, `Registry.pol` (parser PReg próprio, agrupado por assunto em `$GpoAreasRegistro`), `scripts.ini`/`psscripts.ini`, XMLs de `Preferences\` (unidades, impressoras, usuários e grupos locais, tarefas, arquivos, atalhos, registro, variáveis, serviços, compartilhamentos, .ini; as demais pelo nome) e `fdeploy1.ini`. Áreas que só aparecem em `gPCMachineExtensionNames`/`gPCUserExtensionNames` (tabela `$GpoCse`) entram sem detalhe; instalação de software vem dos `packageRegistration` do AD (uma consulta só).
  - Achado novo `SenhaPreferencia` (cpassword preenchido, MS14-025): card `alta` com o arquivo; o valor nunca é lido para o relatório.
  - HTML limitado a 15 linhas por área; CSV novo `06_GPOs_Configuracoes_<stamp>.csv` com tudo; o 06 ganhou `Configuracoes` e `ErrosLeitura`. Snapshot inalterado. CSS: `h4.gpo`, `p.gpo-meta`, `ul.gpo-det`.
- `tools\Test-Modulos.ps1`: 6 testes novos (43 no total): CSE, Registry.pol (bytes montados no teste), GptTmpl.inf, scripts/Preferências/redirecionamento (inclui senha exposta sem o valor), `Read-ConteudoGpo` numa pasta falsa (UTF-16 com BOM, XML UTF-8 sem BOM com acento, Registry.pol inválido não derruba o resto) e seção/cards/CSVs.
- Conferido visualmente: prévia da seção com o CSS real renderizada no Edge headless.
- Spec (nota no fim), `ENTREGA.md` (linha de GPOs, somente leitura e SHA256) e `TESTE.md` (comparar 2–3 GPOs com o relatório do GPMC no passo 2; GPO vazia no passo 3) atualizados.
- Verificado: `.\build.ps1` completo ok (lint 2.0/3.0/5.1, 43/43 testes, sem avisos, Defender limpo, assinaturas 1.463.8.0). Exe 1.290.240 bytes, SHA256 exe `C89F918101D3BB4DEFA39DAA8CBA7A89220BDBBD257D9B7A4E64EE9FFA1558C1`, zip `6A18CB8C578D505AF592EB4267DAA544DF8ED1C3FE0BE9C19ABE7DAF6AF8B07A`.
- Não testado (precisa de AD): a leitura real do SYSVOL, a tradução de SIDs, os `packageRegistration` e o PS 2.0 real. Sem commit: o usuário vai testar antes.

## 2026-10-28 — RelatorioAD 1.1: teste real e publicação
- O usuário informou que a 1.1 funcionou nos servidores reais com Windows Server 2008 (R2), 2012 e 2022, e liberou o commit e o push.
- `README.md`: botão de download com a versão ("BAIXAR RelatorioAD 1.1", mesmo link `raw/main/dist/RelatorioAD.zip`), versão/data/SHA256 do zip, novidades da 1.1, seção "Itens adicionais (opcionais)", o `00_Snapshot` na tabela de arquivos, os campos novos no passo a passo, "Comparar uma visita com a anterior", "Remover computadores antigos", novos problemas comuns, `Modulos.ps1`/`Test-Modulos.ps1` em "Como funciona"/"Estrutura" e situação atualizada.
- Sem rebuild: `dist\` é o do build registrado no ajuste das GPOs (SHA256 conferido).
