# Spec: Auditoria do domínio (módulos opcionais)

> Baseada em `docs/features/001-auditoria-dominio/PRD.md` · Data: 2026-10-27 · Projeto: RelatorioAD 1.0 → 1.1

## Overview
Cinco módulos opcionais (comparação com o levantamento anterior, contas privilegiadas, política de senha, GPOs e computadores) escritos **uma vez** em `scripts/Modulos.ps1`, compatível com PS 2.0. O `build.ps1` embute esse arquivo nas 3 bases. A interface ganha 5 caixas desmarcadas e 3 campos extras. Com tudo desmarcado, o HTML e os CSVs 01/02/03/99 ficam idênticos aos da 1.0; a única saída nova é o `00_Snapshot_<stamp>.csv`.

## Decisões de arquitetura (confirmando o PRD)

1. **Um arquivo de módulos, embutido pelo build.** Cada base ganha uma região vazia:
   ```powershell
   #region MODULOS
   #endregion MODULOS
   ```
   Ela fica logo depois das funções utilitárias do topo: após `Curto` (PS51/PS3) e após `NovoObj` (PS2). O build troca o miolo da região pelo conteúdo de `scripts/Modulos.ps1`, sem o BOM e sem a linha `#Requires`. O script gerado continua sendo **um arquivo só**.
   - A troca é feita com `IndexOf`/`Substring`, **não** com `[regex]::Replace`, porque o `$` do código seria lido como referência de grupo.
   - As bases em `scripts\` deixam de rodar sozinhas com módulos marcados. O script executável é o de `build\scripts\`, como já acontece hoje com o `$VersaoSO`.
2. **`Modulos.ps1` só define funções e tabelas constantes**, sem nenhum efeito colateral no nível do arquivo. Assim `tools/Test-Modulos.ps1` consegue fazer dot-source dele numa máquina sem AD.
3. **Cada módulo tem 3 camadas:**
   - `Read-Mod*`: lê o AD/SYSVOL, recebe tudo por parâmetro e devolve dados brutos.
   - `Get-Achados*`: **pura**, recebe os dados brutos mais os parâmetros (`$hoje`, `$DiasInativo` etc.) e devolve o resultado. É a camada testada pelo `Test-Modulos.ps1`.
   - `Cards*` / `Secao*`: transformam o resultado em HTML (strings) e usam o `HE`, o `DT` e o `Card` da base.
4. **Ganchos nas bases devolvem `''` ou `@()` quando nenhum módulo está marcado**: navegação, cards, seções, CSS e metodologia. É isso que garante o RF1.
5. **CSS novo só entra no HTML quando algum módulo rodou** (`CssModulos`), para não alterar o HTML padrão.
6. **Switches vindos do .exe**: o `ScriptRunner` passa a aceitar um parâmetro com valor `null` e grava só `-Nome`. Não usar `-Nome:$true`, que falha com `-File` no PS 2–4.
7. **Membros de grupos privilegiados** via `LDAP_MATCHING_RULE_IN_CHAIN` (`memberOf:1.2.840.113556.1.4.1941:=<DN>`) com `Get-ADObject`, **não** via `Get-ADGroupMember -Recursive`. Este último falha com membros de outros domínios ou de confiança (FSP) e com grupos grandes no ADWS do 2008 R2.

## Contratos comuns (`scripts/Modulos.ps1`)

```powershell
# Resultado de cada módulo (objeto via New-Object PSObject -Property, nunca [pscustomobject])
#   Ok    [bool]    – $false se o módulo falhou por inteiro
#   Erro  [string]  – mensagem mostrada na seção ("Não foi possível ler: ...")
#   Avisos[string[]]– falhas parciais (ex.: SYSVOL ilegível, domínio raiz inacessível)
#   ...             – dados específicos do módulo

# $MOD (hashtable montada por Invoke-Modulos)
#   Marcados = @('Comparar','ContasDominio','PoliticaSenha','GPOs','Computadores')  # só os marcados, nesta ordem
#   Contas / Senha / Gpo / Comp / Mudancas = resultado do módulo, ou $null se desmarcado
#   Snapshot = caminho do 00_Snapshot gravado (ou $null se falhou)
```

- `Invoke-Modulo([string]$titulo, [scriptblock]$bloco)` executa o bloco em `try/catch`. Em caso de erro, grava no log `Write-Host "[AVISO] Módulo ${titulo}: $msg" -ForegroundColor Yellow` e devolve `Ok=$false`, `Erro=$msg` (RF5). Por regra das bases, `$ErrorActionPreference = 'Stop'`.
- Cada módulo escreve no log uma linha de progresso `"      · <título>..."`, abaixo do `[5/5]`. A linha `Concluído. ...` final **não muda**, porque o `ScriptRunner`/`MainForm` faz parse dela.
- Regras PS 2.0 em todo o arquivo: `foreach` só sobre `@(...)`; `$x = @(if ...)` com ramos sem `@()`; `.Count` só sobre `@(...)` (no PS 2.0 um escalar não tem `.Count`); nada de `[pscustomobject]`, `[ordered]`, `HashSet`, `-in`/`-notin` ou `Where-Object` sem bloco. Conjuntos são `@{}`. A ordem das colunas dos CSVs é fixada com `Select-Object <lista>`.
- Escapar valores em filtro LDAP com `ConvertTo-LdapEscape([string]$s)`: troca `\ * ( ) NUL` por `\5c \2a \28 \29 \00`.

---

## Fase 1: Esqueleto, ganchos nas bases e build

### `scripts/Modulos.ps1` (CRIAR, UTF-8 com BOM, `#Requires -Version 2.0`)
- `Invoke-Modulo($titulo, $bloco)` → objeto resultado (ver contratos).
- `Invoke-Modulos` → `$MOD`. Lê variáveis de escopo do script (`$Comparar`, `$ContasDominio`, `$PoliticaSenha`, `$GPOs`, `$Computadores`, `$dom`, `$U`, `$ADM`, `$alvos`, `$Calc`, `$Saida`, `$stamp`, `$hoje`...). Ordem de execução:
  1. os módulos 2 a 5 marcados (coleta e análise);
  2. snapshot (Fase 2);
  3. comparação, se marcada (Fase 2).
  Nesta fase só monta `Marcados` e deixa os resultados `$null`.
- Ganchos de HTML. Todos devolvem `''` (ou `@()` no caso dos cards) quando `$MOD.Marcados` está vazio:
  - `NavModulos($MOD)` → `<a href="#mod-...">N. Título</a>` para cada módulo marcado, numerados a partir de 6 na ordem Comparar, Contas, Senha, GPOs, Computadores.
  - `CardsModulos($MOD)` → `string[]`, primeiro o card da comparação e depois os dos módulos 2 a 5.
  - `SecoesModulos($MOD)` → as `<section id="mod-...">`, com a mesma numeração da navegação. Para um módulo com `Ok=$false`, a seção traz só `<p class="lead">Não foi possível ler: <erro></p>`.
  - `CssModulos($MOD)` → regras CSS extras (`.t-ok`, `.t-bad`, `table.mini td.sev-*`).
  - `MetodoModulos($MOD)` → frases extras para a metodologia (" Também foram verificados: ...").
- Número da seção: `NumeroSecao($MOD, $id)` → 6 + posição em `Marcados`.

### `scripts/Relatorio-PS51.ps1`, `-PS3.ps1`, `-PS2.ps1` (MODIFICAR, as 3 de forma idêntica)
- `param()`, depois de `$PastasSensiveis`:
  ```powershell
  # Itens adicionais (opcionais) — ver scripts\Modulos.ps1
  [switch]$Comparar,
  [string]$CompararCom,
  [switch]$ContasDominio,
  [switch]$PoliticaSenha,
  [switch]$GPOs,
  [switch]$Computadores,
  # Regex aplicada ao login e à OU para identificar contas de serviço
  [string]$PadraoServico = 'svc|servico|serviço|service|sql|backup',
  [int]$DiasSenhaServico = 365,
  ```
- Região `#region MODULOS` / `#endregion MODULOS`, vazia, após as funções utilitárias do topo.
- No `[5/5]`, logo depois de `$nunca = ...`: `$MOD = Invoke-Modulos`.
- `<style>$css</style>` → `<style>$css$(CssModulos $MOD)</style>`.
- Navegação: ao fim de `<a href="#usuarios">5. Usuários ativos</a>`, acrescentar `$(NavModulos $MOD)`.
- "Pontos de atenção", imediatamente antes de `if (-not $nCards)`:
  ```powershell
  foreach ($c in @(CardsModulos $MOD)) { A $c; $nCards++ }
  ```
- Depois de `A '</tbody></table></section>'` (seção 5) e antes da metodologia: `A (SecoesModulos $MOD)`.
- Metodologia: `...ver arquivo 99_Erros_Acesso." })$(MetodoModulos $MOD)</p>`.

### `build.ps1` (MODIFICAR)
- Etapa 1, antes do laço das versões: ler `scripts\Modulos.ps1` (UTF-8), tirar a linha `#Requires` e normalizar as quebras para `\r\n`.
- No laço, depois da troca do `$VersaoSO`: verificar que existe **exatamente 1** `#region MODULOS` e 1 `#endregion MODULOS` (senão `throw`) e inserir o conteúdo entre os dois com `IndexOf`/`Substring`. A linha do log passa a dizer `(VersaoSO = '...', módulos embutidos)`.
- Nova etapa `2b. Testes dos módulos`: `& tools\Test-Modulos.ps1` no mesmo processo; `$LASTEXITCODE -ne 0` → `throw`.

### `tools/Test-Modulos.ps1` (CRIAR, UTF-8 com BOM)
- Faz dot-source de `scripts\Modulos.ps1` e define stubs mínimos de `HE`, `DT` e `Card`, iguais aos da base PS51.
- Mini framework: `Teste([string]$nome, [scriptblock]$b)` e `Igual($esperado, $obtido, $msg)`. Imprime `OK`/`FALHOU` por caso e termina com `exit 1` se houver falha.
- Casos da Fase 1: com `$MOD = @{ Marcados = @() }`, `NavModulos`, `SecoesModulos`, `CssModulos` e `MetodoModulos` devolvem `''`, e `@(CardsModulos $MOD).Count` é 0. Com `Marcados = @('ContasDominio','GPOs')`, a navegação traz `6. Contas privilegiadas` e `7. GPOs`.
- As fases seguintes acrescentam os seus casos neste arquivo.

### Verificação Fase 1
- [x] `.\build.ps1 -SemScan` passa: os 7 scripts são gerados com os módulos embutidos, o lint aprova (alvos 2.0/3.0/5.1) e o `Test-Modulos.ps1` passa.
- [x] `git diff` das 3 bases mostra os mesmos ganchos, sem outras mudanças.
- [x] Em `build\scripts\Relatorio-2022.ps1`, a região MODULOS contém o `Modulos.ps1` e o arquivo começa com BOM.
- [x] Revisão: em todos os ganchos, `$MOD.Marcados` vazio → nenhuma string nova no HTML.

---

## Fase 2: Snapshot e "Comparar com o levantamento anterior" (módulo 1)

### `scripts/Modulos.ps1` (MODIFICAR)
**Formato do snapshot** (`00_Snapshot_<stamp>.csv`, `Export-Csv` com o `$csvOpts` da base, colunas `Tipo;Chave;Valor`):

| Tipo | Chave | Valor | Quando |
|---|---|---|---|
| `_meta` | `Formato` | `1` | sempre |
| `_meta` | `Dominio` | `$dom.DNSRoot` | sempre |
| `_meta` | `Data` | `yyyy-MM-dd HH:mm` (cultura invariante) | sempre |
| `_meta` | `Servidor` | `$env:COMPUTERNAME` | sempre |
| `_meta` | `Escopo` | `compartilhamentos` ou os `$Caminhos` unidos por `\|` | sempre |
| `_meta` | `Modulos` | módulos 2, 4 e 5 que rodaram com `Ok`, separados por vírgula | sempre |
| `Usuario` | login | nome | sempre |
| `Admin` | login | | sempre |
| `Pasta` | nome do compartilhamento | caminho | sempre |
| `Acesso` | `pasta\|login` | nível efetivo (1, 2 ou 3) | sempre, a partir de `$Calc[pasta].EF` |
| `Privilegiado` | `grupo\|login` | | com ContasDominio |
| `GPO` | GUID | `nome\|flags\|versao` | com GPOs |
| `Computador` | nome | `ativo\|SO` | com Computadores |

- `Get-LinhasSnapshot($MOD)` → linhas `Tipo/Chave/Valor`, montadas com as variáveis da base.
- `Write-Snapshot($linhas, $arquivo)`: `Export-Csv`; se falhar → `[AVISO]` no log e `$MOD.Snapshot = $null` (não derruba).
- `Read-Snapshot([string]$arquivo)` → `@{ Meta = @{...}; Dados = @{ Tipo = @{ Chave = Valor } } }`, lido com `Import-Csv -Delimiter ';'` (o PS 2.0 detecta o BOM UTF-8). Formato diferente de `1` ou sem `Dominio` → `throw "arquivo não é um levantamento do RelatorioAD"`.
- `Find-SnapshotAnterior([string]$pasta, [string]$dominio, [string]$atual)` → caminho ou `$null`. Percorre `Get-ChildItem -LiteralPath $pasta -Filter '00_Snapshot_*.csv'`, ignora `$atual` e os arquivos ilegíveis, filtra por `Meta.Dominio` igual (sem diferenciar maiúsculas) e fica com a maior `Meta.Data`.
- **Ordem em `Invoke-Modulos`**: (1) módulos 2 a 5; (2) se `Comparar`: localizar o anterior (`$CompararCom` ou `Find-SnapshotAnterior`) **antes** de gravar; (3) gravar o snapshot atual; (4) comparar.
- `CompararCom` de outro domínio → a seção diz "O arquivo escolhido é do domínio X; a comparação só é feita entre levantamentos do mesmo domínio" e não compara.
- `Compare-Snapshot($antes, $depois, $logins)` → lista `{Categoria; Item; Mudanca; Antes; Depois}`:
  - **Contas**: `nova` / `saiu`. Para as que saíram, `Read-ModComparar` busca uma vez `Get-ADUser -Filter 'Enabled -eq $false'` (só `SamAccountName`) e classifica em `desativada` ou `excluída`. Para as novas, o `Depois` mostra "acesso a N pastas".
  - **Administradores do servidor**: `entrou` / `saiu`.
  - **Pastas**: `nova` / `removida`. Se o `Escopo` mudou entre as execuções, a seção avisa "o escopo de pastas mudou" e não lista pastas novas nem removidas.
  - **Acesso**: só para pastas e logins presentes nos dois levantamentos: `ganhou` (0 → n), `perdeu` (n → 0) e `mudou` (n → m), com o texto do nível (Leitura/Modificar/Controle total).
  - **Grupos privilegiados**: `entrou` / `saiu` por grupo, só se `ContasDominio` estiver em `Meta.Modulos` dos **dois** lados.
  - **GPOs**: `nova` / `removida` / `alterada` (a versão mudou) / `status` (os flags mudaram), só se `GPOs` estiver dos dois lados.
  - **Computadores**: `novo` / `removido` / `desativado` / `reativado` / `SO mudou`, só se `Computadores` estiver dos dois lados.
  - Categorias puladas porque o levantamento anterior não tinha o módulo: listadas na seção ("não comparado: o levantamento anterior não incluiu GPOs").
- CSV `04_Mudancas_<stamp>.csv`: `Categoria;Item;Mudanca;Antes;Depois`. Só é gravado com Comparar marcado e com pelo menos 1 mudança.
- `SecaoMudancas($r, $num)`: título `Mudanças desde dd/mm/aaaa` (data do anterior). Mostra uma tabela-resumo por categoria e uma tabela por categoria em `.scroll`. Sem anterior: `<p class="lead">Nenhum levantamento anterior encontrado nesta pasta. Este levantamento foi salvo e servirá de base para a próxima comparação.</p>`. Sem mudanças: "Nenhuma mudança encontrada."
- `CardsMudancas($r)`: um card de resumo, com link `#mod-mudancas`. `media` se entrou alguém em Administradores do servidor ou em grupo privilegiado, ou se alguém ganhou acesso a uma pasta que casa com `$PastasSensiveis`; senão `baixa`. Sem anterior ou sem mudanças, nenhum card.

### `tools/Test-Modulos.ps1` (MODIFICAR)
- Ida e volta: `Get-LinhasSnapshot` com dados falsos → `Write-Snapshot` num diretório temporário → `Read-Snapshot` devolve os mesmos valores, com acentos preservados.
- `Find-SnapshotAnterior`: 3 arquivos (2 do domínio A com datas diferentes e 1 do domínio B) → devolve o mais recente do A e ignora o `$atual`.
- `Compare-Snapshot`: um caso por categoria (incluindo `mudou` de nível, pastas e logins fora da interseção ignorados, e categoria pulada quando falta o módulo de um lado).
- Arquivo com `Formato=2` → erro tratado.

### Verificação Fase 2
- [x] Os testes da Fase 2 passam; build e lint ok.
- [x] Revisão: sem `Comparar`, o único arquivo novo na pasta é o `00_Snapshot_*.csv` (o 04 não é gravado).

---

## Fase 3: Contas privilegiadas e de risco no domínio (módulo 2)

### `scripts/Modulos.ps1` (MODIFICAR)
- `Read-ModContas($dom)` → dados brutos:
  - **Grupos** (por SID): `$dom.DomainSID` + `-512` (Domain Admins); `S-1-5-32-544`, `-548`, `-549`, `-550` e `-551` (Builtin); do **domínio raiz**: `-519` (Enterprise Admins) e `-518` (Schema Admins). O domínio raiz vem de `(Get-ADForest).RootDomain`; se for diferente de `$dom.DNSRoot`, usar `Get-ADDomain -Server <raiz>` e consultar com `-Server <raiz>`. Se a raiz falhar, gravar em `Avisos` ("Enterprise/Schema Admins não puderam ser lidos no domínio raiz <x>: ...") e seguir com os outros grupos.
  - **Membros**: `Get-ADGroup -Identity <sid> [-Server]` → DN; depois `Get-ADObject [-Server] -LDAPFilter "(&(objectClass=user)(memberOf:1.2.840.113556.1.4.1941:=$(ConvertTo-LdapEscape $dn)))" -Properties sAMAccountName, displayName, userAccountControl, lastLogonTimestamp, pwdLastSet, objectClass, distinguishedName`. Inclui contas de computador (objectClass `computer` herda de `user`). Ativa = `-not ($uac -band 2)`. As datas vêm de `[datetime]::FromFileTime` (0 ou vazio = nunca).
  - **Usuários ativos com atributos de risco**: `Get-ADUser -Filter 'Enabled -eq $true' -Properties PasswordNotRequired, TrustedForDelegation, DoesNotRequirePreAuth, AllowReversiblePasswordEncryption, ServicePrincipalNames, PasswordLastSet, LastLogonDate, whenCreated, DisplayName`.
  - **Computadores com delegação irrestrita**: `Get-ADComputer -Filter 'TrustedForDelegation -eq $true' -Properties PrimaryGroupID, LastLogonDate, OperatingSystem`, excluindo `PrimaryGroupID` 516 (DC) e 521 (RODC).
  - **Administrador interno**: `Get-ADUser -Identity "$domSid-500" -Properties PasswordLastSet, LastLogonDate`.
- `Get-AchadosContas($bruto, $hoje, $DiasInativo, $PadraoServico, $DiasSenhaServico, $PadraoTerceiros)` → `Contas` (uma linha por conta privilegiada **ou** com achado) e `Achados` por tipo:
  - `SenhaNaoExigida`, `DelegacaoIrrestrita` (usuários + computadores não-DC), `SemPreAuth`, `Reversivel`;
  - `NuncaLogou`: `LastLogonDate` vazio e `whenCreated` < `$hoje - $DiasInativo`;
  - `Servico`: tem SPN **ou** o login/DN casa com `$PadraoServico`. Alerta se `PasswordLastSet` estiver vazio ou for mais antigo que `$DiasSenhaServico` dias;
  - `PrivInativo`: privilegiada ativa sem logon há mais de `$DiasInativo` dias;
  - `PrivTerceiro`: privilegiada cuja OU casa com `$PadraoTerceiros`;
  - `Admin500`: ativo + nome `Administrator`/`Administrador` + idade da senha.
- `CardsContas($r)`. Um card por tipo com ocorrências, cada um com uma frase simples explicando o risco (RF6) e uma `table.mini`:

  | Achado | Sev. | Texto-base |
  |---|---|---|
  | Senha não exigida | alta | "a conta pode ficar com senha em branco" |
  | Delegação irrestrita | alta | "quem controla a conta pode se passar por qualquer usuário que se autentique nela" |
  | Pré-autenticação desativada | alta | "permite tentar descobrir a senha offline (AS-REP roasting)" |
  | Criptografia reversível | alta | "a senha fica armazenada de forma recuperável" |
  | Terceiro com privilégio no domínio | alta | — |
  | Privilegiada sem uso | media | — |
  | Conta de serviço com senha antiga | media | — |
  | Administrador interno | baixa (media se senha > `DiasSenhaServico`) | — |
  | Nunca fez logon | baixa | — |

- `SecaoContas($r, $num)`: tabela grupo × quantidade de membros ativos; tabela de contas privilegiadas (Conta, Login, Grupos, Ativa, Último logon, Senha alterada, tags `terceiro` / `serviço` / `sem uso`); `Avisos` em `<p class="lead">`. Nota fixa: "membros pelo grupo primário não aparecem".
- CSV `05_Contas_Privilegiadas_<stamp>.csv`: `Login;Nome;Tipo;Ativa;GruposPrivilegiados;ContaServico;SPN;Achados;UltimoLogon;SenhaAlteradaEm;CriadaEm;OU`.
- Linhas do snapshot `Privilegiado` (Fase 2).

### `tools/Test-Modulos.ps1` (MODIFICAR)
- Fixtures de objetos falsos com `New-Object PSObject -Property`. Um caso por achado e as fronteiras: exatamente `DiasInativo` dias; SPN sem casar com o padrão; padrão casando só pela OU; DC com delegação ignorado.
- `ConvertTo-LdapEscape 'CN=A (B)*'` → `CN=A \28B\29\2a`.

### Verificação Fase 3
- [x] Testes, build e lint ok.
- [x] Revisão: nenhuma chamada a `Get-ADGroupMember`; todos os grupos por SID. (As 3 chamadas que existem são do `Resolve-Principal` da 1.0, nas bases, fora do módulo.)

---

## Fase 4: Política de senha e bloqueio (módulo 3)

### `scripts/Modulos.ps1` (MODIFICAR)
- `Read-ModSenha` → `Padrao = Get-ADDefaultDomainPasswordPolicy`; `Refinadas = @(Get-ADFineGrainedPasswordPolicy -Filter * -Properties AppliesTo, Precedence)`. Se a leitura das FGPP falhar, gravar em `Avisos` e seguir só com a política padrão. `AppliesTo` (DNs) → nome curto via `CN=`.
- `Get-AvaliacaoSenha($pol)` (pura) → linhas `{Item; Atual; Referencia; Sev; Explicacao}`, com `Sev` em `ok|baixa|media|alta|info`:

  | Item | Propriedade | Regra |
  |---|---|---|
  | Tamanho mínimo | `MinPasswordLength` | < 8 alta · 8–13 media · ≥ 14 ok |
  | Complexidade | `ComplexityEnabled` | desligada alta |
  | Histórico | `PasswordHistoryCount` | < 24 baixa |
  | Criptografia reversível | `ReversibleEncryptionEnabled` | ligada alta |
  | Bloqueio por tentativas | `LockoutThreshold` | 0 alta · > 10 baixa · 1–10 ok (texto: "a baseline do Windows Server 2025 usa 3") |
  | Duração do bloqueio | `LockoutDuration` | só se o limiar for > 0: `TimeSpan.Zero` = "até desbloqueio manual" ok · < 15 min baixa |
  | Zerar contador após | `LockoutObservationWindow` | só se o limiar for > 0: < 15 min baixa |
  | Idade mínima | `MinPasswordAge` | < 1 dia baixa |
  | Validade máxima | `MaxPasswordAge` | info (`TimeSpan.Zero` = "nunca expira"; texto: a Microsoft deixou de recomendar expiração periódica) |

- `CardsSenha($r)`: um card, com a severidade mais alta entre os itens abaixo da referência (política padrão + FGPP), listando os itens. Nenhum item abaixo → sem card.
- `SecaoSenha($r, $num)`: tabela da política padrão com a coluna Situação (`<span class="tag t-ok">ok</span>` ou `<span class="tag t-bad">abaixo</span>`), explicação em uma linha por item, depois uma tabela por FGPP (nome, precedência, aplica-se a) e a nota da baseline com a data de referência. **Sem CSV** (RF4).

### `tools/Test-Modulos.ps1` (MODIFICAR)
- `Get-AvaliacaoSenha` com objetos falsos (`TimeSpan`s): cada fronteira da tabela (7/8/13/14; limiar 0/1/10/11; duração 0/14/15 min; histórico 23/24).

### Verificação Fase 4
- [x] Testes, build e lint ok.

---

## Fase 5: GPOs via LDAP (módulo 4)

### `scripts/Modulos.ps1` (MODIFICAR)
- `Read-ModGpo($dom)`:
  - GPOs: `Get-ADObject -SearchBase "CN=Policies,CN=System,$($dom.DistinguishedName)" -LDAPFilter '(objectClass=groupPolicyContainer)' -Properties displayName, flags, versionNumber, gPCFileSysPath, whenCreated, whenChanged`. O GUID é o `Name` (`{...}`).
  - Vínculos: `Get-ADObject -LDAPFilter '(gPLink=*)' -Properties gPLink, gPOptions` (domínio + OUs) **e** `Get-ADObject -SearchBase "CN=Sites,$((Get-ADRootDSE).configurationNamingContext)" -LDAPFilter '(gPLink=*)' -Properties gPLink, gPOptions` (sites). Sem os sites, uma GPO vinculada só a site seria apontada como "sem vínculo". Se a leitura dos sites falhar, gravar em `Avisos`.
  - SYSVOL: primeiro testar `Test-Path -LiteralPath "\\$($dom.DNSRoot)\SYSVOL\$($dom.DNSRoot)\Policies"`. Se der falso, `SysvolOk = $false`, gravar em `Avisos` e **não** testar GPO por GPO. Se der verdadeiro, `Test-Path -LiteralPath $gpo.gPCFileSysPath` para cada GPO.
- `ConvertFrom-GPLink([string]$s)` (pura) → `{Guid; Dn; Opcao; Desativado = ($op -band 1); Imposto = ($op -band 2)}` por `[LDAP://cn={GUID},<dn>;<op>]`, com regex sem diferenciar maiúsculas. GUIDs de outro domínio (o DN não termina no DN do domínio) ficam marcados como `Externo`.
- `Get-AchadosGpo($gpos, $alvosVinculo, $domDN, $sysvolOk)` (pura) → por GPO: status (0 ativa, 1 usuário desativado, 2 computador desativado, 3 toda desativada), lista de onde está vinculada (DN → caminho legível `dominio/OU/Sub`), vínculos desativados e imposições; mais o bloqueio de herança (`gPOptions = 1`) por OU. Achados:
  - `SemVinculo` (nenhum vínculo ativo nem desativado) — baixa;
  - `Desativada` (flags 3) — baixa;
  - `VinculoDesativado` — baixa;
  - `Vazia` (`versionNumber` 0) — baixa;
  - `Orfao` (GUID no `gPLink` sem GPO correspondente e não `Externo`) — media;
  - `SemSysvol` (só se `sysvolOk`) — media.
- `CardsGpo($r)`: um card `media` "GPOs com problema de integridade" (órfãos + sem SYSVOL) e um card `baixa` "GPOs para limpeza" (os outros 4 tipos, agrupados por tipo). Cada card só aparece se tiver ocorrências.
- `SecaoGpo($r, $num)`: tabela GPO × status × versão × alterada em × vinculada em (com tags `desativado` / `imposto`), lista de OUs com herança bloqueada e `Avisos`.
- CSV `06_GPOs_<stamp>.csv`: `GPO;GUID;Status;Versao;CriadaEm;AlteradaEm;VinculadaEm;VinculosDesativados;PastaSYSVOL;Achados`. Os vínculos órfãos entram como linhas com `GPO = (inexistente)`.
- Linhas do snapshot `GPO` (Fase 2).

### `tools/Test-Modulos.ps1` (MODIFICAR)
- `ConvertFrom-GPLink` com 0, 1 e 3 vínculos, opções 0/1/2/3, maiúsculas/minúsculas e string vazia.
- `Get-AchadosGpo`: um caso por achado, GPO ligada só a site (não é `SemVinculo`), GUID externo (não é `Orfao`) e `sysvolOk = $false` (sem `SemSysvol`).

### Verificação Fase 5
- [x] Testes, build e lint ok.
- [x] Revisão: nenhum `Import-Module GroupPolicy` nem `Get-GPO`/`Get-GPOReport`.

---

## Fase 6: Computadores do domínio (módulo 5)

### `scripts/Modulos.ps1` (MODIFICAR)
- `Read-ModComputadores` → `Get-ADComputer -Filter * -Properties OperatingSystem, OperatingSystemVersion, LastLogonDate, whenCreated`. O `Enabled` já vem por padrão.
- `$TabelaSuporte`: tabela constante no topo do arquivo, sem efeito colateral. **Datas conferidas no Microsoft Lifecycle em 2026-10-27**: o fim é o último dia com suporte, e o Lifecycle mostra o dia seguinte às 06:59:59 PT.

  | Regra (nome em `OperatingSystem`, build em `OperatingSystemVersion` `x.y (build)`) | Fim |
  |---|---|
  | `Windows XP` / `Vista` / `Windows 7` / `Windows 8.1` / `Windows 8` (testar 8.1 antes de 8) | 2014-04-08 / 2017-04-11 / 2020-01-14 / 2023-01-10 / 2016-01-12 |
  | `Windows 10` + `LTSB`/`LTSC`: build 10240 / 14393 / 17763 / 19044 | 2025-10-14 / 2026-10-13 / 2029-01-09 / 2027-01-12 |
  | `Windows 10`, qualquer outro | 2025-10-14 |
  | `Windows 11` + `LTSC`, build 26100 | 2029-10-09 |
  | `Windows 11` por build: 22000 / 22621 / 22631 / 26100 / 26200 / 26300 (26H2) / 28000 (26H1) | 2023-10-10 / 2024-10-08 / 2025-11-11 / 2026-10-13 / 2027-10-12 / 2028-10-09 / 2028-03-14 |
  | `Windows 11`, build maior que o maior da tabela | sem data: "versão mais nova que a tabela", considerado com suporte |
  | `Windows Server 2003` / `2008` (inclui R2) / `2012` (inclui R2) | 2015-07-14 / 2020-01-14 / 2023-10-10 |
  | `Windows Server 2016` / `2019` / `2022` / `2025` | 2027-01-12 / 2029-01-09 / 2031-10-14 / **2034-11-14** (o PRD dizia 10/10; corrigido) |
  | `Windows Server` sem ano: build 14393 / 17763 / 20348 / 26100 → mesma data do ano correspondente; outro build → 25398 = 23H2 (canal anual); build maior que 26100 = "mais nova que a tabela"; outro = canal semestral | 23H2: **2026-05-12** (conferido; o 2025-10-24 foi estendido) / semestral: 2022-08-09 |
  | Sem SO, ou SO que não é Windows (Samba, NAS, Linux, macOS) | "Não Windows / desconhecido", sem alerta |

  Para Windows 11 não LTSC, o texto avisa: "datas de Home/Pro; as edições Enterprise/Education têm 12 meses a mais".
- `Get-FimSuporte([string]$so, [string]$versao)` (pura) → `{Familia; Fim ([datetime] ou $null); Obs}`.
- `Get-AchadosComputadores($comps, $hoje, $DiasInativo)` (pura) → por computador: `Familia`, `Fim`, `ForaSuporte` (ativo e `Fim` < `$hoje`), `Inativo` (ativo e sem logon há mais de `$DiasInativo` dias, ou nunca logou e foi criado antes do limite) e `OU`. Também devolve o resumo por família (total, ativos, inativos, fora de suporte) e a contagem de desativados.
- `CardsComputadores($r)`:
  - `alta`: "N computadores com sistema fora de suporte", com contagem por família (sem a lista inteira);
  - `media`: "N computadores ativos sem uso há mais de D dias";
  - `baixa`: "N contas de computador desativadas" (sugerir limpeza).
- `SecaoComputadores($r, $num)`: tabela-resumo por família, depois as listas de fora de suporte e de inativos em `.scroll`, **limitadas a 300 linhas cada**, com "+N no arquivo 07_Computadores" (RNF3).
- CSV `07_Computadores_<stamp>.csv` com **todos** os computadores: `Nome;SO;Versao;Ativo;UltimoLogon;CriadoEm;Familia;FimSuporte;ForaDeSuporte;Inativo;OU`.
- Linhas do snapshot `Computador` (Fase 2).

### `tools/Test-Modulos.ps1` (MODIFICAR)
- `Get-FimSuporte`: uma linha de teste por regra da tabela, com strings reais do AD: `Windows Server 2008 R2 Standard` / `6.1 (7601)`, `Windows Server® 2008 Standard` / `6.0 (6002)`, `Windows 10 Enterprise LTSC` / `10.0 (17763)`, `Windows 11 Pro` / `10.0 (26100)`, `Windows Server Datacenter` / `10.0 (20348)`, `Windows Server Standard` / `10.0 (25398)`, `Linux`, vazio.
- `Get-AchadosComputadores` com 5.000 objetos falsos: termina em menos de 10 s (PS 5.1) e o HTML gerado fica abaixo de 1 MB.

### Verificação Fase 6
- [x] Testes, build e lint ok.

---

## Fase 7: Interface (C#)

### `src/ScriptRunner.cs` (MODIFICAR)
- Em `Iniciar`: `args.Append(" -").Append(p.Key); if (p.Value != null) args.Append(' ').Append(Quote(p.Value));`. Comentário: "valor null = switch".

### `src/MainForm.cs` (MODIFICAR, gravar em UTF-8 **com BOM**)
- Campos: `CheckBox chkComparar, chkContas, chkSenha, chkGpos, chkComputadores; TextBox txtCompararCom, txtPadraoServico; NumericUpDown numDiasServico; Button btnCompararCom;` e `const string PadraoServico = "svc|servico|serviço|service|sql|backup";`.
- `MontarTela`, depois da linha do `chkAbrir`:
  - `AddTitulo("Itens adicionais (opcionais)")` (helper novo: `Label` em negrito, `SetColumnSpan` 4, margem superior 10);
  - `FlowLayoutPanel` com as 5 caixas, todas `Checked = false`, via `AddLinha("", fluxo)`. Textos: "Comparar com o levantamento anterior", "Contas privilegiadas e de risco", "Política de senha", "GPOs", "Computadores do domínio";
  - "Comparar com:" → `txtCompararCom` + `btnCompararCom` ("..."). Extrair o padrão de `linhaSaida` para um helper `LinhaComBotao(TextBox, Button)` e reusar nos dois. `AddDica("Vazio = levantamento mais recente do mesmo domínio na pasta de saída")`;
  - `AddPar("Contas de serviço:", txtPadraoServico, "Dias senha serviço:", numDiasServico)`, com `NovoNumero(1, 3650, 365)`.
- `AtualizarOpcionais()`: `txtCompararCom`/`btnCompararCom` habilitados = `chkComparar.Checked`; `txtPadraoServico`/`numDiasServico` habilitados = `chkContas.Checked`. Ligado ao `CheckedChanged` das duas caixas e chamado ao fim de `MontarTela` (RF7). O `Bloquear` não muda, porque desabilita o `tblParametros` inteiro e preserva o estado dos filhos.
- `EscolherSnapshot()`: `OpenFileDialog` com `InitialDirectory` na pasta de saída (se existir) e filtro `"Levantamentos do RelatorioAD (00_Snapshot_*.csv)|00_Snapshot_*.csv|Todos os arquivos (*.*)|*.*"`.
- `Gerar()`:
  - validação: com `chkContas`, `RegexValida(txtPadraoServico, "Padrão de contas de serviço")`; com `chkComparar` e `txtCompararCom` preenchido, `File.Exists`, senão `Erro(...)`;
  - parâmetros: `if (chkComparar.Checked) { p["Comparar"] = null; AddParam(p, "CompararCom", txtCompararCom.Text); }`, `if (chkContas.Checked) { p["ContasDominio"] = null; AddParam(p, "PadraoServico", ...); p["DiasSenhaServico"] = ...; }`, `if (chkSenha.Checked) p["PoliticaSenha"] = null;` e o mesmo para `GPOs` e `Computadores`. Desmarcado = nada é passado.
- Janela: `ClientSize` 900×680 → **900×760**, para o log continuar com pelo menos ~6 linhas visíveis.

### `src/AssemblyInfo.cs` (MODIFICAR)
- `AssemblyVersion` e `AssemblyFileVersion` → `1.1.0.0`.

### Verificação Fase 7
- [x] Build sem avisos (`/warnaserror+`).
- [x] Ao abrir, as 5 caixas estão desmarcadas e os 3 campos extras desabilitados; marcar "Comparar" ou "Contas" habilita os campos correspondentes (teste por `WM_SETTEXT`/`BM_CLICK` + `PrintWindow`, como na 1.0).
- [x] Linha de comando registrada no log (sem AD, rodando o 2022): sem módulos, idêntica à da 1.0; com todos marcados, contém `-Comparar -ContasDominio -PadraoServico "..." -DiasSenhaServico 365 -PoliticaSenha -GPOs -Computadores`, sem valor depois dos switches.
- [x] Em 1024×768, o log mostra pelo menos 6 linhas (~7 linhas; 5,9 descontando a barra de tarefas, ver progress.md).

---

## Fase 8: Entrega e roteiro de teste real

### `docs/ENTREGA.md` (MODIFICAR)
- Seção "Itens adicionais (opcionais)": o que cada um verifica, os arquivos 00 e 04 a 07, a observação de que o programa continua somente leitura, e que o `00_Snapshot` é gravado sempre e não deve ser apagado se o cliente quiser comparar na próxima visita.
- SHA256 novos do exe e do zip.

### `docs/features/001-auditoria-dominio/TESTE.md` (CRIAR)
Roteiro para o técnico, na ordem 2022 → 2008 R2 → 2012 R2 → 2016/2019/2025:
1. Sem nenhum item marcado: comparar o HTML e os CSVs com os da 1.0 no mesmo servidor (`fc /b` nos CSVs; nos HTML, ignorar só a data/hora).
2. Todos marcados, primeira execução: "Nenhum levantamento anterior".
3. Desativar uma conta de teste, criar uma GPO vazia, rodar de novo com Comparar → a conta aparece como "desativada" e a GPO como "nova" e "vazia".
4. Servidor membro (não-DC) num domínio filho, se houver: o aviso de Enterprise/Schema Admins aparece e não derruba o relatório.
5. Usuário sem acesso ao SYSVOL: aviso na seção GPOs, sem um "SemSysvol" por GPO.

### Verificação Fase 8
- [x] `.\build.ps1` completo: lint, testes, compilação sem avisos e Defender limpo (`dist\` e `build\scripts\`).
- [x] `ENTREGA.md` com os SHA256 do build final.
- [x] Registro em `docs/progress.md`.

## Testes
- [x] Unitário (`tools/Test-Modulos.ps1`, roda no build): ganchos vazios, snapshot (ida e volta, busca do anterior, comparação), contas, política de senha, `gPLink`, GPOs, fim de suporte e desempenho com 5.000 computadores.
- [x] Lint: os 7 scripts gerados nos alvos 2.0/3.0/5.1.
- [ ] Ponta a ponta: só nos servidores do usuário (`TESTE.md`). **Limitação conhecida:** este Windows 11 não tem o motor do PS 2.0, então os testes unitários rodam no 5.1; a compatibilidade com o 2.0 é garantida só pelo lint e pelas regras de escrita.

## Ordem de implementação
1. Fase 1 (esqueleto + build) → 2. Fase 2 (snapshot/comparar) → 3. Fase 3 (contas) → 4. Fase 4 (senha) → 5. Fase 5 (GPOs) → 6. Fase 6 (computadores) → 7. Fase 7 (interface) → 8. Fase 8 (entrega).

As fases 3 a 6 são independentes entre si. A Fase 2 depende delas só para as linhas `Privilegiado`/`GPO`/`Computador` do snapshot, que cada fase acrescenta à medida que é implementada.

## Notas / Não fazer
- **Não** usar `Get-ADGroupMember` (FSP e limite do ADWS), o módulo `GroupPolicy`, `Get-GPOReport`, RSAT extra, `ConvertTo-Json`/`ConvertFrom-Json` ou qualquer acesso à internet.
- **Não** gravar nada no AD nem no SYSVOL: só `Get-*` e `Test-Path`.
- **Não** alterar o CSS, a navegação, os KPIs, o resumo nem a linha `Concluído. ...` quando nenhum módulo estiver marcado (RF1). O `ScriptRunner` faz parse de `Concluído.` e `Relatório: `.
- **Não** usar `[regex]::Replace` para embutir o `Modulos.ps1` no build (`$` no conteúdo).
- **Não** passar switches como `-Nome:$true` pela linha de comando (falha com `-File` no PS 2–4).
- No C#: sem LINQ, `HashSet` ou `Func`/`Action` (System.Core não é referenciado); `.cs` em UTF-8 com BOM.
- `Modulos.ps1` em UTF-8 com BOM, para o arquivo poder ser lido e testado direto; o build tira o BOM ao embutir.
- Padrões a copiar: `Card`/`table.mini`/`.scroll` da seção "Pontos de atenção" (`scripts/Relatorio-PS51.ps1` ~linhas 572–622); `$csvOpts` para todos os CSVs; `NovoObj` (PS2) como modelo para criar objetos. Em `Modulos.ps1`, usar `New-Object PSObject -Property` direto, porque o `NovoObj` só existe na base PS2.
- Decisões já tomadas (PRD): caixas desmarcadas por padrão; snapshot sempre gravado; bloqueio alerta só em 0 (alta) e > 10 (baixa); Windows 11 com data de Home/Pro e aviso de Enterprise/Education; política de senha sem CSV.
- **Correção em relação ao PRD:** fim do suporte estendido do Windows Server 2025 = **2034-11-14** (Microsoft Lifecycle), e não 2034-10-10. Também foram acrescentadas as regras de Windows 10 LTSB/LTSC, Windows 11 LTSC 2024 e 26H1 (build 28000), e o Windows Server sem ano (por build).
- Conferido na Fase 6 (2026-10-28): Windows Server 23H2 (build 25398) teve suporte até 2026-05-12 (declaração de fim de atualizações da Microsoft); Windows 11 26H2 = build 26300, Home/Pro até 2028-10-09 (Lifecycle mostra 10/10/2028 06:59:59).
- **Alteração pedida pelo usuário (2026-10-28, depois da Fase 8): computadores sem logon há mais de 7 meses.** Constante `$CompMesesAntigo = 7` (sem campo na interface). `Antigo` = último logon (ou criação, se nunca logou) antes de `$hoje.AddMonths(-7)`, **ativo ou desativado** = candidato à remoção. `Inativo` passa a excluir os antigos (faixa `DiasInativo` dias a 7 meses) e `ForaSuporte` também (só ativos que logaram nos últimos 7 meses). Cards: `media` = "N computadores sem logon há mais de 7 meses (candidatos à remoção)" (ativos/desativados); o de sem uso cai para `baixa`; o de desativados cita quantos já estão na lista de remoção. Seção: coluna nova no resumo e lista própria (Situação, Último logon, Dias sem logon, Criado em, OU), do mais antigo para o mais recente, antes da lista de sem uso. CSV 07 ganha `DiasSemLogon` (vazio se nunca logou) e `CandidatoRemocao`: `Nome;SO;Versao;Ativo;UltimoLogon;DiasSemLogon;CriadoEm;Familia;FimSuporte;ForaDeSuporte;Inativo;CandidatoRemocao;OU`. Snapshot inalterado.
- **Alteração pedida pelo usuário (2026-10-28, depois da Fase 8): explicação e conteúdo das GPOs.** A seção GPOs ganha uma explicação de como as GPOs funcionam e do que o relatório lê, e um bloco "O que cada GPO faz" por GPO. O conteúdo vem dos arquivos no SYSVOL (`GptTmpl.inf`, `audit.csv`, `Registry.pol`, `scripts.ini`/`psscripts.ini`, XMLs de `Preferences\`, `fdeploy1.ini`), dos atributos `gPCMachineExtensionNames`/`gPCUserExtensionNames` (áreas sem detalhe) e dos `packageRegistration` no AD (instalação de software). Continua **sem** o módulo GroupPolicy e sem `Get-GPOReport`. Novas funções puras: `Get-CseGpo`, `ConvertFrom-Ini`, `ConvertFrom-GptTmpl`, `ConvertFrom-RegistryPol`, `ConvertTo-ItensRegistro`, `ConvertFrom-ScriptsIni`, `ConvertFrom-PreferenciaXml`, `ConvertFrom-Fdeploy`, `Get-ItensGpo`; leitura em `Read-ConteudoGpo` (um `try` por arquivo, erro vira `ErrosLeitura`). Achado novo `SenhaPreferencia` (cpassword preenchido, MS14-025), card `alta`; o valor nunca é lido para o relatório. Aviso de política de senha vinculada fora da raiz do domínio. HTML limitado a 15 linhas por área (`$GpoLimiteDetalhes`); CSV novo `06_GPOs_Configuracoes_<stamp>.csv` (`GPO;GUID;Parte;Area;Configuracao`) com tudo; o 06 ganha `Configuracoes` e `ErrosLeitura`. Snapshot inalterado.
