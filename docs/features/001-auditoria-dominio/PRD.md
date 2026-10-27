# PRD: Auditoria do domínio (módulos opcionais)

> Data: 2026-10-27 · Status: aprovado · Projeto: RelatorioAD 1.0 → 1.1

## 1. Problema
Hoje o relatório cobre usuários ativos, administradores do servidor e permissões das pastas. Numa revisão de segurança com o cliente, o técnico ainda precisa levantar à mão: o que mudou desde a última visita, as contas privilegiadas do domínio, a política de senha, as GPOs e os computadores do domínio. Isso fica fora do relatório e não tem padrão.

## 2. Objetivo e resultado esperado
- Cinco **módulos opcionais**, cada um com uma caixa de seleção **desmarcada por padrão** na tela (pedido do usuário: "que eu tenha que marcar caso queira a informação"):
  1. **Comparar com o levantamento anterior**
  2. **Contas privilegiadas e de risco no domínio**
  3. **Política de senha e bloqueio**
  4. **GPOs**
  5. **Computadores do domínio**
- Com todos desmarcados, o HTML e os CSVs saem **iguais aos de hoje**. A única diferença é o arquivo de snapshot da seção 4.
- Os módulos marcados entram no **mesmo HTML**: os achados vão para "Pontos de atenção" e cada módulo ganha a sua seção, além de um CSV próprio.
- Funciona nas 7 versões (2008 R2 a 2025), com as mesmas regras de compatibilidade de hoje.

## 3. Usuários e cenários de uso
- **Técnico Compumaq/Werner**: roda o programa no DC ou no servidor de arquivos do cliente, marca os módulos que quer e gera o relatório. Nas visitas periódicas, marca "Comparar" para mostrar ao cliente o que mudou desde o último levantamento.

## 4. Escopo
### Dentro
**Módulo 1 — Comparar com o levantamento anterior**
- Toda execução grava `00_Snapshot_<data>.csv` na pasta de saída, mesmo com o módulo desmarcado, para servir de base à próxima comparação. O arquivo é pequeno, em CSV legível.
- Com o módulo marcado, a comparação é feita com o snapshot mais recente da pasta de saída do **mesmo domínio**, excluindo o atual. Na tela também dá para escolher outro arquivo (`Comparar com...`); vazio = automático.
- Mudanças mostradas:
  - contas ativas novas e contas que saíram (desativadas ou excluídas);
  - administradores do servidor que entraram ou saíram;
  - acesso efetivo por pasta: quem ganhou acesso, quem perdeu e quem mudou de nível;
  - pastas compartilhadas novas e removidas;
  - membros dos grupos privilegiados do domínio, GPOs e computadores, quando o módulo correspondente estiver marcado nas **duas** execuções.
- Seção "Mudanças desde dd/mm/aaaa", um card-resumo em "Pontos de atenção" e o CSV `04_Mudancas_<data>.csv`.
- Sem snapshot anterior: a seção diz "Nenhum levantamento anterior encontrado nesta pasta" e o relatório segue sem erro.

**Módulo 2 — Contas privilegiadas e de risco no domínio**
- Membros (com expansão recursiva) de: Domain Admins (`<domínio>-512`), Enterprise Admins (`<raiz>-519`), Schema Admins (`<raiz>-518`), Administrators do domínio (`S-1-5-32-544`), Account Operators (`-548`), Server Operators (`-549`), Print Operators (`-550`) e Backup Operators (`-551`). Os grupos são identificados por SID, para funcionar em AD em português.
- Checagens por conta ativa:
  - **senha não exigida** (`PasswordNotRequired`, flag `PASSWD_NOTREQD`);
  - **delegação irrestrita** (`TrustedForDelegation`), sem contar os DCs;
  - **pré-autenticação Kerberos desativada** (`DoesNotRequirePreAuth`, vulnerável a AS-REP roasting);
  - **senha com criptografia reversível** (`AllowReversiblePasswordEncryption`);
  - **contas que nunca fizeram logon**, criadas há mais de `DiasInativo` dias;
  - **Administrador interno** (`<domínio>-500`): ativo, nome não alterado, data da última troca de senha.
- **Contas de serviço**: têm `servicePrincipalName` **ou** o login/OU bate com o novo parâmetro `PadraoServico` (padrão `svc|servico|serviço|service|sql|backup`). Alerta quando a senha tem mais de `DiasSenhaServico` dias (padrão 365).
- Seção "Contas privilegiadas do domínio", cards de atenção e o CSV `05_Contas_Privilegiadas_<data>.csv`.

**Módulo 3 — Política de senha e bloqueio**
- Política padrão do domínio (`Get-ADDefaultDomainPasswordPolicy`) e políticas refinadas (FGPP, `Get-ADFineGrainedPasswordPolicy -Filter *`), com a quem se aplicam.
- Comparação com a **Microsoft Security Baseline** (decisão do usuário):

  | Item | Referência | Severidade se pior |
  |---|---|---|
  | Tamanho mínimo | ≥ 14 | alta se < 8, média se 8–13 |
  | Complexidade | ligada | alta |
  | Histórico | ≥ 24 | baixa |
  | Criptografia reversível | desligada | alta |
  | Bloqueio por tentativas | 1 a 10 (a baseline do 2025 usa 3) | alta se 0 (sem bloqueio), baixa se > 10 |
  | Duração do bloqueio / zerar contador | ≥ 15 min | baixa |
  | Idade mínima | ≥ 1 dia | baixa |
  | Validade máxima | só informativo (a Microsoft deixou de recomendar expiração periódica) | — |

- Seção "Política de senha" com a tabela valor atual × referência, explicada em linguagem simples, mais um card em "Pontos de atenção" quando houver itens abaixo da referência.

**Módulo 4 — GPOs (lidas pelo LDAP, sem depender do módulo GroupPolicy)**
- GPOs: `groupPolicyContainer` em `CN=Policies,CN=System,<domínio>`, com nome, GUID, status (`flags`: 0 = ativa, 1 = usuário desativado, 2 = computador desativado, 3 = toda desativada), criação, alteração e versão.
- Vínculos: atributo `gPLink` do domínio e das OUs (`[LDAP://cn={GUID},...;opção]`, onde a opção 0 = ativo, 1 = desativado, 2 = imposto e 3 = desativado + imposto), mais o bloqueio de herança (`gPOptions = 1`).
- Achados:
  - GPO **sem nenhum vínculo**;
  - GPO **toda desativada**;
  - **vínculos desativados**;
  - **vínculo para GPO inexistente** (órfão no `gPLink`);
  - GPO **vazia** (versão 0);
  - **pasta da GPO ausente no SYSVOL** (`gPCFileSysPath` inacessível).
- Seção "GPOs" (tabela GPO × onde está vinculada) e o CSV `06_GPOs_<data>.csv`.

**Módulo 5 — Computadores do domínio**
- `Get-ADComputer -Filter *` com nome, SO, versão, ativo, último logon, criação e OU.
- Achados:
  - computadores ativos **sem logon há mais de `DiasInativo` dias**;
  - **SO fora de suporte** em relação à data do levantamento, por uma tabela de fim de suporte embutida no script (abaixo);
  - contas de computador **desativadas** (contagem).
- Seção "Computadores" com resumo por SO, a lista de computadores fora de suporte e de inativos, e o CSV `07_Computadores_<data>.csv`.

**Interface (C#)**
- Novo grupo "Itens adicionais (opcionais)" com 5 caixas desmarcadas, mais o campo `Comparar com...` (habilitado só com "Comparar" marcado) e os campos `Padrão de contas de serviço` e `Dias senha de serviço` (habilitados só com "Contas privilegiadas" marcado).
- Os módulos viram parâmetros do script (`-Comparar`, `-CompararCom`, `-ContasDominio`, `-PoliticaSenha`, `-GPOs`, `-Computadores`, `-PadraoServico`, `-DiasSenhaServico`), passados só quando marcados.

### Fora (não fazer agora)
- Exportar em PDF, vários servidores, agendamento, e-mail e perfis de cliente (itens 4, 5, 8 a 11 da lista de ideias).
- Ler o conteúdo das GPOs (configurações dentro de cada GPO): seria preciso o `Get-GPOReport`/GPMC, que nem sempre está instalado.
- Idade da senha do `krbtgt`, LAPS, níveis funcionais, confianças entre domínios: ficam para uma próxima feature.
- Corrigir qualquer coisa: o programa continua **somente leitura**.

## 5. Restrições
- Mesma stack: 3 bases de script (PS 5.1, PS 3 e PS 2) e app C# 3 / csc v3.5, sem libs externas.
- Os módulos usam só o módulo `ActiveDirectory`, que já é exigido. Nada de módulo GroupPolicy, RSAT extra ou internet.
- Somente leitura no AD e no SYSVOL.
- Código novo compatível com **PowerShell 2.0**: sem `ConvertTo-Json` (o snapshot é em CSV, lido com `Import-Csv -Delimiter ';'`, que existe no PS 2.0) e com as regras do lint, incluindo o `foreach` sobre `@(...)` e o cuidado com variáveis `$null`.

## 6. Contexto técnico encontrado
### Código existente para reutilizar / copiar padrão
- `scripts/Relatorio-PS51.ps1`, `-PS3.ps1` e `-PS2.ps1`: as 3 bases.
  - Pontos de encaixe: `param()` (linha ~19); fim do cálculo `[4/5]` (~linha 305); HTML com `A`, `Card($sev,$titulo,$corpo)`, `HE`, `NM`, `DT` (~linhas 338–365 e 526+).
  - Navegação `<nav class="toc">` (~linha 545), "Pontos de atenção" (`$nCards`, ~linha 572), seção de usuários (~linha 692) e metodologia (~linha 712).
- `Resolve-Principal` / `UsersOf` / `Get-NomeSid`: expansão de grupos (reusar para os grupos privilegiados do domínio).
- `$csvOpts = @{ NoTypeInformation; Encoding='UTF8'; Delimiter=';' }`: padrão dos CSVs.
- `$U` (login → objeto), `$ADM`, `$Calc[share].EF` (acesso efetivo): a origem dos dados do snapshot.
- `tools/Test-ScriptCompat.ps1`: o lint; o código novo tem que passar no alvo 2.0.
- `src/MainForm.cs`: layout em `TableLayoutPanel` de 4 colunas e a montagem dos parâmetros em `Gerar` (juntar com `|`, não passar os vazios).

**Proposta de arquitetura (para a Spec confirmar):** escrever os 5 módulos **uma única vez**, num arquivo `scripts/Modulos.ps1` compatível com PS 2.0. O `build.ps1` **embute** esse arquivo nas 3 bases num marcador (`#region MODULOS`), de modo que cada script gerado continua sendo **um arquivo só**, legível. As bases ganham só os parâmetros e as chamadas aos módulos (coleta depois do `[4/5]`, mais os cards, a navegação e as seções no HTML). Isso evita triplicar o código e o risco do PS 2.0 em três cópias.

### Documentação e APIs (com fonte)
- `Get-ADDefaultDomainPasswordPolicy` e `Get-ADFineGrainedPasswordPolicy` existem no módulo AD do **Windows Server 2008 R2**: [Retrieve Details of a Default Domain Password Policy (2008 R2)](https://learn.microsoft.com/en-us/previous-versions/windows/it-pro/windows-server-2008-r2-and-2008/dd378870(v=ws.10)) · [Use PowerShell to Configure Domain Password Policy](https://devblogs.microsoft.com/scripting/use-windows-powershell-to-configure-domain-password-policy).
  ```powershell
  Get-ADDefaultDomainPasswordPolicy   # MinPasswordLength, PasswordHistoryCount, ComplexityEnabled,
                                      # MaxPasswordAge, MinPasswordAge, LockoutThreshold,
                                      # LockoutDuration, LockoutObservationWindow, ReversibleEncryptionEnabled
  Get-ADFineGrainedPasswordPolicy -Filter * -Properties AppliesTo, Precedence
  ```
- Baseline: o mínimo de 14 caracteres está no [CIS Windows Server 2025 Benchmark](https://itsecure.hu/wp-content/uploads/2025/05/CIS_Microsoft_Windows_Server_2025_Benchmark_v1.0.0.pdf). A baseline da Microsoft para o 2025 baixou o bloqueio de 10 para **3** tentativas, com 15 min ([anúncio da Windows Server 2025 Security Baseline](https://techcommunity.microsoft.com/blog/microsoft-security-baselines/-/4358733), [reprodução no ElevenForum](https://www.elevenforum.com/t/security-baseline-package-for-windows-server-2025-now-available.32918/latest)), e o STIG do 2025 pede ≤ 3 ([V-278034](https://www.stigviewer.com/stigs/microsoft_windows_server_2025/2026-02-20/finding/V-278034)). **Correção da pergunta feita ao usuário:** lá eu disse "10 tentativas"; a referência atual é 3. A tabela da seção 4 usa "1 a 10 = aceitável, 3 = ideal".
- Atributos de conta no módulo AD (2008 R2+): `PasswordNotRequired`, `TrustedForDelegation`, `DoesNotRequirePreAuth`, `AllowReversiblePasswordEncryption`, `ServicePrincipalNames`, `PasswordLastSet`, `LastLogonDate`, `whenCreated`.
- GPO via LDAP (sem o GroupPolicy):
  ```powershell
  Get-ADObject -SearchBase "CN=Policies,CN=System,$($dom.DistinguishedName)" -LDAPFilter '(objectClass=groupPolicyContainer)' `
      -Properties displayName, flags, versionNumber, gPCFileSysPath, whenCreated, whenChanged
  Get-ADObject -LDAPFilter '(gPLink=*)' -Properties gPLink, gPOptions   # domínio + OUs
  ```
- Fim de suporte (tabela embutida, comparada com a data do levantamento):

  | Sistema | Fim do suporte |
  |---|---|
  | Windows XP / Vista / 7 / 8 / 8.1 | 2014-04-08 / 2017-04-11 / 2020-01-14 / 2016-01-12 / 2023-01-10 |
  | Windows 10 (22H2, sem ESU) | 2025-10-14 |
  | Windows 11 Home/Pro 21H2 / 22H2 / 23H2 / 24H2 / 25H2 (por build 22000 / 22621 / 22631 / 26100 / 26200) | 2023-10-10 / 2024-10-08 / 2025-11-11 / 2026-10-13 / 2027-10-12 |
  | Windows Server 2003 / 2008 e 2008 R2 / 2012 e 2012 R2 | 2015-07-14 / 2020-01-14 / 2023-10-10 |
  | Windows Server 2016 / 2019 / 2022 / 2025 | 2027-01-12 / 2029-01-09 / 2031-10-14 / 2034-10-10 |

  Fonte do 24H2: [BleepingComputer](https://bleepingcomputer.com/news/microsoft/windows-11-24h2-home-and-pro-reach-end-of-support-in-90-days), [Digital Citizen](https://www.digitalcitizen.life/windows-11-24h2-support-ends-october-13-leaving-home-and-pro-pcs-without-security-updates/); as edições Enterprise/Education do 24H2 vão até 2027-10-12. As demais datas são as do Microsoft Lifecycle e precisam ser **conferidas na Spec**. O AD não informa a edição do Windows 11, então o relatório usa a data de Home/Pro e avisa que Enterprise/Education têm um ano a mais.

## 7. Requisitos
### Funcionais
- RF1: Cinco caixas de seleção, desmarcadas por padrão. Sem nenhuma marcada, o HTML e os CSVs 01/02/03/99 ficam iguais aos de hoje.
- RF2: Toda execução grava `00_Snapshot_<data>.csv`, com o domínio, a data, a versão do formato e as linhas `Tipo;Chave;Valor`.
- RF3: "Comparar" usa o snapshot mais recente do mesmo domínio na pasta de saída, ou o arquivo escolhido. Não havendo nenhum, avisa e segue.
- RF4: Cada módulo marcado gera: a seção no HTML (com link na navegação), os cards em "Pontos de atenção" quando houver achados, e o seu CSV (04 a 07). O módulo 3 só tem seção e cards, sem CSV.
- RF5: A falha de um módulo (ex.: sem permissão para ler o SYSVOL) é registrada no log e na seção do módulo ("não foi possível ler: ..."), sem derrubar o relatório.
- RF6: Os textos explicam cada achado em linguagem simples, como os cards atuais.
- RF7: A interface mostra os campos extras só quando o módulo correspondente está marcado (ou os deixa desabilitados).

### Não funcionais
- RNF1: Compatível com PS 2.0 a 5.1; o lint aprova os 7 scripts gerados.
- RNF2: Sem novos módulos do PowerShell nem dependências; somente leitura.
- RNF3: Desempenho: o módulo de computadores deve aguentar ~5.000 computadores sem pesar no HTML. Mostrar no HTML só os achados e a contagem por SO; a lista completa fica no CSV.
- RNF4: O Defender continua limpo e o build sem avisos.

## 8. Riscos e dúvidas em aberto
- [x] **Limiar de bloqueio**: alerta só para 0 (sem bloqueio, alta) e > 10 (baixa); de 1 a 10 é aceitável, e o texto cita que a baseline do 2025 usa 3. (Aprovado com o PRD.)
- [x] **Snapshot sempre gravado**, mesmo com "Comparar" desmarcado. (Aprovado com o PRD.)
- [ ] Em servidor **membro** (não-DC), Enterprise/Schema Admins ficam no domínio raiz da floresta. Num domínio filho, essa leitura pode exigir consulta ao DC raiz; se falhar, avisar na seção.
- [ ] Validação real só nos servidores do usuário; o 2008 R2 continua sendo o de maior risco. Testar primeiro o 1.0 corrigido, depois a 1.1.
- [ ] Conferir as datas de fim de suporte no Microsoft Lifecycle durante a Spec.
