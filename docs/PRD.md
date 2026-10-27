# PRD: RelatorioAD — Relatório de acessos multi-versão do Windows Server

> Data: 2026-10-27 · Status: aprovado

## 1. Problema
A Compumaq/Werner tem um script PowerShell que gera o relatório de usuários ativos do AD e das permissões de compartilhamento/NTFS em HTML, mas ele só funciona no Windows Server 2016–2025 (PowerShell 5.1). Os clientes têm servidores do 2008 R2 ao 2025. Hoje o técnico precisa adaptar o script na mão e rodá-lo pelo console, o que é lento e sujeito a erro.

## 2. Objetivo e resultado esperado
- Um único `.exe` com interface gráfica que roda em qualquer Windows Server do 2008 R2 ao 2025 sem instalar nada.
- O técnico escolhe a versão do SO (com detecção automática como sugestão), preenche os parâmetros e gera o mesmo relatório HTML + CSVs do script atual.
- Cada versão de SO tem o seu próprio script, aproveitando o que a versão oferece (ex.: `Get-SmbShare` e `-Depth` onde existem) e com fallback compatível onde não existem.
- O programa não é bloqueado pelo Windows Defender e reduz ao máximo os falsos positivos em outros antivírus.

## 3. Usuários e cenários de uso
- **Técnico Compumaq/Werner**: copia o `.exe` para o servidor do cliente (DC ou servidor de arquivos), roda como administrador, escolhe a versão, clica em "Gerar relatório" e abre o HTML para apresentar ao cliente.

## 4. Escopo
### Dentro
- Interface WinForms com:
  - lista de versões: 2008 R2, 2012, 2012 R2, 2016, 2019, 2022, 2025, com a versão detectada pré-selecionada e aviso se a escolha divergir do SO real;
  - campos com os parâmetros do script: Cliente, Responsável, Profundidade, Pasta de saída, Caminhos (opcional), Dias inativo, Padrão de terceiros, Pastas sensíveis, Abrir ao final;
  - botão "Gerar relatório", área de log ao vivo (saída do script), barra de status e botões "Abrir relatório" e "Abrir pasta";
  - botão "Exportar script" (salva o .ps1 da versão escolhida, para auditoria ou execução manual).
- 7 scripts embutidos (um por versão), todos gerando o mesmo HTML/CSV do script de referência:
  - **2016 / 2019 / 2022 / 2025** (PS 5.1): o script atual, com ajustes pontuais por versão.
  - **2012 R2** (PS 4.0) e **2012** (PS 3.0): sem `-Depth` no `Get-ChildItem` → recursão manual limitada por profundidade; `Get-SmbShare` disponível.
  - **2008 R2** (PS 2.0, CLR 2): sem `SmbShare` → `Win32_Share` + `Win32_LogicalShareSecuritySetting` (WMI); sem `[pscustomobject]`, `[ordered]`, `-Directory`, `-in`, `SortedSet`, `WebUtility` → equivalentes compatíveis.
- Verificações antes de rodar: executando como administrador, módulo `ActiveDirectory` disponível, versão do PowerShell compatível com o script escolhido.
- Medidas anti-falso-positivo (ver 6.3).
- Script de build (`build.ps1`) que compila o .exe com o `csc.exe` do Windows, sem Visual Studio.

### Fora (não fazer agora)
- Execução remota em outro servidor (o script roda local).
- Certificado comercial de assinatura de código (o usuário não tem; o build deixa a etapa de assinatura pronta para usar depois).
- Suíte de testes automatizados contra AD real ou mocks (o usuário pediu para pular os testes). Fica só a verificação mínima: compilação sem erros e análise sintática dos scripts.
- Instalador (.msi). O programa é um .exe portátil.

## 5. Restrições
- Stack: C# para .NET Framework 3.5/CLR 2 (roda também no CLR 4), WinForms, `csc.exe` em `C:\Windows\Microsoft.NET\Framework64\v3.5\` (disponível nesta máquina).
- Onde roda: local no Windows Server 2008 R2 → 2025, x64, como administrador.
- Custo: zero (sem libs pagas, sem certificado por enquanto).

## 6. Contexto técnico encontrado
### 6.1 Código existente para reutilizar
- `referencia/Relatorio-AD-Permissoes-2022.ps1` — script base fornecido pelo usuário (2016–2025). **Precisa ser salvo pelo usuário nessa pasta** antes da Spec. É a fonte do HTML/CSS, do cálculo de acesso efetivo e dos CSVs; todas as variantes devem produzir o mesmo resultado.

### 6.2 Compatibilidade por versão
| SO | PowerShell nativo | .NET nativo | SmbShare | `-Depth` / `-Directory` | Observações |
|---|---|---|---|---|---|
| 2008 R2 | 2.0 | 3.5.1 (CLR 2) | não (WMI) | não / não | sem `[pscustomobject]`/`[ordered]`; `HashSet` exige `System.Core`; não existe `SortedSet` |
| 2012 | 3.0 | 4.5 | sim | não / sim | `[pscustomobject]`, `[ordered]`, `-in` disponíveis |
| 2012 R2 | 4.0 | 4.5.1 | sim | não / sim | igual ao 2012 |
| 2016 | 5.1 | 4.6.2 | sim | sim / sim | script atual |
| 2019 | 5.1 | 4.7.2 | sim | sim / sim | script atual |
| 2022 | 5.1 | 4.8 | sim | sim / sim | script de referência |
| 2025 | 5.1 (+ pwsh opcional) | 4.8.1 | sim | sim / sim | rodar no Windows PowerShell 5.1 (`powershell.exe`), não no pwsh |

Fonte: documentação da Microsoft sobre versões do PowerShell e do .NET Framework por versão do Windows (a confirmar links na Spec).

Detecção do SO no C#: ler `Environment.OSVersion.Version` (6.1 = 2008 R2, 6.2 = 2012, 6.3 = 2012 R2) e, a partir do 10.0, o `CurrentBuild` em `HKLM\SOFTWARE\Microsoft\Windows NT\CurrentVersion` (14393 = 2016, 17763 = 2019, 20348 = 2022, 26100 = 2025).

### 6.3 Antivírus — o que reduz falso positivo
- **Não usar** PS2EXE, empacotadores (UPX), ofuscação, `-EncodedCommand`, scripts em base64 nem `-WindowStyle Hidden`. São os padrões que mais disparam heurísticas.
- Executar os scripts de forma transparente: extrair o .ps1 para uma pasta visível (`<Saída>\scripts\`) e rodar `powershell.exe -NoProfile -ExecutionPolicy Bypass -File "<caminho>"` com a saída mostrada na janela. Assim o AMSI/Defender vê um script legível.
- Metadados completos no .exe (`AssemblyCompany`, `AssemblyProduct`, `AssemblyVersion`, descrição, ícone) e manifest com `requestedExecutionLevel="requireAdministrator"`, sem pedir elevação por truques.
- Sem acesso à internet a partir do .exe (o HTML já carrega a fonte do Google Fonts, o que é normal para um navegador).
- **Assinatura:** o `build.ps1` terá a etapa de assinatura (Authenticode, `Set-AuthenticodeSignature` para .exe e .ps1) ativada quando houver certificado. Opção interna: certificado autoassinado distribuído por GPO para "Editores Confiáveis" (vale só dentro da rede do cliente, não cria reputação no SmartScreen).
- Se o Defender acusar mesmo assim: enviar o .exe ao portal Microsoft Security Intelligence (WDSI) como falso positivo. É o caminho oficial e costuma ser resolvido em 1–3 dias.

## 7. Requisitos
### Funcionais
- RF1: Detectar a versão do Windows Server e pré-selecioná-la; permitir trocar manualmente.
- RF2: Avisar (sem bloquear) quando a versão escolhida for diferente da detectada.
- RF3: Expor todos os parâmetros do script na interface, com os mesmos valores padrão.
- RF4: Executar o script da versão escolhida mostrando o log ao vivo, sem travar a janela; permitir cancelar.
- RF5: Ao terminar, mostrar o resumo (usuários, pastas, administradores, sem uso) e habilitar "Abrir relatório" e "Abrir pasta".
- RF6: Exportar o script da versão escolhida para um arquivo.
- RF7: Os 7 scripts geram o mesmo HTML e os mesmos CSVs (01, 02, 03, 99) do script de referência.

### Não funcionais
- RNF1: Um único .exe, com menos de 1 MB, sem dependências, rodando do 2008 R2 ao 2025.
- RNF2: Sem padrões típicos de malware (ver 6.3). O Defender não deve acusar o .exe nem os scripts nesta máquina.
- RNF3: Mensagens de erro claras (sem módulo AD, sem privilégio de administrador, sem compartilhamentos).
- RNF4: Verificação mínima: o build compila sem avisos e cada script passa no parser do PowerShell (`[System.Management.Automation.Language.Parser]::ParseFile`).

## 8. Riscos e dúvidas em aberto
- [ ] Usuário precisa salvar o script de referência em `C:\RelatorioAD\referencia\Relatorio-AD-Permissoes-2022.ps1`.
- [ ] Sem teste real, as variantes 2008 R2/2012/2012 R2 só serão validadas quando rodarem num servidor dessas versões. O 2008 R2 é o de maior risco (reescrita via WMI e PS 2.0).
- [x] 2008 R2 com WMF 5.1 instalado: usar o script 2008 R2 mesmo assim. Ele é escrito para PS 2.0 e lê os compartilhamentos via WMI, então funciona em qualquer versão de PowerShell daquele SO. (Decisão: "usar o que funciona".)
- [x] 2025: mesmo script do 2022, que já funciona em PS 5.1. Muda só a identificação da versão no relatório. (Decisão: "usar o que funciona".)

