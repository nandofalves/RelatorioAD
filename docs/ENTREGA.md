# RelatorioAD 1.1 — Entrega

Relatório de usuários ativos do AD e das permissões de compartilhamento/NTFS, em HTML, para Windows Server 2008 R2 a 2025.

## Conteúdo do pacote

`RelatorioAD.zip` contém **dois arquivos, que precisam ficar juntos na mesma pasta**:

| Arquivo | Para que serve |
|---|---|
| `RelatorioAD.exe` | O programa (WinForms, .NET Framework, sem instalação). |
| `RelatorioAD.exe.config` | Diz ao Windows que o programa roda no .NET 4.x **ou** no 2.0/3.5. Sem ele, no 2012 em diante (onde o .NET 3.5 vem desativado) o Windows pede para instalar o .NET 3.5. |

## Como usar

1. Copie o `RelatorioAD.zip` para o servidor (o DC ou o servidor de arquivos membro).
2. Se o zip veio pela internet ou por e-mail, antes de extrair: **botão direito → Propriedades → marcar "Desbloquear" → OK** (ou `Unblock-File .\RelatorioAD.zip`). Isso evita o aviso do SmartScreen e do "editor desconhecido" em cada arquivo extraído.
3. Extraia os dois arquivos para uma pasta local (ex.: `C:\RelatorioAD`).
4. Abra o `RelatorioAD.exe`. Ele pede elevação (UAC) sozinho, porque precisa rodar como administrador.
5. Confira a linha **Detectado**: versão do Windows, PowerShell, módulo AD e administrador. A versão detectada já vem selecionada na lista. Se aparecer o aviso amarelo, leia antes de continuar (versão diferente da detectada, PowerShell antigo, módulo AD ausente ou sem privilégio de administrador).
6. Preencha os parâmetros (todos têm padrão) e clique em **Gerar relatório**. O log mostra o andamento ao vivo e, no final, o relatório abre no navegador (se a caixa "Abrir relatório ao terminar" estiver marcada).

Requisitos no servidor:
- Módulo **Active Directory para PowerShell** (RSAT-AD-PowerShell). Num DC ele já existe. Num servidor membro: `Install-WindowsFeature RSAT-AD-PowerShell` (2012+) ou `Import-Module ServerManager; Add-WindowsFeature RSAT-AD-PowerShell` (2008 R2).
- Se não houver compartilhamentos, informe as pastas no campo **Caminhos** (uma por linha).
- Se o log mostrar erro de política de execução (`AuthorizationManager` / `execution policy`), uma GPO está forçando a política (`MachinePolicy`/`UserPolicy`) e o `-ExecutionPolicy Bypass` não tem efeito. Peça uma exceção ao responsável pela GPO ou use a versão assinada (veja "Certificado").

Onde ficam os arquivos: na **Pasta de saída** (padrão `C:\Relatorio_AD`) ficam o HTML, os CSVs e a subpasta `scripts\` com o `.ps1` que foi executado.

## Itens adicionais (opcionais)

Novos na 1.1. Ficam no grupo **Itens adicionais (opcionais)**, todos **desmarcados** por padrão. Desmarcados, o HTML e os CSVs 01/02/03/99 saem iguais aos da 1.0. Cada item marcado acrescenta um card em "Pontos de atenção" (quando há achado), uma seção numerada no HTML e, em alguns casos, um CSV. Se um item falhar (sem acesso, módulo AD sem o cmdlet, domínio raiz fora do ar), o log mostra `[AVISO]`, a seção diz o que não pôde ser lido e o restante do relatório sai normalmente.

| Item | O que verifica | Arquivo |
|---|---|---|
| **Comparar com o levantamento anterior** | O que mudou desde a última execução no mesmo domínio: contas ativas novas, desativadas e excluídas; quem entrou ou saiu dos administradores do servidor; pastas novas e removidas; acessos ganhos, perdidos ou com nível diferente; e, se o item correspondente estava marcado nas duas execuções, grupos privilegiados, GPOs e computadores. O campo **Comparar com** escolhe um levantamento específico; vazio = o mais recente do mesmo domínio na pasta de saída. | `04_Mudancas_<data>.csv` (só quando há levantamento anterior) |
| **Contas privilegiadas e de risco** | Membros (inclusive indiretos) de Domain Admins, Administrators, Account/Server/Print/Backup Operators, Enterprise Admins e Schema Admins; contas com senha não exigida, pré-autenticação Kerberos desativada ou criptografia reversível; contas e computadores com delegação irrestrita (fora os DCs); contas de terceiros (OUs de terceiros/fornecedores) com privilégio; privilegiadas sem uso; contas de serviço (com SPN ou que casam com o padrão do campo **Contas de serviço**, pelo login ou pela OU) com senha mais antiga que **Dias senha serviço**; o Administrador interno com o nome padrão; contas que nunca fizeram logon. | `05_Contas_Privilegiadas_<data>.csv` |
| **Política de senha** | A política padrão do domínio (tamanho mínimo, complexidade, histórico, criptografia reversível, bloqueio por tentativas, duração do bloqueio, idade mínima e validade) comparada com uma referência, e as políticas refinadas (FGPP) com precedência e a quem se aplicam. | sem CSV (só a seção do HTML) |
| **GPOs** | A seção abre com uma explicação de como as GPOs funcionam (vínculo, partes Computador/Usuário, ordem de aplicação, imposição e bloqueio de herança) e do que o relatório lê ou não lê. Depois vêm a lista de GPOs com os vínculos e um bloco **"O que cada GPO faz"**, com as configurações de cada parte (Computador/Usuário) por área: política de senha e bloqueio, auditoria, direitos de usuário, opções de segurança, grupos restritos, serviços, Modelos administrativos agrupados por assunto (Windows Update, Firewall, Área de Trabalho Remota...), scripts, mapeamento de unidades, impressoras, usuários e grupos locais, tarefas agendadas, instalação de software, redirecionamento de pastas e as demais Preferências. Tudo é lido do AD e dos arquivos de cada GPO no SYSVOL, **sem o módulo GroupPolicy**. Achados: **senha gravada em Preferências (cpassword, MS14-025)**, vínculos para GPOs inexistentes, GPOs sem a pasta no SYSVOL, GPOs sem vínculo, totalmente desativadas, vínculos desativados, GPOs vazias (versão 0) e OUs com herança bloqueada. Também avisa quando uma política de senha está vinculada fora da raiz do domínio (onde ela não vale para as contas do domínio). O HTML mostra até 15 linhas por área. O `06_GPOs_Configuracoes` tem todas. | `06_GPOs_<data>.csv`, `06_GPOs_Configuracoes_<data>.csv` |
| **Computadores do domínio** | Sistema de cada computador comparado com a data de fim de suporte da Microsoft (tabela embutida, conferida em 28/10/2026). Os computadores **sem logon há mais de 7 meses** (ativos ou já desativados) ficam numa lista própria de **candidatos à remoção**, do mais antigo para o mais recente, separados dos que ainda logaram nesse período. À parte, os ativos sem uso há mais de **Dias inativo** dias e menos de 7 meses (verificar antes de desativar) e a contagem de contas desativadas. Fora de suporte só conta os ativos que logaram nos últimos 7 meses. No HTML as listas ficam limitadas a 300 linhas; o CSV tem todos, com as colunas `DiasSemLogon` e `CandidatoRemocao` para filtrar no Excel. | `07_Computadores_<data>.csv` |

Observações:
- **O programa continua somente leitura.** Os itens adicionais só consultam o AD (`Get-ADUser`, `Get-ADGroup`, `Get-ADObject`, `Get-ADComputer`, `Get-ADDomain`, `Get-ADForest`, políticas de senha) e leem os arquivos das GPOs no SYSVOL (`Test-Path`, leitura de `GptTmpl.inf`, `Registry.pol`, `scripts.ini` e XMLs de Preferências). Nada é gravado no AD nem no SYSVOL. O valor das senhas encontradas em Preferências nunca vai para o relatório: só o nome do arquivo.
- **O `00_Snapshot_<data>.csv` é gravado em toda execução**, marcada ou não a comparação. Ele é o "levantamento" que a próxima execução usa para comparar. **Não apague esses arquivos** se o cliente quiser ver as mudanças na próxima visita; para comparar, rode de novo com a mesma pasta de saída (ou aponte o arquivo no campo **Comparar com**).
- As listas de grupos privilegiados, GPOs e computadores entram no snapshot só quando o item correspondente está marcado. Para comparar esses itens, marque-os nas duas execuções.

## Antivírus

### O que o programa faz
- Mostra uma janela, lê o registro (versão do Windows e do PowerShell) e verifica se o módulo AD existe.
- Grava em `<Pasta de saída>\scripts\Relatorio-<versão>.ps1` o script PowerShell daquela versão. Esse arquivo é **texto legível**, idêntico ao que o botão **Exportar script...** salva, e pode ser conferido antes ou depois da execução.
- Executa esse arquivo com `powershell.exe -NoProfile -NonInteractive -ExecutionPolicy Bypass -File "<script>" ...`, com a janela do PowerShell oculta pelo próprio .NET (a saída aparece no log do programa).
- O script **só lê** o AD (`Get-ADUser`, `Get-ADGroupMember` e, com os itens adicionais, os `Get-AD*` listados acima), os compartilhamentos e as ACLs, e grava o relatório na pasta de saída. Não altera permissões, contas nem configurações.

### O que o programa não faz
- Não usa empacotadores (PS2EXE, UPX), ofuscação, base64 nem `-EncodedCommand`.
- Não baixa nada nem se conecta à internet.
- Não executa o script dentro do próprio processo, não injeta código e não persiste nada (sem serviço, tarefa agendada ou chave Run).
- Não precisa de instalação nem de exclusões no antivírus.

### Verificação desta versão

| Item | Valor |
|---|---|
| Versão | 1.1.0.0 |
| Data do build | 2026-10-28 |
| `RelatorioAD.exe` | 1.290.240 bytes |
| SHA256 do `RelatorioAD.exe` | `C89F918101D3BB4DEFA39DAA8CBA7A89220BDBBD257D9B7A4E64EE9FFA1558C1` |
| SHA256 do `RelatorioAD.zip` | `6A18CB8C578D505AF592EB4267DAA544DF8ED1C3FE0BE9C19ABE7DAF6AF8B07A` |
| Microsoft Defender | sem ameaças (assinaturas 1.463.8.0), no `dist\` (exe, config e zip) e nos 7 scripts embutidos |
| Assinatura digital | não assinado (`NotSigned`) |

> O csc grava um identificador novo a cada compilação, então **cada build gera um SHA256 diferente**. Os valores acima valem só para este pacote. Num build novo, use os que o `build.ps1` imprime no final e atualize esta tabela.

Para conferir no servidor:
```powershell
Get-FileHash .\RelatorioAD.exe -Algorithm SHA256
```

### Se algum antivírus acusar (falso positivo)
1. Não crie exclusão no antivírus do cliente sem antes conferir o SHA256 acima.
2. Envie o arquivo para análise do fabricante, informando que é uma ferramenta administrativa interna e que é um falso positivo:
   - **Microsoft Defender**: https://www.microsoft.com/wdsi/filesubmission (escolher "Software developer" se tiver conta, ou "Home customer"/"Enterprise customer", e marcar "Incorrectly detected as malware/malicious").
   - **Outros fabricantes**: todos têm um portal de envio de falsos positivos (procure por "<fabricante> false positive submission"). Kaspersky, ESET, Sophos, Trend Micro, Bitdefender e Avast/AVG também aceitam o envio por e-mail ou formulário.
3. Enquanto a análise não sai: o botão **Exportar script...** (de outro computador) ou a pasta `<Pasta de saída>\scripts` dão o `.ps1` da versão, que pode ser executado direto como administrador:
   ```powershell
   powershell.exe -ExecutionPolicy Bypass -File .\Relatorio-2022.ps1 -Saida C:\Relatorio_AD -Abrir
   ```
4. A solução definitiva é assinar o programa (veja abaixo): binários assinados com reputação raramente são bloqueados.

## Certificado (assinatura de código)

O `build.ps1` assina os 7 scripts (antes de embuti-los) e o `.exe`, com carimbo de tempo da DigiCert, quando recebe a impressão digital de um certificado de assinatura de código instalado em `CurrentUser\My` ou `LocalMachine\My`:

```powershell
.\build.ps1 -CertThumbprint 0123456789ABCDEF0123456789ABCDEF01234567
Get-AuthenticodeSignature .\dist\RelatorioAD.exe   # deve mostrar Valid
```

Sem `-CertThumbprint` a etapa é pulada (é o caso desta versão).

### Opção 1 — certificado comercial (recomendado para distribuir a clientes)
Compre um certificado de **assinatura de código** (OV ou EV) de uma autoridade (DigiCert, Sectigo, GlobalSign etc.). Os emitidos desde 2023 ficam num token USB ou HSM: instale o driver do token, confirme que o certificado aparece em `Cert:\CurrentUser\My` e rode o build com `-CertThumbprint`. A assinatura vale em qualquer servidor, sem configuração extra.

### Opção 2 — certificado autoassinado + GPO (uso interno, num domínio)
Vale só nos computadores que recebem a GPO.

1. Criar o certificado (num Windows 10/2016 ou mais novo, no PC do build):
   ```powershell
   $c = New-SelfSignedCertificate -Type CodeSigningCert -Subject 'CN=Compumaq RelatorioAD' `
        -CertStoreLocation Cert:\CurrentUser\My -NotAfter (Get-Date).AddYears(5)
   Export-Certificate -Cert $c -FilePath C:\Temp\RelatorioAD-assinatura.cer
   $c.Thumbprint
   ```
   Guarde a chave privada (exporte um `.pfx` com senha, se precisar compilar em outro PC). O `.cer` só tem a parte pública e é o que vai para a GPO.
2. Na GPO aplicada aos servidores: **Configuração do Computador → Políticas → Configurações do Windows → Configurações de Segurança → Políticas de Chave Pública**, importe o `.cer` em:
   - **Autoridades de Certificação Raiz Confiáveis** (o autoassinado é a própria raiz), e
   - **Editores Confiáveis** (para o PowerShell e o Windows não perguntarem se confiam no editor).
3. `gpupdate /force` nos servidores e build com `.\build.ps1 -CertThumbprint <impressão digital>`.
4. Com o certificado em **Editores Confiáveis**, os scripts assinados rodam até com a política `AllSigned` imposta por GPO.

Observação: o autoassinado não dá reputação no SmartScreen nem nos antivírus de terceiros, ao contrário do comercial. Ele serve para a confiança dentro do domínio.
