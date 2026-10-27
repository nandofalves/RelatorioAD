# RelatorioAD 1.0 — Entrega

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

## Antivírus

### O que o programa faz
- Mostra uma janela, lê o registro (versão do Windows e do PowerShell) e verifica se o módulo AD existe.
- Grava em `<Pasta de saída>\scripts\Relatorio-<versão>.ps1` o script PowerShell daquela versão. Esse arquivo é **texto legível**, idêntico ao que o botão **Exportar script...** salva, e pode ser conferido antes ou depois da execução.
- Executa esse arquivo com `powershell.exe -NoProfile -NonInteractive -ExecutionPolicy Bypass -File "<script>" ...`, com a janela do PowerShell oculta pelo próprio .NET (a saída aparece no log do programa).
- O script **só lê** o AD (`Get-ADUser`, `Get-ADGroupMember`), os compartilhamentos e as ACLs, e grava o relatório na pasta de saída. Não altera permissões, contas nem configurações.

### O que o programa não faz
- Não usa empacotadores (PS2EXE, UPX), ofuscação, base64 nem `-EncodedCommand`.
- Não baixa nada nem se conecta à internet.
- Não executa o script dentro do próprio processo, não injeta código e não persiste nada (sem serviço, tarefa agendada ou chave Run).
- Não precisa de instalação nem de exclusões no antivírus.

### Verificação desta versão

| Item | Valor |
|---|---|
| Data do build | 2026-10-27 |
| `RelatorioAD.exe` | 376.832 bytes |
| SHA256 do `RelatorioAD.exe` | `3E3817C83A896195E1F1C8E63D67FF856AE4612EA9D9D90342D055A78C0923AF` |
| SHA256 do `RelatorioAD.zip` | `DA541685BEAACA74DDFF2415A47BBF2E58BBB99FCC6FFBD21D11A98A19E89F3F` |
| Microsoft Defender | sem ameaças (assinaturas 1.461.225.0), no `dist\` (exe, config e zip) e nos 7 scripts embutidos |
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
