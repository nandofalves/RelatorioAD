# RelatorioAD

[![Baixar RelatorioAD 1.1](https://img.shields.io/badge/BAIXAR-RelatorioAD%201.1-2ea44f?style=for-the-badge&logo=windows&logoColor=white)](https://github.com/nandofalves/RelatorioAD/raw/main/dist/RelatorioAD.zip)

**Versão 1.1** (28/10/2026) · SHA256 do zip: `6A18CB8C578D505AF592EB4267DAA544DF8ED1C3FE0BE9C19ABE7DAF6AF8B07A`

O zip tem o `RelatorioAD.exe` e o `RelatorioAD.exe.config`: extraia os dois na mesma pasta do servidor e abra o `.exe` (veja [Como usar](#como-usar)).

**Novidades da 1.1:** cinco itens adicionais e opcionais de auditoria do domínio. São eles: comparação com o levantamento anterior, contas privilegiadas e de risco, política de senha, GPOs (com o que cada uma faz) e computadores do domínio (fora de suporte e candidatos à remoção). Veja [Itens adicionais](#itens-adicionais-opcionais).

Programa para Windows Server que gera, em poucos cliques, um **relatório em HTML dos usuários ativos do Active Directory e de quem acessa cada pasta compartilhada da rede**, com as permissões de compartilhamento e NTFS já cruzadas por usuário.

Roda do **Windows Server 2008 R2 ao 2025** com um único `.exe`, sem instalação e sem dependências.

## O que o relatório mostra

O resultado é um arquivo HTML, pronto para apresentar ao cliente, com:

- **Resumo:** número de contas ativas no AD, pastas compartilhadas analisadas, contas com privilégio de administrador no servidor e contas ativas sem logon há mais de N dias.
- **Pontos de atenção**, classificados em alta, média e baixa prioridade:
  - contas que são administradoras do servidor (com destaque para contas de terceiros);
  - compartilhamentos que dão controle total para *Todos*;
  - quem acessa as pastas sensíveis (RH, financeiro, diretoria, jurídico, folha...);
  - permissões dadas diretamente a usuários, em vez de a grupos;
  - contas habilitadas sem uso;
  - contas com senha que nunca expira;
  - possíveis contas duplicadas (mesmo nome de exibição);
  - pastas abertas para todos os usuários.
- **Matriz de acesso:** usuário × pasta, agrupada por setor (OU), com o nível efetivo (leitura, modificar ou controle total). O nível efetivo combina a permissão do compartilhamento com a do NTFS, expande os grupos de forma recursiva e considera as permissões de negação.
- **Detalhe por pasta:** permissões do compartilhamento, permissões NTFS e as subpastas com permissão própria (até a profundidade escolhida).
- **Detalhe por usuário:** setor, grupos, último logon e pastas a que tem acesso.
- **Como o levantamento foi feito:** servidor, versão do Windows, data e critérios usados.

Além do HTML, são gravados CSVs para análise em planilha:

| Arquivo | Conteúdo |
|---|---|
| `01_Usuarios_Ativos_<data>.csv` | contas ativas do AD com setor, cargo, e-mail, último logon e grupos |
| `02_Permissoes_Compartilhamento_<data>.csv` | permissões de cada compartilhamento |
| `03_Permissoes_NTFS_por_Usuario_<data>.csv` | permissões NTFS usuário a usuário, com a origem (direta ou por grupo) |
| `99_Erros_Acesso_<data>.csv` | pastas que não puderam ser lidas (só aparece se houver) |
| `00_Snapshot_<data>.csv` | o "levantamento" desta execução, usado na comparação da próxima visita (gravado sempre) |

O programa **só lê** o AD, os compartilhamentos e as permissões. Ele não altera contas, grupos, permissões nem configurações do servidor.

## Itens adicionais (opcionais)

Ficam no grupo **Itens adicionais (opcionais)** da janela, todos **desmarcados** por padrão. Desmarcados, o relatório sai igual ao da 1.0. Cada item marcado acrescenta uma seção numerada no HTML, cards em "Pontos de atenção" e, em alguns casos, um CSV. Se um item falhar (sem acesso, domínio raiz fora do ar...), o log mostra `[AVISO]` e o restante do relatório sai normalmente.

| Item | O que mostra | Arquivo |
|---|---|---|
| **Comparar com o levantamento anterior** | O que mudou desde a última execução no mesmo domínio: contas novas, desativadas e excluídas; quem entrou ou saiu dos administradores do servidor; pastas novas e removidas; acessos ganhos, perdidos ou alterados; e, se o item estava marcado nas duas execuções, grupos privilegiados, GPOs e computadores. | `04_Mudancas_<data>.csv` |
| **Contas privilegiadas e de risco** | Membros (diretos e indiretos) dos grupos administrativos do domínio; contas com senha não exigida, pré-autenticação Kerberos desativada ou criptografia reversível; delegação irrestrita; terceiros com privilégio; privilegiadas sem uso; contas de serviço com senha antiga; Administrador interno com o nome padrão; contas que nunca fizeram logon. | `05_Contas_Privilegiadas_<data>.csv` |
| **Política de senha** | Política padrão do domínio (tamanho, complexidade, histórico, bloqueio, validade...) comparada com uma referência, e as políticas refinadas (FGPP). | só no HTML |
| **GPOs** | Explicação de como as GPOs funcionam e, para cada GPO, onde está vinculada e **o que ela faz**: política de senha, auditoria, direitos de usuário, opções de segurança, Modelos administrativos por assunto (Windows Update, Firewall, Área de Trabalho Remota...), scripts, mapeamento de unidades, impressoras, grupos locais, tarefas agendadas, instalação de software etc. Achados: **senha gravada em Preferências (MS14-025)**, vínculos para GPOs inexistentes, GPOs sem pasta no SYSVOL, sem vínculo, desativadas ou vazias, herança bloqueada. Lido direto do AD e do SYSVOL, sem o console de GPO. | `06_GPOs_<data>.csv`, `06_GPOs_Configuracoes_<data>.csv` |
| **Computadores do domínio** | Sistema de cada computador × data de fim de suporte da Microsoft; lista de **candidatos à remoção** (sem logon há mais de 7 meses, ativos ou desativados); ativos sem uso recente; contas desativadas. | `07_Computadores_<data>.csv` |

## Como usar

### Requisitos no servidor
- Windows Server 2008 R2, 2012, 2012 R2, 2016, 2019, 2022 ou 2025.
- Executar como **administrador** (o programa pede a elevação sozinho).
- **Módulo Active Directory do PowerShell**. Num controlador de domínio ele já vem instalado. Num servidor de arquivos membro:
  ```powershell
  Install-WindowsFeature RSAT-AD-PowerShell                            # 2012 em diante
  Import-Module ServerManager; Add-WindowsFeature RSAT-AD-PowerShell   # 2008 R2
  ```

### Passo a passo
1. Copie o `RelatorioAD.zip` para o servidor. Se ele veio da internet ou por e-mail, antes de extrair: botão direito → **Propriedades** → marque **Desbloquear**.
2. Extraia os dois arquivos, **`RelatorioAD.exe` e `RelatorioAD.exe.config`, na mesma pasta**. O `.config` é o que permite rodar no 2012 em diante sem instalar o .NET 3.5.
3. Abra o `RelatorioAD.exe` e aceite o UAC.
4. Confira a linha **Detectado** (versão do Windows, PowerShell, módulo AD e administrador). A versão do servidor já vem selecionada. Se aparecer um aviso amarelo, leia antes de continuar.
5. Ajuste os parâmetros, se quiser (todos têm valor padrão):

   | Campo | Para que serve | Padrão |
   |---|---|---|
   | Cliente | nome que aparece no relatório | nome NetBIOS do domínio |
   | Responsável | quem fez o levantamento | Compumaq · Werner Informática |
   | Profundidade | quantos níveis de subpastas analisar abaixo de cada compartilhamento | 2 |
   | Pasta de saída | onde gravar o HTML, os CSVs e o script executado | `C:\Relatorio_AD` |
   | Caminhos | pastas específicas, uma por linha, em vez dos compartilhamentos | vazio = todos os compartilhamentos |
   | Dias inativo | a partir de quantos dias sem logon a conta é considerada sem uso | 90 |
   | Padrão de terceiros | expressão regular aplicada à OU para marcar contas de terceiros | `Terceir\|Tercer\|Fornecedor\|Externo` |
   | Pastas sensíveis | expressão regular aplicada ao nome do compartilhamento | `^(rh\|pessoal\|dp\|financeiro\|...)` |

6. Marque os **itens adicionais** que quiser (todos opcionais):

   | Campo | Para que serve | Padrão |
   |---|---|---|
   | Comparar com o levantamento anterior | liga a seção Mudanças | desmarcado |
   | Comparar com | um `00_Snapshot_*.csv` específico (botão **...**) | vazio = o mais recente do mesmo domínio na pasta de saída |
   | Contas privilegiadas e de risco | liga a seção Contas privilegiadas | desmarcado |
   | Contas de serviço | expressão regular (login ou OU) que identifica contas de serviço, além das que têm SPN | `svc\|servico\|serviço\|service\|sql\|backup` |
   | Dias senha serviço | idade máxima da senha de uma conta de serviço antes do alerta | 365 |
   | Política de senha / GPOs / Computadores do domínio | ligam as seções correspondentes | desmarcados |

   Os campos **Comparar com**, **Contas de serviço** e **Dias senha serviço** só ficam habilitados com o item correspondente marcado.

7. Clique em **Gerar relatório**. O log mostra o andamento ao vivo, nas etapas de 1/5 a 5/5 (os itens adicionais aparecem como linhas `·` abaixo da etapa 5). Ao terminar, o relatório abre no navegador (se a opção estiver marcada) e os botões **Abrir relatório** e **Abrir pasta** ficam disponíveis.

### Comparar uma visita com a anterior
1. Na primeira visita, gere o relatório normalmente. O `00_Snapshot_<data>.csv` fica na pasta de saída. **Não apague esse arquivo.**
2. Na visita seguinte, use a **mesma pasta de saída**, marque **Comparar com o levantamento anterior** e os mesmos itens adicionais da primeira vez (grupos privilegiados, GPOs e computadores só são comparados se estavam marcados nas duas execuções).
3. A seção **Mudanças** e o `04_Mudancas_<data>.csv` listam o que mudou. Se o levantamento anterior estiver em outra pasta, aponte o arquivo no campo **Comparar com**.

### Remover computadores antigos
Na seção **Computadores**, a lista "Sem logon há mais de 7 meses (candidatos à remoção)" já vem do mais antigo para o mais recente. No `07_Computadores`, filtre `CandidatoRemocao = True` para ter a lista completa. Sugestão: confirmar com o cliente, desativar, mover para uma OU de desativados e excluir depois de algumas semanas sem reclamação.

### Problemas comuns
| Mensagem | O que fazer |
|---|---|
| Módulo ActiveDirectory ausente | Instalar o RSAT-AD-PowerShell (veja os requisitos). |
| Nenhum compartilhamento encontrado | Informar as pastas no campo **Caminhos**. |
| Erro de política de execução (`AuthorizationManager` / `execution policy`) | Uma GPO está forçando a política de execução do PowerShell. Peça uma exceção ou use a versão assinada (veja `docs/ENTREGA.md`). |
| `[AVISO] Módulo ...` no log | Aquele item adicional não pôde ser lido por inteiro; a seção dele diz o motivo e o resto do relatório não é afetado. |
| "Nenhum levantamento anterior encontrado" | É a primeira execução com essa pasta de saída, ou os `00_Snapshot` foram apagados. A próxima execução já terá com o que comparar. |
| Card "GPO com senha exposta no SYSVOL" | Achado real de segurança: trocar a senha da conta indicada e remover a senha da GPO. |

### Executar o script sem o programa
O botão **Exportar script...** salva o `.ps1` da versão escolhida. O script executado também fica em `<Pasta de saída>\scripts`. Ele pode ser lido, auditado e rodado direto, como administrador:
```powershell
powershell.exe -ExecutionPolicy Bypass -File .\Relatorio-2022.ps1 -Saida C:\Relatorio_AD -Abrir
```

## Como funciona

O `.exe` é um programa WinForms em C# que traz embutidos 7 scripts PowerShell, um por versão do Windows Server. Ao gerar o relatório, ele grava o script da versão escolhida na pasta de saída e o executa com `powershell.exe -File`, mostrando a saída na janela.

Os itens adicionais ficam em `scripts/Modulos.ps1`, escrito uma vez e compatível com o PowerShell 2.0. O build embute esse arquivo nas 3 bases, e cada script gerado continua sendo um arquivo só.

Os 7 scripts saem de 3 bases, conforme o PowerShell nativo de cada versão:

| Windows Server | PowerShell | Base | Diferenças |
|---|---|---|---|
| 2016, 2019, 2022, 2025 | 5.1 | `scripts/Relatorio-PS51.ps1` | script de referência |
| 2012, 2012 R2 | 3.0 / 4.0 | `scripts/Relatorio-PS3.ps1` | varredura de subpastas sem `-Depth` |
| 2008 R2 | 2.0 | `scripts/Relatorio-PS2.ps1` | compartilhamentos via WMI e só construções do PS 2.0 |

Medidas contra falso positivo de antivírus: sem empacotadores (PS2EXE, UPX), sem ofuscação, sem `-EncodedCommand` e sem download. O script executado fica legível em disco. O build escaneia o resultado com o Microsoft Defender.

## Compilar

Precisa só do Windows com o .NET Framework 3.5 (usa o `csc.exe` do próprio Windows; não precisa de Visual Studio). No PowerShell 5.1, como administrador:

```powershell
.\build.ps1                          # gera os scripts, valida, compila, empacota e escaneia
.\build.ps1 -SemScan                 # sem o scan do Defender
.\build.ps1 -CertThumbprint <x>      # assina scripts e exe com um certificado de assinatura de código
```

O resultado vai para `dist\`: `RelatorioAD.exe`, `RelatorioAD.exe.config` e `RelatorioAD.zip`. O `build.ps1` também roda o `tools/Test-ScriptCompat.ps1`, que verifica se cada script usa só o que existe na versão de PowerShell do seu servidor, e o `tools/Test-Modulos.ps1`, com os testes unitários dos itens adicionais.

## Estrutura

```
build.ps1                    build completo
scripts/                     as 3 bases dos scripts e o Modulos.ps1 (itens adicionais)
referencia/                  script original (2022), a fonte da verdade do relatório
src/                         aplicativo C# (WinForms)
tools/Test-ScriptCompat.ps1  lint de compatibilidade por versão de PowerShell
tools/Test-Modulos.ps1       testes unitários do Modulos.ps1 (sem AD)
docs/                        PRD, Spec, progresso, guia de entrega (ENTREGA.md) e features/
```

## Situação

Versão 1.1, testada nos servidores reais com Windows Server 2008 R2, 2012 e 2022. Roteiro de teste em `docs/features/001-auditoria-dominio/TESTE.md`; histórico em `docs/progress.md`.
