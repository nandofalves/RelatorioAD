# RelatorioAD

[![Baixar RelatorioAD.zip](https://img.shields.io/badge/BAIXAR-RelatorioAD.zip-2ea44f?style=for-the-badge&logo=windows&logoColor=white)](https://github.com/nandofalves/RelatorioAD/raw/main/dist/RelatorioAD.zip)

O zip tem o `RelatorioAD.exe` e o `RelatorioAD.exe.config`: extraia os dois na mesma pasta do servidor e abra o `.exe` (veja [Como usar](#como-usar)).

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

O programa **só lê** o AD, os compartilhamentos e as permissões. Ele não altera contas, grupos, permissões nem configurações do servidor.

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

6. Clique em **Gerar relatório**. O log mostra o andamento ao vivo, nas etapas de 1/5 a 5/5. Ao terminar, o relatório abre no navegador (se a opção estiver marcada) e os botões **Abrir relatório** e **Abrir pasta** ficam disponíveis.

### Problemas comuns
| Mensagem | O que fazer |
|---|---|
| Módulo ActiveDirectory ausente | Instalar o RSAT-AD-PowerShell (veja os requisitos). |
| Nenhum compartilhamento encontrado | Informar as pastas no campo **Caminhos**. |
| Erro de política de execução (`AuthorizationManager` / `execution policy`) | Uma GPO está forçando a política de execução do PowerShell. Peça uma exceção ou use a versão assinada (veja `docs/ENTREGA.md`). |

### Executar o script sem o programa
O botão **Exportar script...** salva o `.ps1` da versão escolhida. O script executado também fica em `<Pasta de saída>\scripts`. Ele pode ser lido, auditado e rodado direto, como administrador:
```powershell
powershell.exe -ExecutionPolicy Bypass -File .\Relatorio-2022.ps1 -Saida C:\Relatorio_AD -Abrir
```

## Como funciona

O `.exe` é um programa WinForms em C# que traz embutidos 7 scripts PowerShell, um por versão do Windows Server. Ao gerar o relatório, ele grava o script da versão escolhida na pasta de saída e o executa com `powershell.exe -File`, mostrando a saída na janela.

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

O resultado vai para `dist\`: `RelatorioAD.exe`, `RelatorioAD.exe.config` e `RelatorioAD.zip`. O `build.ps1` também roda o `tools/Test-ScriptCompat.ps1`, que verifica se cada script usa só o que existe na versão de PowerShell do seu servidor.

## Estrutura

```
build.ps1                    build completo
scripts/                     as 3 bases dos scripts
referencia/                  script original (2022), a fonte da verdade do relatório
src/                         aplicativo C# (WinForms)
tools/Test-ScriptCompat.ps1  lint de compatibilidade por versão de PowerShell
docs/                        PRD, Spec, progresso e guia de entrega (ENTREGA.md)
```

## Situação

Versão 1.0, em validação nos servidores reais. O primeiro teste está sendo feito no 2008 R2 com AD; as etapas 1 a 3 já rodaram. Veja `docs/progress.md`.
