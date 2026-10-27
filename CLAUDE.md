# CLAUDE.md

## Projeto
RelatorioAD: programa Windows (.exe com interface gráfica) que gera o relatório de usuários ativos do AD e das permissões de compartilhamento/NTFS em HTML. O técnico escolhe a versão do Windows Server e o programa executa o script PowerShell embutido específico daquela versão.

## Stack
- C# compilado com `csc.exe` do .NET Framework v3.5 (CLR 2), com `app.config` declarando `supportedRuntime` v4.0 e v2.0.50727 → um único .exe roda do 2008 R2 ao 2025 sem instalar nada.
- WinForms (sem WPF, sem bibliotecas externas).
- Scripts PowerShell embutidos como recurso (`/resource:`), um por versão de SO.
- Roda localmente no servidor (DC ou servidor de arquivos membro), como administrador.

## Fluxo de trabalho
- Este projeto usa Spec-Driven Development: `docs/PRD.md` → `docs/Spec.md` → implementação por fases.
- Implementar uma fase por vez e registrar em `docs/progress.md`.
- Commits em Conventional Commits.

## Padrões
- Textos da interface e dos relatórios em português (pt-BR).
- Scripts `.ps1` salvos em UTF-8 **com BOM** (o PowerShell 2.0–5.1 lê UTF-8 sem BOM como ANSI e quebra os acentos).
- Cada script deve ser compatível com a versão de PowerShell nativa do seu SO (2008 R2 = 2.0, 2012 = 3.0, 2012 R2 = 4.0, 2016+ = 5.1).
- Script de referência (versão 2022): `referencia/Relatorio-AD-Permissoes-2022.ps1`.

## Restrições
- Não usar PS2EXE, empacotadores, ofuscação nem `-EncodedCommand` (disparam antivírus).
- Não adicionar libs sem registrar na Spec.
