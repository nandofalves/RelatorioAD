# Roteiro de teste real — RelatorioAD 1.1 (itens adicionais)

Para o técnico, nos servidores dos clientes ou num laboratório. Os testes automáticos do build rodam num Windows 11 sem AD e sem o motor do PowerShell 2.0, então só este roteiro confirma o comportamento com AD de verdade.

**Ordem dos servidores:** 2022 → 2008 R2 → 2012 R2 → 2016 / 2019 / 2025. O 2022 vem primeiro porque é o mais parecido com a máquina do build; o 2008 R2 logo depois porque é o de maior risco (PowerShell 2.0, ADWS antigo).

## Preparação

- Pacote 1.1: `dist\RelatorioAD.zip`. Confira o SHA256 com o do `docs\ENTREGA.md` (`Get-FileHash .\RelatorioAD.zip -Algorithm SHA256`).
- Pacote 1.0, para o passo 1: o `dist\RelatorioAD.zip` do commit `57a418c` no GitHub: https://github.com/nandofalves/RelatorioAD/raw/57a418c/dist/RelatorioAD.zip (SHA256 `DA541685BEAACA74DDFF2415A47BBF2E58BBB99FCC6FFBD21D11A98A19E89F3F`). Extraia as duas versões em pastas separadas (ex.: `C:\RAD10` e `C:\RAD11`).
- Use uma pasta de saída **nova e vazia** para cada versão (ex.: `C:\Relatorio_AD_10` e `C:\Relatorio_AD_11`), com os mesmos demais parâmetros nas duas.
- Para o passo 3, crie antes do passo 2 uma **conta de teste** ativa (ex.: `teste.rad`), de preferência com acesso a uma das pastas do relatório.
- Ao final de cada passo, guarde o log do programa (copiar da janela) se algo sair diferente do esperado.

## Passo 1 — Nenhum item marcado: igual à 1.0

1. Rode a 1.0 e, logo em seguida, a 1.1 com **todos os itens adicionais desmarcados** e os mesmos parâmetros.
2. CSVs: os nomes trazem data/hora, então compare um a um:
   ```bat
   fc /b C:\Relatorio_AD_10\01_*.csv C:\Relatorio_AD_11\01_*.csv
   ```
   Repita para o 02, o 03 e o 99. Esperado: `nenhuma diferença encontrada`.
3. HTML: compare com `fc` (modo texto). Esperado: só as linhas com a data/hora da geração diferem.
4. Na pasta da 1.1, o único arquivo a mais é o `00_Snapshot_<data>.csv` (além da subpasta `scripts\`). Não há 04, 05, 06 nem 07.
5. O log da 1.1 tem a linha nova `Comando: powershell.exe ...`, sem nenhum parâmetro além dos da 1.0. A linha `Concluído. ...` tem o mesmo formato.

> Diferenças aceitáveis: se alguém fizer logon entre as duas execuções, o último logon dessa conta pode mudar. Rode as duas em sequência e, havendo diferença, confira se é só isso.

## Passo 2 — Todos marcados, primeira execução

1. Na pasta de saída da 1.1, apague os `00_Snapshot_*.csv` do passo 1 (ou use outra pasta vazia).
2. Marque os 5 itens, deixe **Comparar com** vazio e rode.
3. Esperado:
   - Seção **Mudanças**: "Nenhum levantamento anterior encontrado nesta pasta...". Não é gravado o 04.
   - Seções Contas privilegiadas, Política de senha, GPOs e Computadores preenchidas, e os arquivos `05_Contas_Privilegiadas`, `06_GPOs`, `07_Computadores` e `00_Snapshot` na pasta.
   - Nenhum `[AVISO]` no log num DC do domínio raiz com acesso normal.
   - Conferir por amostragem: os membros de Domain Admins batem com o "Usuários e Computadores do AD"; a política de senha bate com `Get-ADDefaultDomainPasswordPolicy`; a quantidade de GPOs bate com o Gerenciamento de Política de Grupo; o fim de suporte de 2 ou 3 computadores bate com a tabela do `ENTREGA.md`.
   - Seção GPOs, bloco "O que cada GPO faz": escolha 2 ou 3 GPOs bem diferentes (uma de segurança/senha, uma de mapeamento de unidades ou impressoras, uma com Modelos administrativos) e compare com o relatório do Gerenciamento de Política de Grupo (botão direito na GPO → Salvar relatório). As áreas e os itens principais têm que bater. Os Modelos administrativos aparecem pela chave do registro, não pelo nome da política. No 2008 R2, confira se os acentos dos caminhos e nomes saem certos. Se aparecer o card **"GPO com senha exposta no SYSVOL"**, confirme o arquivo indicado e avise o cliente: é um achado real de segurança.
   - Seção Computadores: a lista "Sem logon há mais de 7 meses (candidatos à remoção)" traz só máquinas com último logon (ou criação, se nunca logaram) há mais de 7 meses, do mais antigo para o mais recente, com a situação ativo/desativado. Nenhuma delas aparece em "Fora de suporte" nem em "Sem uso há mais de N dias". Conferir 2 ou 3 com o `LastLogonDate` no AD (`Get-ADComputer <nome> -Properties LastLogonDate`). No 07, filtrar `CandidatoRemocao = True` dá a mesma lista.
   - Abrir o HTML: navegação com os itens 6 em diante, cards em "Pontos de atenção", acentos corretos.

## Passo 3 — Mudanças detectadas na segunda execução

1. Desative a conta de teste.
2. Crie uma GPO **vazia** (sem editar nenhuma configuração) e **não vincule** a lugar nenhum.
3. Rode de novo com os 5 itens marcados e **Comparar com** vazio.
4. Esperado:
   - Seção **Mudanças**, categoria Contas: a conta de teste como `desativada`.
   - Categoria GPOs: a GPO nova como `nova`.
   - Seção **GPOs**: a mesma GPO com os achados "Sem nenhum vínculo" e "Vazia (versão 0)". No bloco dela, em "O que cada GPO faz": "Nenhuma configuração encontrada: a GPO não faz nada."
   - Arquivo `04_Mudancas_<data>.csv` com essas linhas.
5. Repita com **Comparar com** apontando para o `00_Snapshot` do passo 2: mesmo resultado.
6. Limpeza: apague a GPO de teste e a conta de teste (o programa não altera nada no AD; quem desfaz é o técnico).

## Passo 4 — Servidor membro num domínio filho (se houver)

1. Num servidor membro (não DC) de um domínio filho, com o RSAT-AD-PowerShell, marque **Contas privilegiadas e de risco** e rode.
2. Esperado: se o domínio raiz não puder ser consultado, o log e a seção mostram o aviso "Enterprise/Schema Admins não puderam ser lidos no domínio raiz ..." e o restante da seção e do relatório sai normalmente. Se o domínio raiz responder, Enterprise Admins e Schema Admins aparecem com os membros do raiz.

## Passo 5 — Sem acesso ao SYSVOL

**Só em laboratório** (não altere o SYSVOL de cliente). Uma forma: no DC de laboratório, negar leitura em `\\<domínio>\SYSVOL\<domínio>\Policies` para a conta usada no teste; ou rodar num servidor membro que não resolva o nome DNS do domínio.

1. Marque **GPOs** e rode.
2. Esperado: na seção GPOs, um único aviso "a pasta ... não pôde ser acessada; as pastas das GPOs no SYSVOL não foram verificadas"; **nenhuma** GPO com o achado "Pasta ausente no SYSVOL"; a coluna `PastaSYSVOL` do 06 com `não verificada`.
3. Desfaça a mudança de permissão.

Sem laboratório, registre o passo como "não testado".

## Resultado

| Servidor | Passo 1 | Passo 2 | Passo 3 | Passo 4 | Passo 5 | Observações |
|---|---|---|---|---|---|---|
| 2022 | | | | | | |
| 2008 R2 | | | | | | |
| 2012 R2 | | | | | | |
| 2016 | | | | | | |
| 2019 | | | | | | |
| 2025 | | | | | | |

Se algum passo falhar, envie o log do programa, o HTML e os CSVs daquela execução (sem dados sensíveis do cliente, se possível) para correção.
