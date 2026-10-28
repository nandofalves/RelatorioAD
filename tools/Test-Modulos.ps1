#Requires -Version 5.1
<#
.SYNOPSIS
  Testes unitários de scripts\Modulos.ps1 (camadas puras e ganchos de HTML), sem AD.
  Imprime OK/FALHOU por caso; exit code 1 se algum falhar. Roda no build (etapa 2b).

.EXEMPLO
  .\tools\Test-Modulos.ps1
#>
$ErrorActionPreference = 'Stop'
. (Join-Path $PSScriptRoot '..\scripts\Modulos.ps1')

# Stubs iguais aos da base PS51
function HE([object]$s) { [System.Net.WebUtility]::HtmlEncode([string]$s) }
function DT($d) { if ($d) { ([datetime]$d).ToString('dd/MM/yyyy') } else { 'nunca' } }
function Card($sev, $titulo, $corpo) {
    $rot = @{ alta = 'Prioridade alta'; media = 'Prioridade média'; baixa = 'Para validar' }[$sev]
    "<div class=""card sev-$sev""><div class=""sev"">$rot</div><h3>$titulo</h3>$corpo</div>"
}

# ---- mini framework ----
$script:falhas = 0
$script:total = 0
function Igual($esperado, $obtido, [string]$msg) {
    if ([string]$esperado -cne [string]$obtido) { throw "${msg}: esperado [$esperado], obtido [$obtido]" }
}
function Teste([string]$nome, [scriptblock]$b) {
    $script:total++
    try { & $b; Write-Host "  OK      $nome" -ForegroundColor Green }
    catch { $script:falhas++; Write-Host "  FALHOU  $nome — $($_.Exception.Message)" -ForegroundColor Red }
}

Write-Host 'Testes de scripts\Modulos.ps1'

# ---- fixtures ----
function Obj([hashtable]$h) { New-Object PSObject -Property $h }
# Variáveis da base usadas por Get-LinhasSnapshot / Invoke-Modulos (dot-source dentro de cada teste)
$amb = {
    $csvOpts = @{ NoTypeInformation = $true; Encoding = 'UTF8'; Delimiter = ';' }
    $dom = Obj @{ DNSRoot = 'empresa.local' }
    $hoje = [datetime]'2026-10-27 14:05'
    $stamp = $hoje.ToString('yyyyMMdd_HHmm')
    $Caminhos = $null
    $PastasSensiveis = '^(rh|financeiro)'
    $U = @{ 'jose' = (Obj @{ Nome = 'José Conceição' }); 'ana' = (Obj @{ Nome = 'Ana Lúcia' }) }
    $ADM = @('ana')
    $alvos = @((Obj @{ Nome = 'Financeiro'; Path = 'D:\Dados\Financeiro' }), (Obj @{ Nome = 'Ação'; Path = 'D:\Dados\Ação' }))
    $Calc = @{ 'Financeiro' = (Obj @{ EF = @{ 'jose' = 2; 'ana' = 3 } }); 'Ação' = (Obj @{ EF = @{ 'ana' = 3 } }) }
    $Comparar = $null; $CompararCom = $null; $ContasDominio = $null; $PoliticaSenha = $null; $GPOs = $null; $Computadores = $null
}
function NovoDir { $d = Join-Path ([IO.Path]::GetTempPath()) ('relad_' + [guid]::NewGuid().ToString('N')); New-Item -ItemType Directory -Path $d | Out-Null; $d }
# Snapshot a partir de linhas 'Tipo;Chave;Valor'
function Snap([string[]]$l) { ConvertTo-Snapshot @($l | ForEach-Object { $p = $_ -split ';', 3; NovaLinhaSnap $p[0] $p[1] $p[2] }) }
$metaBase = @('_meta;Formato;1', '_meta;Dominio;empresa.local', '_meta;Escopo;compartilhamentos')
function Mud($r, [string]$cat) { @(@($r.Mudancas) | Where-Object { $_.Categoria -eq $cat } | ForEach-Object { "$($_.Mudanca):$($_.Item):$($_.Antes)>$($_.Depois)" }) -join ' / ' }

# ---- Fase 1: ganchos ----
Teste 'ganchos vazios sem módulo marcado' {
    $MOD = @{ Marcados = @() }
    Igual '' (NavModulos $MOD) 'NavModulos'
    Igual '' (SecoesModulos $MOD) 'SecoesModulos'
    Igual '' (CssModulos $MOD) 'CssModulos'
    Igual '' (MetodoModulos $MOD) 'MetodoModulos'
    Igual 0 @(CardsModulos $MOD).Count 'CardsModulos'
}
Teste 'navegação numerada a partir de 6' {
    $MOD = @{ Marcados = @('ContasDominio', 'GPOs') }
    $nav = NavModulos $MOD
    Igual '<a href="#mod-contas">6. Contas privilegiadas</a><a href="#mod-gpos">7. GPOs</a>' $nav 'NavModulos'
    Igual 7 (NumeroSecao $MOD 'GPOs') 'NumeroSecao'
}
Teste 'Invoke-Modulos monta Marcados na ordem fixa' {
    . $amb
    $Saida = NovoDir
    try {
        $GPOs = [switch]$true; $Computadores = [switch]$true; $Comparar = [switch]$true
        $MOD = Invoke-Modulos 6>$null
        Igual 'Comparar,GPOs,Computadores' ($MOD.Marcados -join ',') 'Marcados'
        $GPOs = $null; $Computadores = $null; $Comparar = [switch]$false
        Igual 0 @((Invoke-Modulos).Marcados).Count 'nenhum marcado'
    } finally { Remove-Item -LiteralPath $Saida -Recurse -Force }
}
Teste 'módulo com falha vira seção de erro e não gera card' {
    $MOD = @{ Marcados = @('PoliticaSenha'); Senha = (New-Object PSObject -Property @{ Ok = $false; Erro = 'acesso <negado>'; Avisos = @() }) }
    Igual '<section id="mod-senha"><h2><span>6</span>Política de senha</h2><p class="lead">Não foi possível ler: acesso &lt;negado&gt;</p></section>' (SecoesModulos $MOD) 'SecoesModulos'
    Igual 0 @(CardsModulos $MOD).Count 'CardsModulos'
}
Teste 'Invoke-Modulo captura o erro' {
    $r = Invoke-Modulo 'Teste' { throw 'falhou de propósito' } 6>$null
    Igual $false $r.Ok 'Ok'
    Igual 'falhou de propósito' $r.Erro 'Erro'
}

# ---- Fase 2: snapshot e comparação ----


Teste 'snapshot: ida e volta com acentos' {
    . $amb
    $dir = NovoDir
    try {
        $arq = Write-Snapshot @(Get-LinhasSnapshot @{ Marcados = @() }) (Join-Path $dir '00_Snapshot_x.csv')
        Igual (Join-Path $dir '00_Snapshot_x.csv') $arq 'caminho devolvido'
        $s = Read-Snapshot $arq
        Igual '1' $s.Meta['Formato'] 'Formato'
        Igual 'empresa.local' $s.Meta['Dominio'] 'Dominio'
        Igual '2026-10-27 14:05' $s.Meta['Data'] 'Data'
        Igual 'compartilhamentos' $s.Meta['Escopo'] 'Escopo'
        Igual '' $s.Meta['Modulos'] 'Modulos'
        Igual 'José Conceição' $s.Dados['Usuario']['jose'] 'Usuario'
        Igual $true $s.Dados['Admin'].ContainsKey('ana') 'Admin'
        Igual 'D:\Dados\Ação' $s.Dados['Pasta']['Ação'] 'Pasta'
        Igual '2' $s.Dados['Acesso']['Financeiro|jose'] 'Acesso'
        Igual '3' $s.Dados['Acesso']['Ação|ana'] 'Acesso com acento'
        Igual 3 @($s.Dados['Acesso'].Keys).Count 'linhas de acesso'
        Igual 'empresa.local' (Read-Snapshot $arq -SoMeta).Meta['Dominio'] 'SoMeta'
        $Caminhos = @('D:\A', 'E:\B')
        Igual 'D:\A|E:\B' (ConvertTo-Snapshot @(Get-LinhasSnapshot @{ Marcados = @() })).Meta['Escopo'] 'Escopo com -Caminhos'
    } finally { Remove-Item -LiteralPath $dir -Recurse -Force }
}
Teste 'snapshot: falha ao gravar só avisa' {
    . $amb
    $r = Write-Snapshot @(NovaLinhaSnap '_meta' 'Formato' '1') 'Z:\nao\existe\00_Snapshot_x.csv' 6>$null
    Igual $null $r 'retorno'
}
Teste 'Read-Snapshot recusa outro formato' {
    . $amb
    $dir = NovoDir
    try {
        $f2 = Join-Path $dir 'f2.csv'
        Write-Snapshot @((NovaLinhaSnap '_meta' 'Formato' '2'), (NovaLinhaSnap '_meta' 'Dominio' 'empresa.local')) $f2 | Out-Null
        $msg = ''
        try { Read-Snapshot $f2 | Out-Null } catch { $msg = $_.Exception.Message }
        Igual 'arquivo não é um levantamento do RelatorioAD' $msg 'Formato=2'
        $txt = Join-Path $dir 'txt.csv'
        [IO.File]::WriteAllText($txt, "a;b`r`n1;2`r`n")
        $msg = ''
        try { Read-Snapshot $txt | Out-Null } catch { $msg = $_.Exception.Message }
        Igual 'arquivo não é um levantamento do RelatorioAD' $msg 'CSV qualquer'
    } finally { Remove-Item -LiteralPath $dir -Recurse -Force }
}
Teste 'Find-SnapshotAnterior: mais recente do mesmo domínio' {
    . $amb
    $dir = NovoDir
    try {
        $snap = { param($nome, $dominio, $data) Write-Snapshot @((NovaLinhaSnap '_meta' 'Formato' '1'), (NovaLinhaSnap '_meta' 'Dominio' $dominio), (NovaLinhaSnap '_meta' 'Data' $data)) (Join-Path $dir $nome) | Out-Null }
        & $snap '00_Snapshot_20260901_1000.csv' 'empresa.local' '2026-09-01 10:00'
        & $snap '00_Snapshot_20261001_0900.csv' 'EMPRESA.local' '2026-10-01 09:00'
        & $snap '00_Snapshot_20261020_0800.csv' 'outro.local' '2026-10-20 08:00'
        & $snap '00_Snapshot_20261027_1405.csv' 'empresa.local' '2026-10-27 14:05'
        [IO.File]::WriteAllText((Join-Path $dir '00_Snapshot_lixo.csv'), 'lixo')
        $r = Find-SnapshotAnterior $dir 'empresa.local' (Join-Path $dir '00_Snapshot_20261027_1405.csv')
        Igual (Join-Path $dir '00_Snapshot_20261001_0900.csv') $r 'anterior'
        Igual $null (Find-SnapshotAnterior $dir 'nenhum.local' '') 'outro domínio'
        Igual $null (Find-SnapshotAnterior (Join-Path $dir 'nao') 'empresa.local' '') 'pasta inexistente'
    } finally { Remove-Item -LiteralPath $dir -Recurse -Force }
}
Teste 'Compare-Snapshot: contas, administradores e pastas' {
    $a = Snap ($metaBase + @('Usuario;ana;Ana', 'Usuario;bia;Bia', 'Usuario;caio;Caio', 'Admin;ana;', 'Admin;bia;',
        'Pasta;RH;D:\RH', 'Pasta;Velha;D:\Velha'))
    $d = Snap ($metaBase + @('Usuario;ana;Ana', 'Usuario;davi;Davi', 'Usuario;eva;eva', 'Admin;ana;', 'Admin;davi;',
        'Pasta;RH;D:\RH', 'Pasta;Nova;D:\Nova', 'Acesso;RH|davi;1', 'Acesso;Nova|davi;2'))
    $c = Compare-Snapshot $a $d @{ 'bia' = 1 }
    Igual 'desativada:Bia (bia):ativa>desativada / excluída:Caio (caio):ativa>excluída / nova:Davi (davi):>acesso a 2 pastas / nova:eva:>acesso a 0 pastas' (Mud $c 'Contas') 'Contas'
    Igual 'saiu:Bia (bia):> / entrou:Davi (davi):>' (Mud $c 'Administradores do servidor') 'Admin'
    Igual 'nova:Nova:>D:\Nova / removida:Velha:D:\Velha>' (Mud $c 'Pastas') 'Pastas'
    Igual '' (Mud $c 'Acesso') 'conta nova não gera Acesso'
    Igual $false $c.EscopoMudou 'EscopoMudou'
}
Teste 'Compare-Snapshot: escopo mudou não lista pastas' {
    $a = Snap ($metaBase + @('Pasta;A;D:\A'))
    $d = Snap (@('_meta;Formato;1', '_meta;Dominio;empresa.local', '_meta;Escopo;D:\B') + @('Pasta;B;D:\B'))
    $c = Compare-Snapshot $a $d @{}
    Igual $true $c.EscopoMudou 'EscopoMudou'
    Igual 0 @($c.Mudancas).Count 'sem mudanças de pasta'
}
Teste 'Compare-Snapshot: acesso só na interseção' {
    $a = Snap ($metaBase + @('Usuario;ana;Ana', 'Usuario;bia;Bia', 'Usuario;caio;Caio', 'Pasta;RH;D:\RH', 'Pasta;Fin;D:\Fin', 'Pasta;Velha;D:\Velha',
        'Acesso;Fin|bia;1', 'Acesso;Fin|caio;3', 'Acesso;Velha|ana;3', 'Acesso;RH|ana;2'))
    $d = Snap ($metaBase + @('Usuario;ana;Ana', 'Usuario;bia;Bia', 'Usuario;caio;Caio', 'Usuario;davi;Davi', 'Pasta;RH;D:\RH', 'Pasta;Fin;D:\Fin', 'Pasta;Nova;D:\Nova',
        'Acesso;RH|bia;1', 'Acesso;Fin|bia;2', 'Acesso;Nova|ana;3', 'Acesso;RH|davi;1', 'Acesso;RH|ana;2'))
    $c = Compare-Snapshot $a $d @{}
    Igual 'mudou:Fin · Bia (bia):Leitura>Modificar / perdeu:Fin · Caio (caio):Controle total>Sem acesso / ganhou:RH · Bia (bia):Sem acesso>Leitura' (Mud $c 'Acesso') 'Acesso'
    Igual 'RH' (@($c.Mudancas) | Where-Object { $_.Mudanca -eq 'ganhou' }).Pasta 'Pasta guardada'
}
Teste 'Compare-Snapshot: módulos só quando rodaram dos dois lados' {
    $a = Snap ($metaBase + @('_meta;Modulos;ContasDominio,GPOs,Computadores', 'Usuario;ana;Ana',
        'Privilegiado;Domain Admins|ana;', 'Privilegiado;Domain Admins|bia;',
        'GPO;{1};Senhas|0|5', 'GPO;{2};Velha|0|1', 'GPO;{3};A|b|0|2',
        'Computador;PC1;1|Windows 10 Pro', 'Computador;PC2;1|Windows 10 Pro', 'Computador;PC3;0|Windows 11 Pro', 'Computador;PC4;1|Windows 7'))
    $d = Snap ($metaBase + @('_meta;Modulos;ContasDominio,GPOs,Computadores', 'Usuario;ana;Ana', 'Usuario;davi;Davi',
        'Privilegiado;Domain Admins|ana;', 'Privilegiado;Enterprise Admins|davi;',
        'GPO;{1};Senhas|0|6', 'GPO;{3};A|b|3|2', 'GPO;{4};Nova|0|0',
        'Computador;PC1;1|Windows 11 Pro', 'Computador;PC2;0|Windows 10 Pro', 'Computador;PC3;1|Windows 11 Pro', 'Computador;PC5;1|Windows 11 Pro'))
    $c = Compare-Snapshot $a $d @{}
    Igual 'saiu:Domain Admins · bia:> / entrou:Enterprise Admins · Davi (davi):>' (Mud $c 'Grupos privilegiados') 'Privilegiado'
    Igual 'alterada:Senhas:versão 5>versão 6 / removida:Velha:ativa> / status:A|b:ativa>toda desativada / nova:Nova:>ativa' (Mud $c 'GPOs') 'GPOs'
    Igual 'SO mudou:PC1:Windows 10 Pro>Windows 11 Pro / desativado:PC2:ativo>desativado / reativado:PC3:desativado>ativo / removido:PC4:Windows 7> / novo:PC5:>Windows 11 Pro' (Mud $c 'Computadores') 'Computadores'
    Igual 0 @($c.NaoComparado).Count 'NaoComparado'

    $a2 = Snap ($metaBase + @('_meta;Modulos;ContasDominio', 'Privilegiado;Domain Admins|ana;'))
    $d2 = Snap ($metaBase + @('_meta;Modulos;GPOs', 'GPO;{4};Nova|0|0'))
    $c2 = Compare-Snapshot $a2 $d2 @{}
    Igual 0 @($c2.Mudancas).Count 'nada comparado'
    Igual 'este levantamento não incluiu os grupos privilegiados | o levantamento anterior não incluiu as GPOs' (@($c2.NaoComparado) -join ' | ') 'NaoComparado'
}
Teste 'CardsMudancas e SecaoMudancas' {
    . $amb
    $base = @{ Ok = $true; Erro = $null; Avisos = @(); Arquivo = 'C:\x\00_Snapshot_20261001_0900.csv'; DataAnterior = [datetime]'2026-10-01 09:00'
        OutroDominio = $null; Mudancas = @(); NaoComparado = @(); EscopoMudou = $false; Csv = $null }
    $semAnt = Obj ($base.Clone()); $semAnt.DataAnterior = $null
    Igual 0 @(CardsMudancas $semAnt).Count 'sem anterior: sem card'
    Igual $true ((SecaoMudancas $semAnt 6) -like '*Nenhum levantamento anterior encontrado nesta pasta.*') 'sem anterior: texto'
    $outro = Obj ($base.Clone()); $outro.OutroDominio = 'outro.local'
    Igual $true ((SecaoMudancas $outro 6) -like '*é do domínio outro.local; a comparação só é feita entre levantamentos do mesmo domínio.*') 'outro domínio'
    $vazio = Obj ($base.Clone())
    Igual 0 @(CardsMudancas $vazio).Count 'sem mudanças: sem card'
    Igual $true ((SecaoMudancas $vazio 6) -like '*Mudanças desde 01/10/2026*Nenhuma mudança encontrada.*') 'sem mudanças: texto'

    $p = Obj ($base.Clone()); $p.Mudancas = @(NovaMudanca 'Pastas' 'Nova' 'nova' '' 'D:\Nova')
    Igual $true ((CardsMudancas $p) -like '<div class="card sev-baixa">*1 mudança desde 01/10/2026*') 'só pasta nova: baixa'
    $x = NovaMudanca 'Acesso' 'Financeiro · Ana' 'ganhou' 'Sem acesso' 'Leitura'; $x.Pasta = 'Financeiro'
    $s = Obj ($base.Clone()); $s.Mudancas = @($x)
    Igual $true ((CardsMudancas $s) -like '*sev-media*pasta sensível*') 'acesso a pasta sensível: media'
    $x.Pasta = 'Compras'
    Igual $true ((CardsMudancas $s) -like '*sev-baixa*') 'pasta comum: baixa'
    $ad = Obj ($base.Clone()); $ad.Mudancas = @(NovaMudanca 'Administradores do servidor' 'Davi' 'entrou' '' ''); $ad.NaoComparado = @('o levantamento anterior não incluiu as GPOs')
    Igual $true ((CardsMudancas $ad) -like '*sev-media*Administradores do servidor*') 'admin entrou: media'
    $sec = SecaoMudancas $ad 7
    Igual $true ($sec -like '<section id="mod-mudancas"><h2><span>7</span>Mudanças desde 01/10/2026</h2>*') 'título'
    Igual $true ($sec -like "*$(HE 'não comparado: o levantamento anterior não incluiu as GPOs')*") 'não comparado'
    Igual $true ($sec -like '*<td>Davi</td><td>entrou</td>*') 'linha da tabela'
}
Teste 'Invoke-Modulos: sem Comparar grava só o 00_Snapshot' {
    . $amb
    $dir = NovoDir
    try {
        $Saida = $dir
        $MOD = Invoke-Modulos
        Igual (Join-Path $dir '00_Snapshot_20261027_1405.csv') $MOD.Snapshot 'Snapshot'
        Igual '00_Snapshot_20261027_1405.csv' (@(Get-ChildItem -LiteralPath $dir | ForEach-Object { $_.Name }) -join ',') 'arquivos'
        Igual $null $MOD.Mudancas 'Mudancas'
    } finally { Remove-Item -LiteralPath $dir -Recurse -Force }
}
Teste 'Invoke-Modulos: Comparar sem anterior, depois com mudança' {
    . $amb
    $dir = NovoDir
    try {
        $Saida = $dir; $Comparar = [switch]$true
        $MOD = Invoke-Modulos 6>$null
        Igual $true $MOD.Mudancas.Ok 'Ok'
        Igual $null $MOD.Mudancas.DataAnterior 'sem anterior'
        Igual 1 @(Get-ChildItem -LiteralPath $dir).Count 'só o 00'

        $hoje = [datetime]'2026-10-28 09:30'; $stamp = $hoje.ToString('yyyyMMdd_HHmm')
        $ADM = @('ana', 'jose')
        $MOD = Invoke-Modulos 6>$null
        Igual '2026-10-27 14:05' $MOD.Mudancas.DataAnterior.ToString('yyyy-MM-dd HH:mm') 'DataAnterior'
        Igual 'entrou:José Conceição (jose):>' (Mud $MOD.Mudancas 'Administradores do servidor') 'mudança'
        Igual (Join-Path $dir '04_Mudancas_20261028_0930.csv') $MOD.Mudancas.Csv 'Csv'
        $csv = @(Import-Csv -LiteralPath $MOD.Mudancas.Csv -Delimiter ';')
        Igual 'Categoria,Item,Mudanca,Antes,Depois' (($csv[0].PSObject.Properties | ForEach-Object { $_.Name }) -join ',') 'colunas do 04'
        Igual 'José Conceição (jose)' $csv[0].Item 'acento no 04'

        $CompararCom = Join-Path $dir 'nao-existe.csv'
        $MOD = Invoke-Modulos 6>$null
        Igual $false $MOD.Mudancas.Ok 'CompararCom inexistente → erro tratado'
        Igual $true ([bool]$MOD.Snapshot) 'snapshot gravado mesmo assim'
    } finally { Remove-Item -LiteralPath $dir -Recurse -Force }
}

# ---- Fase 3: contas privilegiadas e de risco ----
$h3 = [datetime]'2026-10-27 14:05'
# Usuário como o Get-ADUser devolve (padrão: ativo, sem risco, logou ontem, senha de 10 dias)
function Usr([string]$login, [hashtable]$h) {
    $o = @{ SamAccountName = $login; DisplayName = "Nome $login"; DistinguishedName = "CN=$login,OU=TI,DC=empresa,DC=local"; Enabled = $true
        PasswordNotRequired = $false; TrustedForDelegation = $false; DoesNotRequirePreAuth = $false; AllowReversiblePasswordEncryption = $false
        ServicePrincipalNames = @(); PasswordLastSet = $h3.AddDays(-10); LastLogonDate = $h3.AddDays(-1); whenCreated = $h3.AddDays(-1000) }
    if ($h) { foreach ($k in @($h.Keys)) { $o[$k] = $h[$k] } }
    Obj $o
}
# Membro de grupo no formato comum (como o Read-ModContas devolve)
function Mbr([string]$login, [hashtable]$h) { ConvertFrom-UsuarioAD (Usr $login $h) }
function Grp([string]$nome, $membros) { Obj @{ Nome = $nome; Sid = "S-$nome"; Membros = @($membros) } }
function Bruto($grupos, $usuarios, $comps, $adm) { Obj @{ Grupos = @($grupos); Usuarios = @($usuarios); Computadores = @($comps); Admin500 = $adm; Avisos = @() } }
function Achar($b) { Get-AchadosContas $b $h3 90 'svc|servico|serviço|service|sql|backup' 365 'Terceir|Tercer|Fornecedor|Externo' }
function Ids($r, [string]$id) { if ($r.Achados.ContainsKey($id)) { (@($r.Achados[$id]) | ForEach-Object { $_.Login }) -join ',' } else { '' } }

Teste 'ConvertTo-LdapEscape' {
    Igual 'CN=A \28B\29\2a' (ConvertTo-LdapEscape 'CN=A (B)*') 'escape'
    Igual 'a\5cb' (ConvertTo-LdapEscape 'a\b') 'barra'
}
Teste 'ConvertFrom-FileTimeAD' {
    Igual $null (ConvertFrom-FileTimeAD 0) 'zero'
    Igual $null (ConvertFrom-FileTimeAD $null) 'vazio'
    Igual $null (ConvertFrom-FileTimeAD ([int64]::MaxValue)) 'nunca expira'
    Igual '2026-10-27' (ConvertFrom-FileTimeAD $h3.ToFileTime()).ToString('yyyy-MM-dd') 'data'
}
Teste 'contas: atributos de risco' {
    $r = Achar (Bruto @() @((Usr 'normal'), (Usr 'a' @{ PasswordNotRequired = $true }), (Usr 'b' @{ TrustedForDelegation = $true }),
        (Usr 'c' @{ DoesNotRequirePreAuth = $true }), (Usr 'd' @{ AllowReversiblePasswordEncryption = $true; PasswordNotRequired = $true })) @() $null)
    Igual 'a,d' (Ids $r 'SenhaNaoExigida') 'SenhaNaoExigida'
    Igual 'b' (Ids $r 'DelegacaoIrrestrita') 'DelegacaoIrrestrita'
    Igual 'c' (Ids $r 'SemPreAuth') 'SemPreAuth'
    Igual 'd' (Ids $r 'Reversivel') 'Reversivel'
    Igual 'a,b,c,d' ((@($r.Contas) | ForEach-Object { $_.Login }) -join ',') 'conta sem achado fica fora'
    Igual 0 @($r.Privilegiadas).Count 'nenhuma privilegiada'
}
Teste 'contas: nunca fez logon (fronteira de DiasInativo)' {
    $r = Achar (Bruto @() @((Usr 'exato' @{ LastLogonDate = $null; whenCreated = $h3.AddDays(-90) }),
        (Usr 'velha' @{ LastLogonDate = $null; whenCreated = $h3.AddDays(-91) }),
        (Usr 'logou' @{ whenCreated = $h3.AddDays(-91) })) @() $null)
    Igual 'velha' (Ids $r 'NuncaLogou') 'NuncaLogou'
}
Teste 'contas: serviço por SPN, por login e só pela OU' {
    $r = Achar (Bruto @() @(
        (Usr 'web01' @{ ServicePrincipalNames = @('HTTP/web01'); PasswordLastSet = $h3.AddDays(-400) }),
        (Usr 'web02' @{ ServicePrincipalNames = @('HTTP/web02'); PasswordLastSet = $h3.AddDays(-30) }),
        (Usr 'svc_backup' @{ PasswordLastSet = $h3.AddDays(-365) }),
        (Usr 'svc_erp' @{ PasswordLastSet = $h3.AddDays(-366) }),
        (Usr 'joao' @{ DistinguishedName = 'CN=joao,OU=Contas SQL,DC=empresa,DC=local'; PasswordLastSet = $null }),
        (Usr 'maria' @{ DistinguishedName = 'CN=maria,OU=SQL Users,OU=TI,DC=empresa,DC=local'; PasswordLastSet = $h3.AddDays(-30) })) @() $null)
    Igual 'joao,svc_erp,web01' (Ids $r 'Servico') 'Servico (SPN fora do padrão, login, OU; exatamente 365 dias não)'
    Igual 'joao,svc_erp,web01' ((@($r.Contas) | ForEach-Object { $_.Login }) -join ',') 'serviço com senha nova não entra'
    Igual $true ((@($r.Contas) | Where-Object { $_.Login -eq 'joao' }).Servico) 'Servico pela OU'
}
Teste 'contas: privilegiadas sem uso, terceiro e grupos' {
    $g1 = Grp 'Domain Admins' @((Mbr 'adm1'), (Mbr 'exato' @{ LastLogonDate = $h3.AddDays(-90) }), (Mbr 'parado' @{ LastLogonDate = $h3.AddDays(-91) }),
        (Mbr 'off' @{ Enabled = $false; LastLogonDate = $null }), (Mbr 'forn' @{ DistinguishedName = 'CN=forn,OU=Fornecedores,DC=empresa,DC=local' }),
        (Mbr 'novo' @{ LastLogonDate = $null; whenCreated = $h3.AddDays(-5) }))
    $g2 = Grp 'Backup Operators' @((Mbr 'adm1'), (Mbr 'bkp'))
    $r = Achar (Bruto @($g1, $g2) @((Usr 'adm1' @{ PasswordNotRequired = $true }), (Usr 'parado' @{ LastLogonDate = $h3.AddDays(-91) })) @() $null)
    Igual 'parado' (Ids $r 'PrivInativo') 'PrivInativo (exatamente 90 dias não, desativada não, recém-criada não)'
    Igual 'forn' (Ids $r 'PrivTerceiro') 'PrivTerceiro'
    Igual 'adm1' (Ids $r 'SenhaNaoExigida') 'achado de usuário na conta privilegiada'
    Igual 'Domain Admins:5/6, Backup Operators:2/2' ((@($r.Grupos) | ForEach-Object { "$($_.Nome):$($_.Ativos)/$($_.Total)" }) -join ', ') 'contagem por grupo'
    $adm1 = @($r.Contas) | Where-Object { $_.Login -eq 'adm1' }
    Igual 'Domain Admins, Backup Operators' (@($adm1.Grupos) -join ', ') 'uma linha com os dois grupos'
    Igual 7 @($r.Privilegiadas).Count 'privilegiadas'
}
Teste 'contas: computadores com delegação (DC e RODC ignorados)' {
    $pc = { param($n, $pg, $en) Obj @{ Name = $n; SamAccountName = "$n$"; DistinguishedName = "CN=$n,OU=Servidores,DC=empresa,DC=local"; PrimaryGroupID = $pg; Enabled = $en
        LastLogonDate = $h3; PasswordLastSet = $h3; whenCreated = $h3.AddDays(-500) } }
    $r = Achar (Bruto @() @() @((& $pc 'DC01' 516 $true), (& $pc 'RODC1' 521 $true), (& $pc 'APP01' 515 $true), (& $pc 'VELHO' 515 $false)) $null)
    Igual 'APP01$' (Ids $r 'DelegacaoIrrestrita') 'DelegacaoIrrestrita'
    Igual 'computador' (@($r.Achados['DelegacaoIrrestrita'])[0].Tipo) 'Tipo'
}
Teste 'contas: Administrador interno' {
    $r = Achar (Bruto @() @() @() (Usr 'Administrator'))
    Igual 'Administrator' (Ids $r 'Admin500') 'nome padrão'
    Igual 'baixa' $r.Admin500Sev 'senha recente: baixa'
    $r = Achar (Bruto @() @() @() (Usr 'Administrador' @{ PasswordLastSet = $h3.AddDays(-400) }))
    Igual 'media' $r.Admin500Sev 'senha antiga: media'
    Igual '' (Ids (Achar (Bruto @() @() @() (Usr 'adm.empresa'))) 'Admin500') 'renomeado'
    Igual '' (Ids (Achar (Bruto @() @() @() (Usr 'Administrator' @{ Enabled = $false }))) 'Admin500') 'desativado'
}
Teste 'contas: cards, seção, CSV e snapshot' {
    . $amb
    $g = Grp 'Domain Admins' @((Mbr 'ana' @{ DisplayName = 'Ana Lúcia'; DistinguishedName = 'CN=ana,OU=Terceiros,DC=empresa,DC=local'; LastLogonDate = $h3.AddDays(-200) }))
    $r = Achar (Bruto @($g) @((Usr 'x' @{ DoesNotRequirePreAuth = $true })) @() $null)
    $r.Avisos += 'Enterprise/Schema Admins não puderam ser lidos no domínio raiz raiz.local: negado'
    $cards = @(CardsContas $r)
    Igual 3 $cards.Count 'um card por tipo'
    Igual $true ($cards[0] -like '<div class="card sev-alta">*<h3>1 conta sem pré-autenticação Kerberos</h3>*AS-REP roasting*<td>Nome x</td><td>x</td>*') 'card SemPreAuth'
    Igual $true ($cards[1] -like '*sev-alta*1 conta de terceiros com privilégio no domínio*<td>Domain Admins</td>*') 'card PrivTerceiro'
    Igual $true ($cards[2] -like '*sev-media*1 conta privilegiada sem uso há mais de 90 dias*') 'card PrivInativo'
    $sec = SecaoContas $r 7
    Igual $true ($sec -like '<section id="mod-contas"><h2><span>7</span>Contas privilegiadas</h2>*') 'título'
    Igual $true ($sec -like '*raiz raiz.local: negado.*') 'aviso'
    Igual $true ($sec -like '*<td>Domain Admins</td><td class="num">1</td>*') 'tabela de grupos'
    Igual $true ($sec -like '*t-terc">terceiro*sem uso &gt;90d*') 'tags'
    Igual $true ((SecaoContas (Achar (Bruto @() @() @() $null)) 6) -like '*Nenhuma conta nos grupos privilegiados.*') 'sem privilegiadas'

    $dir = NovoDir
    try {
        Write-CsvContas $r (Join-Path $dir '05.csv')
        $csv = @(Import-Csv -LiteralPath $r.Csv -Delimiter ';')
        Igual 'Login,Nome,Tipo,Ativa,GruposPrivilegiados,ContaServico,SPN,Achados,UltimoLogon,SenhaAlteradaEm,CriadaEm,OU' (($csv[0].PSObject.Properties | ForEach-Object { $_.Name }) -join ',') 'colunas do 05'
        Igual 'Ana Lúcia' $csv[0].Nome 'acento no 05'
        Igual 'Terceiro com privilégio no domínio, Privilegiada sem uso' $csv[0].Achados 'Achados'
        Igual 'OU=Terceiros,DC=empresa,DC=local' $csv[0].OU 'OU'
    } finally { Remove-Item -LiteralPath $dir -Recurse -Force }

    $MOD = @{ Marcados = @('ContasDominio'); Contas = $r }
    $s = ConvertTo-Snapshot @(Get-LinhasSnapshot $MOD)
    Igual 'ContasDominio' $s.Meta['Modulos'] 'Modulos'
    Igual 'Domain Admins|ana' (@($s.Dados['Privilegiado'].Keys) -join ',') 'Privilegiado'
    Igual 3 @(CardsModulos $MOD).Count 'CardsModulos'
    Igual $true ((SecoesModulos $MOD) -like '<section id="mod-contas"><h2><span>6</span>*') 'SecoesModulos'
}

# ---- Fase 4: política de senha ----
# Política que passa em tudo; $h altera propriedades
function Pol([hashtable]$h) {
    $p = @{ MinPasswordLength = 14; ComplexityEnabled = $true; PasswordHistoryCount = 24; ReversibleEncryptionEnabled = $false
        LockoutThreshold = 5; LockoutDuration = [TimeSpan]::FromMinutes(15); LockoutObservationWindow = [TimeSpan]::FromMinutes(15)
        MinPasswordAge = [TimeSpan]::FromDays(1); MaxPasswordAge = [TimeSpan]::Zero }
    if ($h) { foreach ($k in @($h.Keys)) { $p[$k] = $h[$k] } }
    Obj $p
}
# Severidade de um item da avaliação
function Sev($pol, [string]$item) { $i = @(Get-AvaliacaoSenha $pol) | Where-Object { $_.Item -eq $item }; if ($i) { $i.Sev } else { '(ausente)' } }

Teste 'senha: política dentro da referência' {
    $itens = @(Get-AvaliacaoSenha (Pol))
    Igual 9 $itens.Count 'itens'
    Igual 'ok,ok,ok,ok,ok,ok,ok,ok,info' ((@($itens) | ForEach-Object { $_.Sev }) -join ',') 'severidades'
    Igual 'nunca expira' $itens[8].Atual 'validade zero'
    Igual '15 min' $itens[5].Atual 'duração'
    $r = Get-ResultadoSenha (Obj @{ Padrao = (Pol); Refinadas = @(); Avisos = @() })
    Igual 0 @($r.Abaixo).Count 'nada abaixo'
    Igual 0 @(CardsSenha $r).Count 'sem card'
}
Teste 'senha: fronteiras de tamanho, histórico e bloqueio' {
    Igual 'alta' (Sev (Pol @{ MinPasswordLength = 7 }) 'Tamanho mínimo') '7'
    Igual 'media' (Sev (Pol @{ MinPasswordLength = 8 }) 'Tamanho mínimo') '8'
    Igual 'media' (Sev (Pol @{ MinPasswordLength = 13 }) 'Tamanho mínimo') '13'
    Igual 'ok' (Sev (Pol @{ MinPasswordLength = 14 }) 'Tamanho mínimo') '14'
    Igual 'baixa' (Sev (Pol @{ PasswordHistoryCount = 23 }) 'Histórico') 'histórico 23'
    Igual 'ok' (Sev (Pol @{ PasswordHistoryCount = 24 }) 'Histórico') 'histórico 24'
    Igual 'alta' (Sev (Pol @{ LockoutThreshold = 0 }) 'Bloqueio por tentativas') 'limiar 0'
    Igual 'ok' (Sev (Pol @{ LockoutThreshold = 1 }) 'Bloqueio por tentativas') 'limiar 1'
    Igual 'ok' (Sev (Pol @{ LockoutThreshold = 10 }) 'Bloqueio por tentativas') 'limiar 10'
    Igual 'baixa' (Sev (Pol @{ LockoutThreshold = 11 }) 'Bloqueio por tentativas') 'limiar 11'
    Igual 'alta' (Sev (Pol @{ ComplexityEnabled = $false }) 'Complexidade') 'complexidade'
    Igual 'alta' (Sev (Pol @{ ReversibleEncryptionEnabled = $true }) 'Criptografia reversível') 'reversível'
}
Teste 'senha: duração, contador e idade mínima' {
    Igual 'ok' (Sev (Pol @{ LockoutDuration = [TimeSpan]::Zero }) 'Duração do bloqueio') 'duração 0 = manual'
    Igual 'até desbloqueio manual' (@(Get-AvaliacaoSenha (Pol @{ LockoutDuration = [TimeSpan]::Zero }))[5].Atual) 'texto manual'
    Igual 'ok' (Sev (Pol @{ LockoutDuration = [TimeSpan]::MinValue }) 'Duração do bloqueio') 'duração "infinita" = manual'
    Igual 'baixa' (Sev (Pol @{ LockoutDuration = [TimeSpan]::FromMinutes(14) }) 'Duração do bloqueio') 'duração 14'
    Igual 'ok' (Sev (Pol @{ LockoutDuration = [TimeSpan]::FromMinutes(15) }) 'Duração do bloqueio') 'duração 15'
    Igual 'baixa' (Sev (Pol @{ LockoutObservationWindow = [TimeSpan]::FromMinutes(14) }) 'Zerar contador após') 'contador 14'
    Igual 'ok' (Sev (Pol @{ LockoutObservationWindow = [TimeSpan]::FromMinutes(30) }) 'Zerar contador após') 'contador 30'
    Igual '(ausente)' (Sev (Pol @{ LockoutThreshold = 0; LockoutDuration = [TimeSpan]::FromMinutes(1) }) 'Duração do bloqueio') 'sem limiar: sem duração'
    Igual 7 @(Get-AvaliacaoSenha (Pol @{ LockoutThreshold = 0 })).Count 'sem limiar: 7 itens'
    Igual 'baixa' (Sev (Pol @{ MinPasswordAge = [TimeSpan]::Zero }) 'Idade mínima') 'idade 0'
    Igual 'baixa' (Sev (Pol @{ MinPasswordAge = [TimeSpan]::FromHours(23) }) 'Idade mínima') 'idade 23h'
    Igual '42 dias' (@(Get-AvaliacaoSenha (Pol @{ MaxPasswordAge = [TimeSpan]::FromDays(42) }))[8].Atual) 'validade 42 dias'
    Igual '1 hora' (TextoDuracao ([TimeSpan]::FromHours(1))) 'hora'
    Igual '0,5 min' (TextoDuracao ([TimeSpan]::FromSeconds(30))) 'fração'
}
Teste 'senha: FGPP, card e seção' {
    . $amb
    $f1 = Obj @{ Nome = 'Servicos'; Precedencia = 20; AplicaA = @('Contas de Serviço'); Politica = (Pol @{ MinPasswordLength = 20; LockoutThreshold = 0 }) }
    $f2 = Obj @{ Nome = 'Admins'; Precedencia = 10; AplicaA = @(); Politica = (Pol) }
    $r = Get-ResultadoSenha (Obj @{ Padrao = (Pol @{ MinPasswordLength = 8; PasswordHistoryCount = 12 }); Refinadas = @($f1, $f2); Avisos = @() })
    Igual 'Admins,Servicos' ((@($r.Refinadas) | ForEach-Object { $_.Nome }) -join ',') 'ordem por precedência'
    Igual 'alta' $r.Sev 'maior severidade (FGPP sem bloqueio)'
    Igual 'Padrão do domínio:Tamanho mínimo,Padrão do domínio:Histórico,Refinada: Servicos:Bloqueio por tentativas' ((@($r.Abaixo) | ForEach-Object { "$($_.Politica):$($_.Linha.Item)" }) -join ',') 'itens abaixo'
    $cards = @(CardsSenha $r)
    Igual 1 $cards.Count 'um card'
    Igual $true ($cards[0] -like '<div class="card sev-alta">*<h3>3 itens da política de senha abaixo da referência</h3>*<td class="sev-media">8 caracteres</td>*') 'card'
    $sec = SecaoSenha $r 8
    Igual $true ($sec -like '<section id="mod-senha"><h2><span>8</span>Política de senha</h2>*') 'título'
    Igual $true ($sec -like '*<td class="sev-media">8 caracteres</td><td>14 ou mais</td><td><span class="tag t-bad">abaixo</span>*') 'linha abaixo'
    Igual $true ($sec -like '*<span class="tag">informativo</span>*') 'informativo'
    Igual $true ($sec -like '*<td>Admins</td><td class="num">10</td><td>ningu*') 'FGPP sem alvo'
    Igual $true ($sec -like '*31/01/2025*') 'nota da baseline'
    Igual 'Grupo, TI' (Get-NomeCN 'CN=Grupo\, TI,OU=Grupos,DC=empresa,DC=local') 'Get-NomeCN'
    $r2 = Get-ResultadoSenha (Obj @{ Padrao = (Pol); Refinadas = @(); Avisos = @('as políticas refinadas (FGPP) não puderam ser lidas: negado') })
    $sec2 = SecaoSenha $r2 6
    Igual $true ($sec2 -like '*Observação: as pol*refinadas (FGPP)*' -and $sec2 -notlike '*Nenhuma política refinada*') 'aviso de FGPP'
    $MOD = @{ Marcados = @('PoliticaSenha'); Senha = $r }
    Igual 1 @(CardsModulos $MOD).Count 'CardsModulos'
    Igual $true ((SecoesModulos $MOD) -like '<section id="mod-senha"><h2><span>6</span>*') 'SecoesModulos'
    Igual '' (@(Get-LinhasSnapshot $MOD) | Where-Object { $_.Tipo -eq '_meta' -and $_.Chave -eq 'Modulos' }).Valor 'snapshot não muda'
}

# ---- Fase 5: GPOs ----
$domDN = 'DC=empresa,DC=local'
$pol = 'cn=policies,cn=system,DC=empresa,DC=local'
# GPO como o Read-ModGpo devolve (padrão: ativa, versão 5, pasta no SYSVOL)
function Gpo([string]$guid, [string]$nome, [hashtable]$h) {
    $o = @{ Guid = $guid; Nome = $nome; Flags = 0; Versao = 5; CriadaEm = $h3.AddDays(-300); AlteradaEm = $h3.AddDays(-30)
        Pasta = "\\empresa.local\SysVol\empresa.local\Policies\$guid"; PastaExiste = $true }
    if ($h) { foreach ($k in @($h.Keys)) { $o[$k] = $h[$k] } }
    Obj $o
}
function Alvo([string]$dn, [string]$link, $opt) { Obj @{ Dn = $dn; GPLink = $link; GPOptions = $opt } }
function Lnk([string]$guid, [int]$op) { "[LDAP://cn=$guid,$pol;$op]" }
function IdsGpo($r, [string]$id) { if ($r.Achados.ContainsKey($id)) { (@($r.Achados[$id]) | ForEach-Object { $_.Guid }) -join ',' } else { '' } }

Teste 'ConvertFrom-GPLink' {
    Igual 0 @(ConvertFrom-GPLink '' $domDN).Count 'vazio'
    $v = @(ConvertFrom-GPLink (Lnk '{aa}' 0) $domDN)
    Igual 1 $v.Count 'um vínculo'
    Igual '{AA}|cn=policies,cn=system,DC=empresa,DC=local|0|False|False|False' "$($v[0].Guid)|$($v[0].Dn)|$($v[0].Opcao)|$($v[0].Desativado)|$($v[0].Imposto)|$($v[0].Externo)" 'campos'
    $s = "[LDAP://CN={B1},CN=Policies,CN=System,DC=EMPRESA,DC=LOCAL;1][ldap://cn={b2},$pol;2][LDAP://cn={B3},$pol;3]"
    $v = @(ConvertFrom-GPLink $s $domDN)
    Igual '{B1}:True:False,{B2}:False:True,{B3}:True:True' ((@($v) | ForEach-Object { "$($_.Guid):$($_.Desativado):$($_.Imposto)" }) -join ',') 'três vínculos, opções 1/2/3 e maiúsculas'
    Igual 'False' $v[0].Externo 'mesmo domínio em maiúsculas'
    Igual 'True' (@(ConvertFrom-GPLink '[LDAP://cn={C1},cn=policies,cn=system,DC=filho,DC=empresa,DC=local;0]' $domDN))[0].Externo 'domínio filho é externo'
}
Teste 'Get-CaminhoDN' {
    Igual 'empresa.local' (Get-CaminhoDN $domDN $domDN) 'domínio'
    Igual 'empresa.local/TI/Sub, Norte' (Get-CaminhoDN 'OU=Sub\, Norte,OU=TI,DC=empresa,DC=local' $domDN) 'OU com vírgula'
    Igual 'site Matriz' (Get-CaminhoDN 'CN=Matriz,CN=Sites,CN=Configuration,DC=empresa,DC=local' $domDN) 'site'
}
Teste 'GPOs: um caso por achado' {
    $gpos = @((Gpo '{01}' 'Normal'), (Gpo '{02}' 'Solta'), (Gpo '{03}' 'Desligada' @{ Flags = 3 }), (Gpo '{04}' 'Meio' @{ Flags = 1 }),
        (Gpo '{05}' 'Nova' @{ Versao = 0 }), (Gpo '{06}' 'SemPasta' @{ PastaExiste = $false }), (Gpo '{07}' 'Pausada'), (Gpo '{08}' 'Site'))
    $vinc = @((Alvo $domDN "$(Lnk '{01}' 0)$(Lnk '{03}' 0)$(Lnk '{04}' 0)$(Lnk '{05}' 0)$(Lnk '{06}' 0)" $null),
        (Alvo 'OU=TI,DC=empresa,DC=local' "$(Lnk '{07}' 1)$(Lnk '{99}' 0)$(Lnk '{01}' 2)" 1),
        (Alvo 'OU=Vazia,DC=empresa,DC=local' '' 1),
        (Alvo 'CN=Matriz,CN=Sites,CN=Configuration,DC=empresa,DC=local' "$(Lnk '{08}' 0)[LDAP://cn={EE},cn=policies,cn=system,DC=outro,DC=local;0]" $null))
    $r = Get-AchadosGpo $gpos $vinc $domDN $true
    Igual '{02}' (IdsGpo $r 'SemVinculo') 'SemVinculo (site conta como vínculo)'
    Igual '{03}' (IdsGpo $r 'Desativada') 'Desativada (flags 1 não)'
    Igual '{07}' (IdsGpo $r 'VinculoDesativado') 'VinculoDesativado'
    Igual '{05}' (IdsGpo $r 'Vazia') 'Vazia'
    Igual '{99}' (IdsGpo $r 'Orfao') 'Orfao (externo não)'
    Igual '{06}' (IdsGpo $r 'SemSysvol') 'SemSysvol'
    Igual 'empresa.local/TI,empresa.local/Vazia' (@($r.Bloqueios) -join ',') 'herança bloqueada'
    Igual '{EE}' (@($r.Externos) | ForEach-Object { $_.Guid }) 'externo'
    $n = @($r.Gpos) | Where-Object { $_.Guid -eq '{01}' }
    Igual 'empresa.local, empresa.local/TI (imposto)' (TextoVinculos $n) 'vínculos da GPO'
    Igual 'usuário desativado' (@($r.Gpos) | Where-Object { $_.Guid -eq '{04}' }).Status 'status 1'
    Igual 'empresa.local/TI' (TextoVinculos $r.Orfaos[0]) 'onde está o órfão'
    $r2 = Get-AchadosGpo @(Gpo '{06}' 'SemPasta' @{ PastaExiste = $null }) @(Alvo $domDN (Lnk '{06}' 0) $null) $domDN $false
    Igual '' (IdsGpo $r2 'SemSysvol') 'SYSVOL inacessível: sem SemSysvol'
    Igual 0 @($r2.Achados.Keys).Count 'nenhum achado'
}
Teste 'GPOs: cards, seção, CSV e snapshot' {
    . $amb
    $gpos = @((Gpo '{01}' 'Senhas & Bloqueio'), (Gpo '{02}' 'Solta' @{ Versao = 0 }), (Gpo '{03}' 'SemPasta' @{ PastaExiste = $false }))
    $vinc = @((Alvo $domDN "$(Lnk '{01}' 2)$(Lnk '{03}' 0)$(Lnk '{99}' 0)" $null), (Alvo 'OU=Ação,DC=empresa,DC=local' (Lnk '{01}' 1) $null))
    $r = Get-AchadosGpo $gpos $vinc $domDN $true
    $r.Avisos = @('os vínculos nos sites não puderam ser lidos: negado')
    $cards = @(CardsGpo $r)
    Igual 2 $cards.Count 'dois cards'
    Igual $true ($cards[0] -like '<div class="card sev-media">*<h3>2 GPOs com problema de integridade</h3>*<td>Vínculo para GPO inexistente</td><td>{99}</td><td>empresa.local</td>*<td>Pasta ausente no SYSVOL</td><td>SemPasta</td>*') 'card integridade'
    Igual $true ($cards[1] -like '<div class="card sev-baixa">*<h3>3 GPOs para limpeza</h3>*') 'card limpeza (sem vínculo + vínculo desativado + vazia)'
    $sec = SecaoGpo $r 9
    Igual $true ($sec -like '<section id="mod-gpos"><h2><span>9</span>GPOs</h2>*') 'título'
    Igual $true ($sec -like '*Observação: os v*nculos nos sites*') 'aviso'
    Igual $true ($sec -like '*<td><a href="#gpo-01">Senhas &amp; Bloqueio</a><span class="tag t-warn">vínculo desativado</span></td><td>ativa</td><td class="num">5</td><td class="num">27/09/2026</td><td>empresa.local<span class="tag">imposto</span><br>empresa.local/A*o<span class="tag t-warn">desativado</span></td>*') 'linha da GPO'
    Igual $true ($sec -like '*Vínculos para GPOs inexistentes*<td>{99}</td>*Nenhuma OU com herança bloqueada.*') 'órfãos e herança'

    $dir = NovoDir
    try {
        Write-CsvGpo $r (Join-Path $dir '06.csv')
        $csv = @(Import-Csv -LiteralPath $r.Csv -Delimiter ';')
        Igual 'GPO,GUID,Status,Versao,CriadaEm,AlteradaEm,VinculadaEm,VinculosDesativados,PastaSYSVOL,Configuracoes,Achados,ErrosLeitura' (($csv[0].PSObject.Properties | ForEach-Object { $_.Name }) -join ',') 'colunas do 06'
        Igual 4 $csv.Count 'GPOs + órfão'
        $linha = { param($guid) @($csv) | Where-Object { $_.GUID -eq $guid } }
        $l1 = & $linha '{01}'; $l2 = & $linha '{02}'; $l3 = & $linha '{03}'; $l9 = & $linha '{99}'
        Igual 'empresa.local (imposto), empresa.local/Ação (desativado)|empresa.local/Ação|ok' "$($l1.VinculadaEm)|$($l1.VinculosDesativados)|$($l1.PastaSYSVOL)" 'linha com acento'
        Igual 'ausente' $l3.PastaSYSVOL 'PastaSYSVOL ausente'
        Igual '(inexistente)|{99}|Vínculo para GPO inexistente' "$($l9.GPO)|$($l9.GUID)|$($l9.Achados)" 'linha do órfão'
        Igual '{99}' $csv[3].GUID 'órfão no fim'
        Igual 'Sem nenhum vínculo, Vazia (versão 0)' $l2.Achados 'Achados'
    } finally { Remove-Item -LiteralPath $dir -Recurse -Force }

    $MOD = @{ Marcados = @('GPOs'); Gpo = $r }
    $s = ConvertTo-Snapshot @(Get-LinhasSnapshot $MOD)
    Igual 'GPOs' $s.Meta['Modulos'] 'Modulos'
    Igual 'Senhas & Bloqueio|0|5' $s.Dados['GPO']['{01}'] 'linha GPO'
    Igual 3 @($s.Dados['GPO'].Keys).Count 'órfão fora do snapshot'
    Igual 2 @(CardsModulos $MOD).Count 'CardsModulos'
    Igual $true ((SecoesModulos $MOD) -like '<section id="mod-gpos"><h2><span>6</span>*') 'SecoesModulos'
}

# ---- Conteúdo das GPOs ----
# Registry.pol falso: cada entrada @(chave, valor, tipo, dado); tipo 4 = DWORD, senão texto
function PolBytes($entradas) {
    $u = [Text.Encoding]::Unicode
    $b = New-Object System.Collections.Generic.List[byte]
    $b.AddRange([Text.Encoding]::ASCII.GetBytes('PReg')); $b.AddRange([BitConverter]::GetBytes([uint32]1))
    foreach ($e in $entradas) {
        $dado = if ($e[2] -eq 4) { [BitConverter]::GetBytes([uint32]$e[3]) } else { $u.GetBytes("$($e[3])`0") }
        $b.AddRange($u.GetBytes("[$($e[0])`0;$($e[1])`0;"))
        $b.AddRange([BitConverter]::GetBytes([uint32]$e[2])); $b.AddRange($u.GetBytes(';'))
        $b.AddRange([BitConverter]::GetBytes([uint32]$dado.Length)); $b.AddRange($u.GetBytes(';'))
        $b.AddRange([byte[]]$dado); $b.AddRange($u.GetBytes(']'))
    }
    , $b.ToArray()
}
function Areas($itens) { (@($itens) | ForEach-Object { "$($_.Lado):$($_.Area)" }) -join ',' }
function Det($itens, [string]$area) { (@(@($itens) | Where-Object { $_.Area -eq $area })[0].Detalhes) -join ' / ' }
$xmlDrives = '<?xml version="1.0" encoding="utf-8"?><Drives clsid="{8FDDCC1A-0C3C-43cd-A6B4-71A6DF20DA8C}"><Drive name="H:" uid="{1}"><Properties action="U" path="\\srv\Usuários\%USERNAME%" label="Pessoal" letter="H"/></Drive><Drive name="S:"><Properties action="D" letter="S" path="\\srv\antigo"/></Drive></Drives>'
$xmlGrupos = '<?xml version="1.0" encoding="utf-8"?><Groups><Group name="Administradores (interno)"><Properties action="U" groupName="Administradores (interno)" deleteAllUsers="0"><Members><Member name="EMPRESA\TI" action="ADD" sid=""/><Member name="EMPRESA\Estagio" action="REMOVE" sid=""/></Members></Properties></Group><User name="Administrador (interno)"><Properties action="U" newName="admlocal" userName="Administrador (interno)" cpassword="j1Uyj3Vx8TY9LtLZil2uAuZkFQA" acctDisabled="0"/></User></Groups>'

Teste 'GPOs: extensões (CSE) e áreas sem detalhe' {
    $ext = '[{00000000-0000-0000-0000-000000000000}{79F92669-4224-476C-9C5C-6EFB4D87DF4A}][{35378EAC-683F-11D2-A89A-00C04FBBCFA2}{0F6B957D-509E-11D1-A7CC-0000F87571E3}][{827d319e-6eac-11d2-a4ea-00c04f79f83a}{803E14A0-B4FB-11D0-A0D0-00A0C90F574B}]'
    Igual '{35378EAC-683F-11D2-A89A-00C04FBBCFA2},{827D319E-6EAC-11D2-A4EA-00C04F79F83A}' (@(Get-CseGpo $ext) -join ',') 'CSEs (GUID zero fora, maiúsculas)'
    $g = Gpo '{10}' 'X' @{ ExtMaquina = $ext; ExtUsuario = '[{5794DAFD-BE60-433F-88A2-1A31939AC01F}{2EA1A81B-48E5-45E9-8BB7-A6E3AC170006}][{AAAAAAAA-0000-0000-0000-000000000000}{BBBBBBBB-0000-0000-0000-000000000000}]' }
    Igual 'Computador:Modelos administrativos,Computador:Configurações de segurança,Usuário:Mapeamento de unidades' (Areas (Get-ItensGpo $g)) 'áreas só pelas extensões (desconhecida ignorada)'
    $c = Obj @{ Itens = @(NovoItemGpo 'Usuário' 'registro' 'Modelos administrativos: Windows Update' @('AU\NoAutoUpdate = 0')); SenhaEm = @(); Erros = @() }
    $g2 = Gpo '{11}' 'Y' @{ ExtMaquina = $ext; ExtUsuario = '[{35378EAC-683F-11D2-A89A-00C04FBBCFA2}{0F6B957D-509E-11D1-A7CC-0000F87571E3}]'; Conteudo = $c }
    Igual 'Computador:Modelos administrativos,Computador:Configurações de segurança,Usuário:Modelos administrativos: Windows Update' (Areas (Get-ItensGpo $g2)) 'área lida não se repete; Computador antes de Usuário'
    Igual 0 @(Get-ItensGpo (Gpo '{12}' 'Z')).Count 'sem extensões e sem conteúdo'
}
Teste 'GPOs: Registry.pol e Modelos administrativos' {
    $b = PolBytes @(
        @('Software\Policies\Microsoft\Windows\WindowsUpdate\AU', 'NoAutoUpdate', 4, 0),
        @('Software\Policies\Microsoft\Windows\WindowsUpdate\AU', 'AUOptions', 4, 4),
        @('Software\Policies\Microsoft\WindowsFirewall\DomainProfile', 'EnableFirewall', 4, 1),
        @('Software\Policies\Microsoft\Windows\Control Panel\Desktop', 'ScreenSaveTimeOut', 1, '600'),
        @('Software\Policies\Microsoft\Windows\WindowsUpdate', '**del.WUServer', 1, ' '),
        @('Software\Policies\Fabricante\App', 'Modo', 1, 'ação'))
    $e = @(ConvertFrom-RegistryPol $b)
    Igual 6 $e.Count 'entradas'
    Igual 'Software\Policies\Microsoft\Windows\WindowsUpdate\AU|AUOptions|4|4' "$($e[1].Chave)|$($e[1].Valor)|$($e[1].Tipo)|$($e[1].Dado)" 'DWORD'
    Igual '600' $e[3].Dado 'texto'
    $i = @(ConvertTo-ItensRegistro $e 'Computador')
    Igual 'Computador:Modelos administrativos: Windows Update,Computador:Modelos administrativos: Firewall do Windows,Computador:Modelos administrativos: Área de trabalho e proteção de tela,Computador:Modelos administrativos: Fabricante' (Areas $i) 'áreas na ordem'
    Igual 'WindowsUpdate\AU\NoAutoUpdate = 0 / WindowsUpdate\AU\AUOptions = 4 / WindowsUpdate\WUServer: removido' (Det $i 'Modelos administrativos: Windows Update') 'Windows Update'
    Igual 'WindowsFirewall\DomainProfile\EnableFirewall = 1' (Det $i 'Modelos administrativos: Firewall do Windows') 'firewall'
    Igual 'Fabricante\App\Modo = ação' (Det $i 'Modelos administrativos: Fabricante') 'área pelo nome da chave'
    $erro = $null; try { ConvertFrom-RegistryPol ([Text.Encoding]::ASCII.GetBytes('XXXXXXXX')) } catch { $erro = $_.Exception.Message }
    Igual 'Registry.pol em formato desconhecido' $erro 'cabeçalho inválido'
}
Teste 'GPOs: GptTmpl.inf (configurações de segurança)' {
    $linhas = @('[Unicode]', 'Unicode=yes', '[System Access]', 'MinimumPasswordLength = 12', 'NewAdministratorName = "admin.local"',
        '[Event Audit]', 'AuditLogonEvents = 3', '[Privilege Rights]', 'SeRemoteInteractiveLogonRight = *S-1-5-32-544,*S-1-5-21-1-2-3-1105', 'SeDenyNetworkLogonRight = Convidado',
        '[Registry Values]', 'MACHINE\System\CurrentControlSet\Control\Lsa\LmCompatibilityLevel=4,5',
        'MACHINE\System\CurrentControlSet\Services\LanmanServer\Parameters\RequireSecuritySignature=4,1',
        'MACHINE\Software\Microsoft\Windows\CurrentVersion\Policies\System\LegalNoticeText=7,Uso restrito',
        '[Group Membership]', '*S-1-5-32-544__Memberof =', '*S-1-5-32-544__Members = *S-1-5-21-1-2-3-512,EMPRESA\suporte',
        '[Service General Setting]', '"Spooler",4,"D:AR(A;;CCLCSWRPWPDTLOCRRC;;;SY)"',
        '[File Security]', '"%SystemRoot%\System32\cmd.exe",2,"D:PAR(A;;FA;;;BA)"', '[Version]', 'signature="$CHICAGO$"')
    $res = { param($sid) if ($sid -eq 'S-1-5-21-1-2-3-1105') { 'EMPRESA\joao' } }
    $i = @(ConvertFrom-GptTmpl $linhas $res)
    Igual 'Computador:Política de contas (senha e bloqueio),Computador:Auditoria,Computador:Direitos de usuário,Computador:Opções de segurança,Computador:Grupos restritos,Computador:Serviços do sistema,Computador:Permissões de arquivos e pastas' (Areas $i) 'áreas na ordem da tabela'
    Igual 'Tamanho mínimo da senha: 12 / Renomear a conta Administrador para: admin.local' (Det $i 'Política de contas (senha e bloqueio)') 'política de contas'
    Igual 'Eventos de logon: sucesso e falha' (Det $i 'Auditoria') 'auditoria'
    Igual 'Permitir logon pela Área de Trabalho Remota: Administrators, EMPRESA\joao / Negar acesso a este computador pela rede: Convidado' (Det $i 'Direitos de usuário') 'direitos (SID resolvido, conhecido e nome)'
    Igual 'Nível de autenticação LAN Manager: 5 / Exigir assinatura SMB (servidor): 1 / Texto do aviso de logon: Uso restrito' (Det $i 'Opções de segurança') 'opções de segurança'
    Igual 'Administrators: membros definidos = S-1-5-21-1-2-3-512, EMPRESA\suporte' (Det $i 'Grupos restritos') 'grupos restritos (Memberof vazio fora)'
    Igual 'Spooler: desativado' (Det $i 'Serviços do sistema') 'serviço'
    Igual '%SystemRoot%\System32\cmd.exe' (Det $i 'Permissões de arquivos e pastas') 'permissão de arquivo'
}
Teste 'GPOs: scripts, Preferências e redirecionamento' {
    Igual 'Logon: mapear.bat /s / Logon: \\empresa.local\NETLOGON\inv.vbs / Logoff: sair.cmd' ((@(ConvertFrom-ScriptsIni @('[Logon]', '0CmdLine=mapear.bat', '0Parameters=/s', '1CmdLine=\\empresa.local\NETLOGON\inv.vbs', '1Parameters=', '[Logoff]', '0CmdLine=sair.cmd', '0Parameters=') '')) -join ' / ') 'scripts.ini'
    Igual 'Inicialização (PowerShell): a.ps1' ((@(ConvertFrom-ScriptsIni @('[Startup]', '0CmdLine=a.ps1') ' (PowerShell)')) -join ' / ') 'psscripts.ini'
    $r = ConvertFrom-PreferenciaXml ([xml]$xmlDrives) 'Drives'
    Igual 'H: → \\srv\Usuários\%USERNAME% / S: → \\srv\antigo (excluir)|False' "$(@($r.Detalhes) -join ' / ')|$($r.Senha)" 'unidades'
    $r = ConvertFrom-PreferenciaXml ([xml]$xmlGrupos) 'Groups'
    Igual 'grupo local Administradores (interno): adiciona EMPRESA\TI; remove EMPRESA\Estagio / usuário local Administrador (interno) → renomear para admlocal (define a senha)|True' "$(@($r.Detalhes) -join ' / ')|$($r.Senha)" 'grupos e senha exposta'
    Igual $false ((@($r.Detalhes) -join ' ') -like '*j1Uyj3*') 'o cpassword não vai para o relatório'
    $r = ConvertFrom-PreferenciaXml ([xml]'<Printers><SharedPrinter name="HP"><Properties action="C" path="\\srv\HP" default="1"/></SharedPrinter></Printers>') 'Printers'
    Igual '\\srv\HP (padrão) (criar)' (@($r.Detalhes) -join ' / ') 'impressora'
    $r = ConvertFrom-PreferenciaXml ([xml]'<ScheduledTasks><TaskV2 name="Backup"><Properties action="R" runAs="NT AUTHORITY\System"><Task><Actions><Exec><Command>C:\bkp.cmd</Command></Exec></Actions></Task></Properties></TaskV2></ScheduledTasks>') 'ScheduledTasks'
    Igual 'Backup: C:\bkp.cmd (executa como NT AUTHORITY\System) (substituir)' (@($r.Detalhes) -join ' / ') 'tarefa agendada'
    Igual 'My Documents → \\srv\docs\%USERNAME%' ((@(ConvertFrom-Fdeploy @('[FolderStatus]', 'My Documents=11', '[My Documents]', 's-1-1-0=\\srv\docs\%USERNAME%'))) -join ' / ') 'redirecionamento'
}
Teste 'GPOs: Read-ConteudoGpo numa pasta de GPO' {
    $dir = NovoDir
    try {
        $novo = { param($rel) $f = Join-Path $dir $rel; [void](New-Item -ItemType Directory -Force -Path (Split-Path $f)); $f }
        [IO.File]::WriteAllLines((& $novo 'Machine\Microsoft\Windows NT\SecEdit\GptTmpl.inf'), [string[]]@('[System Access]', 'MinimumPasswordLength = 14'), [Text.Encoding]::Unicode)
        [IO.File]::WriteAllBytes((& $novo 'Machine\Registry.pol'), (PolBytes @(, @('Software\Policies\Microsoft\Windows\WindowsUpdate\AU', 'NoAutoUpdate', 4, 1))))
        [IO.File]::WriteAllLines((& $novo 'Machine\Scripts\scripts.ini'), [string[]]@('[Startup]', '0CmdLine=início.bat', '0Parameters='), [Text.Encoding]::Unicode)
        $utf8 = New-Object System.Text.UTF8Encoding $false
        [IO.File]::WriteAllText((& $novo 'User\Preferences\Drives\Drives.xml'), $xmlDrives, $utf8)
        [IO.File]::WriteAllText((& $novo 'User\Preferences\Groups\Groups.xml'), $xmlGrupos, $utf8)
        [IO.File]::WriteAllLines((& $novo 'User\Documents & Settings\fdeploy1.ini'), [string[]]@('[FolderStatus]', 'Desktop=11', '[Desktop]', 's-1-1-0=\\srv\desk'), [Text.Encoding]::Unicode)
        [IO.File]::WriteAllText((& $novo 'User\Registry.pol'), 'XXXXXXXX')
        $c = Read-ConteudoGpo $dir $null
        Igual 'Computador:Política de contas (senha e bloqueio),Computador:Modelos administrativos: Windows Update,Computador:Scripts,Usuário:Mapeamento de unidades,Usuário:Usuários e grupos locais,Usuário:Redirecionamento de pastas' (Areas $c.Itens) 'áreas dos dois lados'
        Igual 'Inicialização: início.bat' (Det $c.Itens 'Scripts') 'scripts.ini em UTF-16'
        Igual $true ((Det $c.Itens 'Mapeamento de unidades') -like '*\\srv\Usuários\%USERNAME%*') 'XML UTF-8 sem BOM com acento'
        Igual 'User\Preferences\Groups\Groups.xml' (@($c.SenhaEm) -join ',') 'arquivo com senha'
        Igual 1 @($c.Erros).Count 'um erro'
        Igual $true ((@($c.Erros) -join '') -like 'User\Registry.pol:*') 'Registry.pol inválido não derruba o resto'
    } finally { Remove-Item -LiteralPath $dir -Recurse -Force }
}
Teste 'GPOs: o que cada GPO faz, senha exposta e CSV de configurações' {
    . $amb
    $c1 = Obj @{ Itens = @((NovoItemGpo 'Computador' 'seguranca' 'Política de contas (senha e bloqueio)' @('Tamanho mínimo da senha: 6')),
        (NovoItemGpo 'Usuário' 'pref:Drives' 'Mapeamento de unidades' @(1..20 | ForEach-Object { "X$_" }))); SenhaEm = @('User\Preferences\Groups\Groups.xml'); Erros = @() }
    $gpos = @((Gpo '{21}' 'Mapeamentos' @{ Conteudo = $c1; Flags = 1 }), (Gpo '{22}' 'Vazia' @{ Versao = 0; Conteudo = (Obj @{ Itens = @(); SenhaEm = @(); Erros = @() }) }),
        (Gpo '{23}' 'Sem leitura' @{ ExtMaquina = '[{42B5FAAE-6536-11D2-AE5A-0000F87571E3}{40B6664F-4972-11D1-A7CA-0000F87571E3}]' }))
    $vinc = @((Alvo 'OU=TI,DC=empresa,DC=local' "$(Lnk '{21}' 0)$(Lnk '{23}' 0)" $null))
    $r = Get-AchadosGpo $gpos $vinc $domDN $true
    Igual '{21}' (IdsGpo $r 'SenhaPreferencia') 'SenhaPreferencia'
    $cards = @(CardsGpo $r)
    Igual $true ($cards[0] -like '<div class="card sev-alta">*<h3>1 GPO com senha exposta no SYSVOL</h3>*<td>Mapeamentos</td><td>User\Preferences\Groups\Groups.xml</td>*') 'card alta'
    $sec = SecaoGpo $r 7
    Igual $true ($sec -like '*O que cada GPO faz*<div class="gpo-bloco" id="gpo-21"><h4 class="gpo">Mapeamentos<span class="tag t-bad">senha exposta</span>*') 'bloco da GPO com tag'
    Igual $true ($sec -like '*<td>Usuário<span class="tag t-warn">parte desativada</span></td>*') 'parte Usuário desativada (flags 1)'
    Igual $true ($sec -like '*<li>X15</li><li>+5 no arquivo 06_GPOs_Configuracoes</li>*') 'limite de 15 linhas'
    Igual $true ($sec -like '*só vale para as contas do domínio quando a GPO está vinculada na raiz*') 'aviso da política de senha fora da raiz'
    Igual $true ($sec -like '*id="gpo-22">*Nenhuma configuração encontrada: a GPO não faz nada.*') 'GPO vazia'
    Igual $true ($sec -like '*id="gpo-23">*<td>Computador</td><td>Scripts</td><td><i>área configurada; detalhes não lidos (SYSVOL inacessível)</i></td>*') 'área só pela extensão'
    $semAviso = SecaoGpo (Get-AchadosGpo @(Gpo '{21}' 'Mapeamentos' @{ Conteudo = $c1 }) @(Alvo $domDN (Lnk '{21}' 0) $null) $domDN $true) 7
    Igual $false ($semAviso -like '*vinculada na raiz*') 'vinculada na raiz: sem aviso'

    $dir = NovoDir
    try {
        Write-CsvGpoConfig $r (Join-Path $dir 'c.csv')
        $csv = @(Import-Csv -LiteralPath $r.CsvConfig -Delimiter ';')
        Igual 'GPO,GUID,Parte,Area,Configuracao' (($csv[0].PSObject.Properties | ForEach-Object { $_.Name }) -join ',') 'colunas do 06_GPOs_Configuracoes'
        Igual 22 $csv.Count 'todas as linhas, sem limite'
        Igual 'Sem leitura|Computador|Scripts|(área configurada; SYSVOL não lido)' "$($csv[21].GPO)|$($csv[21].Parte)|$($csv[21].Area)|$($csv[21].Configuracao)" 'área sem detalhe'
        Write-CsvGpo $r (Join-Path $dir '06.csv')
        $l = @(Import-Csv -LiteralPath $r.Csv -Delimiter ';') | Where-Object { $_.GUID -eq '{21}' }
        Igual 'Computador: Política de contas (senha e bloqueio) (1) | Usuário: Mapeamento de unidades (20)' $l.Configuracoes 'resumo no 06'
        Igual $true ($l.Achados -like 'Senha gravada em Prefer*ncias (cpassword)*') 'achado no 06'
    } finally { Remove-Item -LiteralPath $dir -Recurse -Force }
}

# ---- Fase 6: computadores ----
# Computador como o Get-ADComputer devolve (padrão: ativo, logou ontem, criado há 1000 dias)
function Pc([string]$nome, [string]$so, [string]$ver, [hashtable]$h) {
    $o = @{ Name = $nome; DistinguishedName = "CN=$nome,OU=Estações,DC=empresa,DC=local"; Enabled = $true; OperatingSystem = $so
        OperatingSystemVersion = $ver; LastLogonDate = $h3.AddDays(-1); whenCreated = $h3.AddDays(-1000) }
    if ($h) { foreach ($k in @($h.Keys)) { $o[$k] = $h[$k] } }
    Obj $o
}
function Fim([string]$so, [string]$ver) { $f = Get-FimSuporte $so $ver; "$($f.Familia)|$(if ($f.Fim) { $f.Fim.ToString('yyyy-MM-dd') })" }
function Nomes($lista) { (@($lista) | ForEach-Object { $_.Nome }) -join ',' }

Teste 'Get-FimSuporte: uma linha por regra' {
    $casos = @(
        ,@('Windows 2000 Professional', '5.0 (2195)', 'Windows 2000|2010-07-13')
        ,@('Windows XP Professional', '5.1 (2600)', 'Windows XP|2014-04-08')
        ,@('Windows Vista™ Business', '6.0 (6002)', 'Windows Vista|2017-04-11')
        ,@('Windows 7 Professional', '6.1 (7601)', 'Windows 7|2020-01-14')
        ,@('Windows 8.1 Pro', '6.3 (9600)', 'Windows 8.1|2023-01-10')
        ,@('Windows 8 Pro', '6.2 (9200)', 'Windows 8|2016-01-12')
        ,@('Windows 10 Enterprise 2015 LTSB', '10.0 (10240)', 'Windows 10 LTSB 2015|2025-10-14')
        ,@('Windows 10 Enterprise 2016 LTSB', '10.0 (14393)', 'Windows 10 LTSB 2016|2026-10-13')
        ,@('Windows 10 Enterprise LTSC', '10.0 (17763)', 'Windows 10 LTSC 2019|2029-01-09')
        ,@('Windows 10 Enterprise LTSC', '10.0 (19044)', 'Windows 10 LTSC 2021|2027-01-12')
        ,@('Windows 10 Pro', '10.0 (19045)', 'Windows 10|2025-10-14')
        ,@('Windows 10 Enterprise LTSC', '10.0 (19045)', 'Windows 10|2025-10-14')
        ,@('Windows 11 Enterprise LTSC', '10.0 (26100)', 'Windows 11 LTSC 2024|2029-10-09')
        ,@('Windows 11 Pro', '10.0 (22000)', 'Windows 11 21H2|2023-10-10')
        ,@('Windows 11 Pro', '10.0 (22621)', 'Windows 11 22H2|2024-10-08')
        ,@('Windows 11 Pro', '10.0 (22631)', 'Windows 11 23H2|2025-11-11')
        ,@('Windows 11 Pro', '10.0 (26100)', 'Windows 11 24H2|2026-10-13')
        ,@('Windows 11 Enterprise', '10.0 (26200)', 'Windows 11 25H2|2027-10-12')
        ,@('Windows 11 Pro', '10.0 (26300)', 'Windows 11 26H2|2028-10-09')
        ,@('Windows 11 Pro', '10.0 (28000)', 'Windows 11 26H1|2028-03-14')
        ,@('Windows 11 Pro', '10.0 (29500)', 'Windows 11 (mais nova que a tabela)|')
        ,@('Windows 11 Pro', '10.0 (25000)', 'Windows 11 (build 25000)|')
        ,@('Windows Server 2003', '5.2 (3790)', 'Windows Server 2003|2015-07-14')
        ,@('Windows Server® 2008 Standard', '6.0 (6002)', 'Windows Server 2008|2020-01-14')
        ,@('Windows Server 2008 R2 Standard', '6.1 (7601)', 'Windows Server 2008 R2|2020-01-14')
        ,@('Windows Server 2012 R2 Datacenter', '6.3 (9600)', 'Windows Server 2012 R2|2023-10-10')
        ,@('Windows Server 2016 Standard', '10.0 (14393)', 'Windows Server 2016|2027-01-12')
        ,@('Windows Server 2019 Standard', '10.0 (17763)', 'Windows Server 2019|2029-01-09')
        ,@('Windows Server 2022 Standard', '10.0 (20348)', 'Windows Server 2022|2031-10-14')
        ,@('Windows Server 2025 Datacenter', '10.0 (26100)', 'Windows Server 2025|2034-11-14')
        ,@('Windows Server Datacenter', '10.0 (20348)', 'Windows Server 2022|2031-10-14')
        ,@('Windows Server Standard', '10.0 (25398)', 'Windows Server 23H2 (canal anual)|2026-05-12')
        ,@('Windows Server Datacenter', '10.0 (19042)', 'Windows Server (canal semestral)|2022-08-09')
        ,@('Windows Server Datacenter', '10.0 (26500)', 'Windows Server (mais nova que a tabela)|')
        ,@('Windows Embedded Standard', '6.1 (7601)', 'Windows Embedded Standard|')
        ,@('Linux', '', 'Não Windows / desconhecido|')
        ,@('', '', 'Não Windows / desconhecido|')
    )
    foreach ($c in @($casos)) { Igual $c[2] (Fim $c[0] $c[1]) "$($c[0]) / $($c[1])" }
    Igual $Win11Obs (Get-FimSuporte 'Windows 11 Pro' '10.0 (26100)').Obs 'aviso Home/Pro'
    Igual '' (Get-FimSuporte 'Windows 11 Enterprise LTSC' '10.0 (26100)').Obs 'LTSC sem aviso'
}
Teste 'computadores: fora de suporte, sem uso e desativados' {
    $comps = @((Pc 'PC-W10' 'Windows 10 Pro' '10.0 (19045)'), (Pc 'PC-W11' 'Windows 11 Pro' '10.0 (26200)'),
        (Pc 'LTSB' 'Windows 10 Enterprise 2016 LTSB' '10.0 (14393)'),
        (Pc 'OFF7' 'Windows 7 Professional' '6.1 (7601)' @{ Enabled = $false; LastLogonDate = $h3.AddDays(-900) }),
        (Pc 'EXATO' 'Windows 11 Pro' '10.0 (26200)' @{ LastLogonDate = $h3.AddDays(-90) }),
        (Pc 'PARADO' 'Windows 11 Pro' '10.0 (26200)' @{ LastLogonDate = $h3.AddDays(-91) }),
        (Pc 'NUNCA' 'Windows 11 Pro' '10.0 (26200)' @{ LastLogonDate = $null; whenCreated = $h3.AddDays(-91) }),
        (Pc 'NOVO' 'Windows 11 Pro' '10.0 (26200)' @{ LastLogonDate = $null; whenCreated = $h3.AddDays(-5) }),
        (Pc 'NAS' '' ''),
        (Pc 'VELHO' 'Windows 10 Pro' '10.0 (19045)' @{ LastLogonDate = $h3.AddDays(-300) }),
        (Pc 'LIMITE' 'Windows 11 Pro' '10.0 (26200)' @{ LastLogonDate = $h3.AddMonths(-7) }),
        (Pc 'PASSOU' 'Windows 11 Pro' '10.0 (26200)' @{ LastLogonDate = $h3.AddMonths(-7).AddMinutes(-1) }),
        (Pc 'NUNCA2' 'Windows 11 Pro' '10.0 (26200)' @{ LastLogonDate = $null; whenCreated = $h3.AddDays(-400) }), $null)
    $r = Get-AchadosComputadores $comps $h3 90
    Igual 13 @($r.Computadores).Count 'todos'
    Igual 'PC-W10,LTSB' (Nomes $r.ForaSuporte) 'ForaSuporte (desativado não, sem logon há 7 meses não; ordem por data)'
    Igual 'OFF7,NUNCA2,VELHO,PASSOU' (Nomes $r.Antigos) 'Antigos (desativado entra, exatamente 7 meses não; do mais antigo)'
    Igual 'NUNCA,LIMITE,PARADO' (Nomes $r.Inativos) 'Inativos (exatamente 90 dias não, recém-criado não, antigos não)'
    Igual 1 $r.Desativados 'Desativados'
    Igual 'Windows 7:1/0/0/1/0,Windows 10:2/2/0/1/1,Windows 10 LTSB 2016:1/1/0/0/1,Windows 11 25H2:8/8/3/2/0,Não Windows / desconhecido:1/1/0/0/0' ((@($r.Familias) | ForEach-Object { "$($_.Familia):$($_.Total)/$($_.Ativos)/$($_.Inativos)/$($_.Antigos)/$($_.ForaSuporte)" }) -join ',') 'resumo por família'
    Igual 'OU=Estações,DC=empresa,DC=local' (@($r.Computadores) | Where-Object { $_.Nome -eq 'NAS' }).OU 'OU'
    Igual '300|' "$((@($r.Computadores) | Where-Object { $_.Nome -eq 'VELHO' }).DiasSemLogon)|$((@($r.Computadores) | Where-Object { $_.Nome -eq 'NUNCA' }).DiasSemLogon)" 'DiasSemLogon (nunca logou = vazio)'
    $cards = @(CardsComputadores $r)
    Igual $true ($cards[1] -like '<div class="card sev-media">*<h3>4 computadores sem logon há mais de 7 meses (candidatos à remoção)</h3>*3 ainda ativas, 1 já desativada*') 'card candidatos à remoção'
    Igual $true ($cards[3] -like '*1 delas está na lista de candidatos à remoção*') 'card desativados cita a lista'
    # Fronteira: no último dia com suporte ainda não está fora
    Igual '' (Nomes (Get-AchadosComputadores @(Pc 'LTSB' 'Windows 10 Enterprise 2016 LTSB' '10.0 (14393)') ([datetime]'2026-10-13 23:00') 90).ForaSuporte) 'último dia'
    Igual 'LTSB' (Nomes (Get-AchadosComputadores @(Pc 'LTSB' 'Windows 10 Enterprise 2016 LTSB' '10.0 (14393)') ([datetime]'2026-10-14 00:30') 90).ForaSuporte) 'dia seguinte'
}
Teste 'computadores: cards, seção, CSV e snapshot' {
    . $amb
    $comps = @((Pc 'PC-W10' 'Windows 10 Pro' '10.0 (19045)'), (Pc 'PC-W11' 'Windows 11 Pro' '10.0 (26200)'),
        (Pc 'PARADO' 'Windows 11 Pro' '10.0 (26200)' @{ LastLogonDate = $h3.AddDays(-200) }),
        (Pc 'OFF' 'Windows 7 Professional' '6.1 (7601)' @{ Enabled = $false }),
        (Pc 'VELHO' 'Windows 11 Pro' '10.0 (26200)' @{ LastLogonDate = $h3.AddDays(-300) }))
    $r = Get-AchadosComputadores $comps $h3 90
    $cards = @(CardsComputadores $r)
    Igual 4 $cards.Count 'quatro cards'
    Igual $true ($cards[0] -like '<div class="card sev-alta">*<h3>1 computador com sistema fora de suporte</h3>*<td>Windows 10</td><td class="num">1</td><td class="num">14/10/2025</td>*') 'card fora de suporte'
    Igual $true ($cards[1] -like '<div class="card sev-media">*<h3>1 computador sem logon há mais de 7 meses (candidatos à remoção)</h3>*') 'card candidatos à remoção'
    Igual $true ($cards[2] -like '<div class="card sev-baixa">*<h3>1 computador ativo sem uso há mais de 90 dias (e menos de 7 meses)</h3>*') 'card sem uso'
    Igual $true ($cards[3] -like '<div class="card sev-baixa">*<h3>1 conta de computador desativada</h3>*') 'card desativados'
    Igual $false ($cards[3] -like '*candidatos à remoção*') 'card desativados sem citar a lista'
    Igual 0 @(CardsComputadores (Get-AchadosComputadores @(Pc 'PC-W11' 'Windows 11 Pro' '10.0 (26200)') $h3 90)).Count 'sem achados: sem card'
    $sec = SecaoComputadores $r 10
    Igual $true ($sec -like '<section id="mod-computadores"><h2><span>10</span>Computadores</h2>*') 'título'
    Igual $true ($sec -like '*<td>Windows 11 25H2</td><td class="num">3</td><td class="num">3</td><td class="num">1</td><td class="num sev-media">1</td><td class="num">0</td><td class="num">12/10/2027*</td>*') 'linha do resumo com *'
    Igual $true ($sec -like '*Sem logon há mais de 7 meses (candidatos à remoção)</h3>*<td>VELHO</td><td>ativo</td><td>Windows 11 Pro</td><td class="num">31/12/2025</td><td class="num">300</td>*') 'lista candidatos à remoção'
    Igual $false ($sec -like '*Sem uso há mais de 90 dias*<td>VELHO</td>*') 'candidato fora da lista sem uso'
    Igual $true ($sec -like '*<td class="num sev-alta">1</td>*') 'fora de suporte em destaque'
    Igual $true ($sec -like '** Windows 11: datas de Home/Pro*') 'nota do Windows 11'
    Igual $true ($sec -like '*<td>PC-W10</td><td>Windows 10 Pro</td><td class="num">14/10/2025</td><td class="num">26/10/2026</td><td>empresa.local/Esta*es</td>*') 'lista fora de suporte'
    Igual $true ($sec -like '*Sem uso há mais de 90 dias*<td>PARADO</td>*') 'lista sem uso'
    Igual $true ((SecaoComputadores (Get-AchadosComputadores @() $h3 90) 6) -like '*Nenhum computador encontrado.*') 'sem computadores'

    # Listas limitadas a 300 linhas
    $muitos = @(); for ($i = 0; $i -lt 305; $i++) { $muitos += Pc ('XP{0:000}' -f $i) 'Windows XP Professional' '5.1 (2600)' }
    $sec2 = SecaoComputadores (Get-AchadosComputadores $muitos $h3 90) 6
    Igual 300 ([regex]::Matches($sec2, '<td>XP\d{3}</td>')).Count 'até 300 linhas'
    Igual $true ($sec2 -like '*<td colspan="5">+5 no arquivo 07_Computadores</td>*') '+N no arquivo'

    $dir = NovoDir
    try {
        Write-CsvComputadores $r (Join-Path $dir '07.csv')
        $csv = @(Import-Csv -LiteralPath $r.Csv -Delimiter ';')
        Igual 'Nome,SO,Versao,Ativo,UltimoLogon,DiasSemLogon,CriadoEm,Familia,FimSuporte,ForaDeSuporte,Inativo,CandidatoRemocao,OU' (($csv[0].PSObject.Properties | ForEach-Object { $_.Name }) -join ',') 'colunas do 07'
        Igual 5 $csv.Count 'todos no 07'
        $l = @($csv) | Where-Object { $_.Nome -eq 'PC-W10' }
        Igual 'Windows 10|True|False|False|1|OU=Estações,DC=empresa,DC=local' "$($l.Familia)|$($l.ForaDeSuporte)|$($l.Inativo)|$($l.CandidatoRemocao)|$($l.DiasSemLogon)|$($l.OU)" 'linha com acento'
        $l = @($csv) | Where-Object { $_.Nome -eq 'VELHO' }
        Igual 'True|300' "$($l.CandidatoRemocao)|$($l.DiasSemLogon)" 'linha do candidato à remoção'
    } finally { Remove-Item -LiteralPath $dir -Recurse -Force }

    $MOD = @{ Marcados = @('Computadores'); Comp = $r }
    $s = ConvertTo-Snapshot @(Get-LinhasSnapshot $MOD)
    Igual 'Computadores' $s.Meta['Modulos'] 'Modulos'
    Igual '1|Windows 10 Pro' $s.Dados['Computador']['PC-W10'] 'linha ativo'
    Igual '0|Windows 7 Professional' $s.Dados['Computador']['OFF'] 'linha desativado'
    Igual 4 @(CardsModulos $MOD).Count 'CardsModulos'
    Igual $true ((SecoesModulos $MOD) -like '<section id="mod-computadores"><h2><span>6</span>*') 'SecoesModulos'
}
Teste 'computadores: 5.000 objetos em menos de 10 s e HTML abaixo de 1 MB' {
    $sos = @(@('Windows 10 Pro', '10.0 (19045)'), @('Windows 11 Pro', '10.0 (26200)'), @('Windows 7 Professional', '6.1 (7601)'),
        @('Windows Server 2019 Standard', '10.0 (17763)'), @('Windows XP Professional', '5.1 (2600)'))
    $comps = New-Object System.Collections.ArrayList
    for ($i = 0; $i -lt 5000; $i++) {
        $so = $sos[$i % $sos.Count]
        $h = @{ LastLogonDate = $h3.AddDays(-($i % 200)); Enabled = ($i % 17 -ne 0) }
        [void]$comps.Add((Pc ('PC{0:0000}' -f $i) $so[0] $so[1] $h))
    }
    $sw = [Diagnostics.Stopwatch]::StartNew()
    $r = Get-AchadosComputadores $comps $h3 90
    $html = (@(CardsComputadores $r) -join '') + (SecaoComputadores $r 6)
    $sw.Stop()
    Igual 5000 @($r.Computadores).Count 'todos'
    if ($sw.Elapsed.TotalSeconds -ge 10) { throw "demorou $([math]::Round($sw.Elapsed.TotalSeconds, 1)) s" }
    $bytes = [Text.Encoding]::UTF8.GetByteCount($html)
    if ($bytes -ge 1MB) { throw "HTML com $bytes bytes" }
    Write-Host "          ($([math]::Round($sw.Elapsed.TotalSeconds, 1)) s, HTML $([math]::Round($bytes / 1KB)) KB)"
}

Write-Host ''
if ($script:falhas) {
    Write-Host "REPROVADO: $($script:falhas) de $($script:total) teste(s) falharam" -ForegroundColor Red
    exit 1
}
Write-Host "OK: $($script:total) teste(s)" -ForegroundColor Green
exit 0
