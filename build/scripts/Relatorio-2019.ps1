#Requires -Version 5.1
<#
.SYNOPSIS
  Relatório de usuários ativos do AD + permissões de compartilhamento e NTFS,
  com geração automática de relatório HTML para apresentação.
  Testado para Windows Server 2016 / 2019 / 2022 (PowerShell 5.1).

.EXEMPLOS
  # Todos os compartilhamentos do servidor, 2 níveis de subpastas
  .\Relatorio-AD-Permissoes-2022.ps1 -Cliente "Delta Frio"

  # Caminhos específicos (sem compartilhamento), 3 níveis
  .\Relatorio-AD-Permissoes-2022.ps1 -Cliente "Delta Frio" -Caminhos "D:\Dados","E:\Setores" -Profundidade 3

  # Abrir o relatório no navegador ao final
  .\Relatorio-AD-Permissoes-2022.ps1 -Cliente "Delta Frio" -Abrir
#>
[CmdletBinding()]
param(
    [string]$Cliente,
    [string]$Responsavel = 'Compumaq · Werner Informática',
    [int]$Profundidade = 2,
    [string]$Saida = 'C:\Relatorio_AD',
    [string[]]$Caminhos,
    [int]$DiasInativo = 90,
    # Regex aplicada ao caminho da OU para marcar contas de terceiros
    [string]$PadraoTerceiros = 'Terceir|Tercer|Fornecedor|Externo',
    # Regex aplicada ao nome do compartilhamento para destacar pastas sensíveis
    [string]$PastasSensiveis = '^(rh|pessoal|dp|financeiro|contabil|contabilidade|diretoria|juridico|folha)',
    [switch]$Abrir,
    # Versão do Windows Server (gravada pelo build do RelatorioAD)
    [string]$VersaoSO = '2019'
)

# powershell.exe -File não passa arrays: aceita também "C:\A|D:\B"
if ($Caminhos -and @($Caminhos).Count -eq 1 -and $Caminhos[0] -like '*|*') {
    $Caminhos = @($Caminhos[0] -split '\|' | ForEach-Object { $_.Trim() } | Where-Object { $_ })
}

$ErrorActionPreference = 'Stop'
Import-Module ActiveDirectory
Import-Module SmbShare -ErrorAction SilentlyContinue

$hoje   = Get-Date
$stamp  = $hoje.ToString('yyyyMMdd_HHmm')
$dom    = Get-ADDomain
$domSid = $dom.DomainSID.Value
if (-not $Cliente) { $Cliente = $dom.NetBIOSName }
New-Item -ItemType Directory -Force -Path $Saida | Out-Null
$csvOpts = @{ NoTypeInformation = $true; Encoding = 'UTF8'; Delimiter = ';' }

function HE([object]$s) { [System.Net.WebUtility]::HtmlEncode([string]$s) }
function Curto([string]$n) { if ($n -match '\\') { ($n -split '\\', 2)[1] } else { $n } }

# =====================================================================
# 1. USUÁRIOS ATIVOS
# =====================================================================
Write-Host '[1/5] Lendo usuários ativos do AD...' -ForegroundColor Cyan
$adUsers = Get-ADUser -Filter 'Enabled -eq $true' -Properties DisplayName, Mail, Department, Title, Description, LastLogonDate, PasswordLastSet, PasswordNeverExpires, whenCreated, MemberOf

$U = @{}          # login -> objeto
$SidToSam = @{}
foreach ($a in $adUsers) {
    $parent = ($a.DistinguishedName -split '(?<!\\),', 2)[1]
    if     ($parent -match '^OU=([^,]+)')  { $setor = $Matches[1] }
    elseif ($parent -match '^CN=Users,')   { $setor = 'Contas padrão (Users)' }
    else                                   { $setor = 'Raiz do domínio' }
    $grupos = @($a.MemberOf | ForEach-Object { (($_ -split '(?<!\\),')[0] -replace '^CN=') -replace '\\,', ',' } | Sort-Object)
    $nome = if ($a.DisplayName) { $a.DisplayName } else { $a.SamAccountName }
    $U[$a.SamAccountName] = [pscustomobject]@{
        Login = $a.SamAccountName; Nome = $nome; Setor = $setor; OU = $parent
        Mail = $a.Mail; Depto = $a.Department; Cargo = $a.Title; Desc = $a.Description
        Logon = $a.LastLogonDate; SenhaEm = $a.PasswordLastSet; Nunca = [bool]$a.PasswordNeverExpires
        Criado = $a.whenCreated; Grupos = $grupos; Terceiro = ($parent -match $PadraoTerceiros)
    }
    $SidToSam[$a.SID.Value] = $a.SamAccountName
}
$allUsers = @($U.Keys)

$U.Values | Sort-Object Login | Select-Object `
    @{n='Login';e={$_.Login}}, @{n='Nome';e={$_.Nome}}, @{n='Email';e={$_.Mail}},
    @{n='Departamento';e={$_.Depto}}, @{n='Cargo';e={$_.Cargo}}, @{n='Descricao';e={$_.Desc}},
    @{n='UltimoLogon';e={$_.Logon}}, @{n='SenhaAlteradaEm';e={$_.SenhaEm}},
    @{n='SenhaNuncaExpira';e={$_.Nunca}}, @{n='CriadoEm';e={$_.Criado}},
    @{n='OU';e={$_.OU}}, @{n='Grupos';e={$_.Grupos -join ', '}} |
    Export-Csv "$Saida\01_Usuarios_Ativos_$stamp.csv" @csvOpts

# =====================================================================
# Resolução de contas/grupos (por SID - independe do idioma do Windows)
# =====================================================================
$TodosSids   = @('S-1-1-0', 'S-1-5-11', 'S-1-5-32-545', "$domSid-513")   # Todos, Usuários autenticados, BUILTIN\Usuários, Usuários do domínio
$IgnorarSids = @('S-1-5-18', 'S-1-5-19', 'S-1-5-20', 'S-1-3-0', 'S-1-3-1', 'S-1-5-32-549', 'S-1-5-32-550')
$SidAdmins   = 'S-1-5-32-544'
$cacheP = @{}

function Get-NomeSid([string]$sid) {
    try { (New-Object System.Security.Principal.SecurityIdentifier($sid)).Translate([System.Security.Principal.NTAccount]).Value }
    catch { "SID órfão ($sid)" }
}

function Expand-GrupoLocal([string]$sid) {
    # Para servidor membro (não-DC): grupos BUILTIN são locais
    $res = New-Object 'System.Collections.Generic.HashSet[string]' ([StringComparer]::OrdinalIgnoreCase)
    $todos = $false
    try {
        $nome = Curto (Get-NomeSid $sid)
        $g = [ADSI]"WinNT://$env:COMPUTERNAME/$nome,group"
        foreach ($m in @($g.Invoke('Members'))) {
            try {
                $bytes = $m.GetType().InvokeMember('objectSid', 'GetProperty', $null, $m, $null)
                $msid  = (New-Object System.Security.Principal.SecurityIdentifier($bytes, 0)).Value
                $r = Resolve-Principal $msid
                if ($r.Todos) { $todos = $true }
                foreach ($x in $r.Users) { [void]$res.Add($x) }
            } catch {}
        }
    } catch {}
    [pscustomobject]@{ Todos = $todos; Users = @($res) }
}

function Resolve-Principal([string]$sid) {
    if ($cacheP.ContainsKey($sid)) { return $cacheP[$sid] }
    $nome = Get-NomeSid $sid
    $p = [pscustomobject]@{ Sid = $sid; Nome = $nome; Curto = (Curto $nome); Tipo = 'outro'; Todos = $false; Users = @() }
    $cacheP[$sid] = $p
    if ($TodosSids -contains $sid) { $p.Tipo = 'todos'; $p.Todos = $true; return $p }
    if ($SidToSam.ContainsKey($sid)) { $p.Tipo = 'usuario'; $p.Users = @($SidToSam[$sid]); return $p }
    $o = $null
    try { $o = Get-ADObject -LDAPFilter "(objectSid=$sid)" -ErrorAction Stop } catch {}
    if ($o -and $o.ObjectClass -eq 'user')  { $p.Tipo = 'usuario'; return $p }   # usuário desabilitado
    if ($o -and $o.ObjectClass -eq 'group') {
        $p.Tipo = 'grupo'
        try {
            $p.Users = @(Get-ADGroupMember -Identity $o.DistinguishedName -Recursive -ErrorAction Stop |
                Where-Object { $_.objectClass -eq 'user' -and $U.ContainsKey($_.SamAccountName) } |
                Select-Object -ExpandProperty SamAccountName -Unique)
        } catch {
            # Fallback (grupos grandes ou com membros externos): busca em cadeia via LDAP
            $dnEsc = $o.DistinguishedName -replace '\(', '\28' -replace '\)', '\29'
            $p.Users = @(Get-ADUser -LDAPFilter "(memberOf:1.2.840.113556.1.4.1941:=$dnEsc)" |
                Where-Object { $U.ContainsKey($_.SamAccountName) } | Select-Object -ExpandProperty SamAccountName)
        }
        return $p
    }
    if ($sid -like 'S-1-5-32-*') {
        $p.Tipo = 'grupo'
        $r = Expand-GrupoLocal $sid
        $p.Users = $r.Users
        if ($r.Todos) { $p.Tipo = 'todos'; $p.Todos = $true }
    }
    return $p
}
function UsersOf($p) { if ($p.Todos) { $allUsers } else { $p.Users } }

# Nível de acesso a partir da máscara NTFS (inclui direitos genéricos)
function Get-Nivel([uint32]$v) {
    if (($v -band 268435456) -or (($v -band 2032127) -eq 2032127)) { return 3 }
    if (($v -band 1073741824) -or ($v -band 6)) { return 2 }
    return 1
}
function Get-Rotulo([uint32]$v) {
    if (($v -band 268435456) -or (($v -band 2032127) -eq 2032127)) { return 'Controle total' }
    if (($v -band 1073741824) -or ($v -band 6)) {
        if ($v -band 65536) { return 'Modificar' } else { return 'Leitura e gravação' }
    }
    if (($v -band 2147483648) -or ($v -band 1)) { return 'Leitura' }
    return 'Especial'
}
# Efeito de um "Negar": 0 = bloqueia tudo, 1 = bloqueia gravação, 9 = irrelevante
function Get-EfeitoNegar([uint32]$v) {
    if (($v -band 268435456) -or ($v -band 2147483648) -or ($v -band 1)) { return 0 }
    if (($v -band 1073741824) -or ($v -band 6) -or ($v -band 65536)) { return 1 }
    return 9
}

# =====================================================================
# 2. COMPARTILHAMENTOS
# =====================================================================
Write-Host '[2/5] Lendo compartilhamentos...' -ForegroundColor Cyan
$SLV = @{ Full = 3; Change = 2; Read = 1; Custom = 1 }
$alvos = New-Object System.Collections.ArrayList
if ($Caminhos) {
    foreach ($c in $Caminhos) { [void]$alvos.Add([pscustomobject]@{ Nome = (Split-Path $c -Leaf); Path = $c.TrimEnd('\'); Acl = @(); Manual = $true }) }
} else {
    $shares = Get-SmbShare | Where-Object {
        -not $_.Special -and $_.Path -and $_.ShareType -eq 'FileSystemDirectory' -and $_.Name -notin @('NETLOGON', 'SYSVOL', 'print$')
    } | Sort-Object @{e={$_.Name.EndsWith('$')}}, Name
    foreach ($s in $shares) {
        $acl = @(Get-SmbShareAccess -Name $s.Name | ForEach-Object {
            $sid = $null
            try { $sid = (New-Object System.Security.Principal.NTAccount($_.AccountName)).Translate([System.Security.Principal.SecurityIdentifier]).Value } catch {}
            [pscustomobject]@{ Conta = $_.AccountName; Sid = $sid; Tipo = [string]$_.AccessControlType; Perm = [string]$_.AccessRight }
        })
        [void]$alvos.Add([pscustomobject]@{ Nome = $s.Name; Path = $s.Path.TrimEnd('\'); Acl = $acl; Manual = $false })
    }
    $alvos | ForEach-Object { $sh = $_; $sh.Acl | Select-Object @{n='Compartilhamento';e={$sh.Nome}}, @{n='Caminho';e={$sh.Path}},
        @{n='Conta';e={$_.Conta}}, @{n='Tipo';e={$_.Tipo}}, @{n='Permissao';e={$_.Perm}} } |
        Export-Csv "$Saida\02_Permissoes_Compartilhamento_$stamp.csv" @csvOpts
}
if (-not $alvos.Count) { throw 'Nenhum compartilhamento encontrado. Use -Caminhos para informar as pastas.' }

# =====================================================================
# 3. PERMISSÕES NTFS
# =====================================================================
Write-Host "[3/5] Lendo permissões NTFS (profundidade $Profundidade)..." -ForegroundColor Cyan
$rows  = New-Object System.Collections.ArrayList
$erros = New-Object System.Collections.ArrayList
$tipoSid = [System.Security.Principal.SecurityIdentifier]
$pastasComAcl = New-Object 'System.Collections.Generic.HashSet[string]'

foreach ($a in $alvos) {
    Write-Host "   -> $($a.Nome): $($a.Path)"
    $lista = @($a.Path)
    if ($Profundidade -gt 0) {
        $ev = $null
        $lista += @(Get-ChildItem -LiteralPath $a.Path -Directory -Recurse -Depth ($Profundidade - 1) -Force -ErrorAction SilentlyContinue -ErrorVariable ev |
                    Select-Object -ExpandProperty FullName)
        foreach ($e in $ev) { [void]$erros.Add([pscustomobject]@{ Pasta = [string]$e.TargetObject; Erro = $e.Exception.Message }) }
    }
    foreach ($pasta in $lista) {
        try { $acl = Get-Acl -LiteralPath $pasta -ErrorAction Stop }
        catch { [void]$erros.Add([pscustomobject]@{ Pasta = $pasta; Erro = $_.Exception.Message }); continue }
        $raiz = ($pasta -eq $a.Path)
        $prot = $acl.AreAccessRulesProtected
        foreach ($ace in $acl.GetAccessRules($true, $true, $tipoSid)) {
            # Subpastas: só o que foi definido nelas (ou tudo, se a herança estiver desabilitada)
            if (-not $raiz -and $ace.IsInherited -and -not $prot) { continue }
            $sid = $ace.IdentityReference.Value
            if ($IgnorarSids -contains $sid) { continue }
            $v = [BitConverter]::ToUInt32([BitConverter]::GetBytes([int32]$ace.FileSystemRights), 0)
            [void]$rows.Add([pscustomobject]@{
                Share = $a.Nome; Pasta = $pasta; Raiz = $raiz; Sid = $sid
                Nivel = (Get-Nivel $v); Rotulo = (Get-Rotulo $v); Negar = (Get-EfeitoNegar $v)
                Permitir = ($ace.AccessControlType -eq 'Allow'); Herdada = $ace.IsInherited; Protegida = $prot
                Bruta = $ace.FileSystemRights.ToString()
            })
            [void]$pastasComAcl.Add($pasta)
        }
    }
}

# =====================================================================
# 4. CÁLCULO DO ACESSO EFETIVO
# =====================================================================
Write-Host '[4/5] Expandindo grupos e calculando acesso efetivo...' -ForegroundColor Cyan
$csv3 = New-Object System.Collections.ArrayList
$Calc = @{}   # share -> dados calculados
foreach ($a in $alvos) {
    $s = $a.Nome
    $sp = @{}; $cap = @{}
    if ($a.Manual) { foreach ($l in $allUsers) { $sp[$l] = 3 } }
    else {
        foreach ($e in $a.Acl) {
            if (-not $e.Sid) { continue }
            $p = Resolve-Principal $e.Sid
            $lv = $SLV[$e.Perm]
            foreach ($l in (UsersOf $p)) {
                if ($e.Tipo -eq 'Deny') { $c = if ($e.Perm -eq 'Change') { 1 } else { 0 }; if (-not $cap.ContainsKey($l) -or $cap[$l] -gt $c) { $cap[$l] = $c } }
                elseif (-not $sp.ContainsKey($l) -or $sp[$l] -lt $lv) { $sp[$l] = $lv }
            }
        }
    }
    $nt = @{}; $orig = @{}; $np = [ordered]@{}; $exc = New-Object System.Collections.ArrayList
    foreach ($r in @($rows | Where-Object { $_.Share -eq $s })) {
        $p = Resolve-Principal $r.Sid
        $origem = if ($p.Tipo -eq 'usuario') { 'Direta' } elseif ($p.Todos) { $p.Curto } else { 'Grupo: ' + $p.Nome }
        # CSV detalhado (usuário a usuário)
        $alvo = if ($p.Todos) { @('(TODOS OS USUARIOS)') } elseif ($p.Users.Count) { $p.Users } elseif ($p.Tipo -eq 'grupo') { @('(grupo sem membros ativos)') } elseif ($p.Tipo -ne 'usuario') { @($p.Nome) } else { @() }
        foreach ($l in $alvo) {
            [void]$csv3.Add([pscustomobject]@{ Compartilhamento = $s; Pasta = $r.Pasta; Usuario = $l
                Nome = $(if ($U.ContainsKey($l)) { $U[$l].Nome } else { '' }); Permissao = $r.Rotulo
                Tipo = $(if ($r.Permitir) { 'Permitir' } else { 'NEGAR' }); Origem = $origem
                Herdada = $(if ($r.Herdada) { 'Sim' } else { 'Nao' }); HerancaPasta = $(if ($r.Protegida) { 'Desabilitada' } else { 'Ativa' })
                PermissaoBruta = $r.Bruta })
        }
        if (-not $r.Raiz) {
            if ($r.Sid -ne $SidAdmins -and $r.Sid -notlike "$domSid-500") { [void]$exc.Add($r) }
            continue
        }
        $k = $r.Sid + '|' + $r.Permitir
        if (-not $np.Contains($k) -or $np[$k].Nivel -lt $r.Nivel) { $np[$k] = $r }
        foreach ($l in (UsersOf $p)) {
            if (-not $r.Permitir) {
                if ($r.Negar -lt 9 -and (-not $cap.ContainsKey($l) -or $cap[$l] -gt $r.Negar)) { $cap[$l] = $r.Negar }
                continue
            }
            if (-not $nt.ContainsKey($l) -or $nt[$l] -lt $r.Nivel) { $nt[$l] = $r.Nivel }
            if (-not $orig.ContainsKey($l)) { $orig[$l] = New-Object 'System.Collections.Generic.SortedSet[string]' }
            [void]$orig[$l].Add($(if ($p.Tipo -eq 'usuario') { 'Direta' } else { $p.Curto }))
        }
    }
    $ef = @{}
    foreach ($l in $allUsers) {
        $n = if ($nt.ContainsKey($l)) { $nt[$l] } else { 0 }
        $m = if ($sp.ContainsKey($l)) { $sp[$l] } else { 0 }
        $e = [Math]::Min($n, $m)
        if ($cap.ContainsKey($l)) { $e = [Math]::Min($e, $cap[$l]) }
        if ($e -gt 0) { $ef[$l] = $e }
    }
    $Calc[$s] = [pscustomobject]@{ SP = $sp; NT = $nt; Orig = $orig; NP = $np; EF = $ef; Exc = $exc }
}
$csv3 | Sort-Object Pasta, Usuario | Export-Csv "$Saida\03_Permissoes_NTFS_por_Usuario_$stamp.csv" @csvOpts
if ($erros.Count) { $erros | Export-Csv "$Saida\99_Erros_Acesso_$stamp.csv" @csvOpts }

# =====================================================================
# 5. RELATÓRIO HTML
# =====================================================================
Write-Host '[5/5] Gerando relatório HTML...' -ForegroundColor Cyan
$TXT = @{ 3 = 'Controle total'; 2 = 'Modificar'; 1 = 'Leitura'; 0 = 'Sem acesso' }
$SPT = @{ Full = 'Controle total'; Change = 'Alteração'; Read = 'Leitura' }
$nSh = $alvos.Count
$nUsers = $allUsers.Count

$ADM  = @((Resolve-Principal $SidAdmins).Users | Sort-Object)
$admSet = @{}; foreach ($l in $ADM) { $admSet[$l] = 1 }
$TERC = @($allUsers | Where-Object { $U[$_].Terceiro })
$limite = $hoje.AddDays(-$DiasInativo)
$inativos = @($allUsers | Where-Object { -not $U[$_].Logon -or $U[$_].Logon -lt $limite } |
              Sort-Object @{e={ if ($U[$_].Logon) { $U[$_].Logon } else { [datetime]::MinValue } }})
$inaSet = @{}; foreach ($l in $inativos) { $inaSet[$l] = 1 }
$nunca = @($allUsers | Where-Object { $U[$_].Nunca })

# Pastas abertas para todos (NTFS) e compartilhamentos "Todos: controle total"
$abertas = @(); $todosFull = @()
foreach ($a in $alvos) {
    $temTodosNtfs = @($Calc[$a.Nome].NP.Values | Where-Object { $_.Permitir -and $TodosSids -contains $_.Sid }).Count -gt 0
    if ($temTodosNtfs) { $abertas += $a.Nome; continue }
    if (@($a.Acl | Where-Object { $_.Tipo -eq 'Allow' -and $_.Perm -eq 'Full' -and $TodosSids -contains $_.Sid }).Count) { $todosFull += $a.Nome }
}
# Permissões diretas a usuários (raiz)
$diretas = [ordered]@{}; $nDiretas = 0
foreach ($a in $alvos) {
    $l = @($Calc[$a.Nome].NP.Values | Where-Object { $_.Permitir -and $SidToSam.ContainsKey($_.Sid) -and $_.Sid -notlike "$domSid-500" } |
           ForEach-Object { $SidToSam[$_.Sid] } | Sort-Object -Unique)
    if ($l.Count) { $diretas[$a.Nome] = $l; $nDiretas += $l.Count }
}
# Possíveis duplicadas (mesmo nome de exibição)
function Norm($n) { ($n.ToLower().Trim() -replace '\s+', ' ') }
$dups = New-Object System.Collections.ArrayList
$lst = @($allUsers | Sort-Object)
for ($i = 0; $i -lt $lst.Count; $i++) {
    for ($j = $i + 1; $j -lt $lst.Count; $j++) {
        $na = Norm $U[$lst[$i]].Nome; $nb = Norm $U[$lst[$j]].Nome
        if (($na -split ' ').Count -lt 2 -or ($nb -split ' ').Count -lt 2) { continue }
        if (($na + $nb) -match '\d$') { continue }
        if ($na -eq $nb -or $nb.StartsWith("$na ") -or $na.StartsWith("$nb ")) { [void]$dups.Add(@($lst[$i], $lst[$j])) }
    }
}
$sensiveis = @($alvos | Where-Object { $_.Nome -match $PastasSensiveis } | ForEach-Object Nome)

# Agrupamento por setor
$setorFim = { param($l) $U[$l].Terceiro -or $U[$l].Setor -in @('Contas padrão (Users)', 'Raiz do domínio') }
$ordUsers = @($allUsers | Sort-Object @{e={ & $setorFim $_ }}, @{e={ $U[$_].Setor }}, @{e={ $U[$_].Nome }})
$setores = [ordered]@{}
foreach ($l in $ordUsers) { $st = $U[$l].Setor; if (-not $setores.Contains($st)) { $setores[$st] = New-Object System.Collections.ArrayList }; [void]$setores[$st].Add($l) }

function NM([string]$l) { if ($U.ContainsKey($l)) { HE $U[$l].Nome } else { HE $l } }
function DT($d) { if ($d) { ([datetime]$d).ToString('dd/MM/yyyy') } else { 'nunca' } }
function Card($sev, $titulo, $corpo) {
    $rot = @{ alta = 'Prioridade alta'; media = 'Prioridade média'; baixa = 'Para validar' }[$sev]
    "<div class=""card sev-$sev""><div class=""sev"">$rot</div><h3>$titulo</h3>$corpo</div>"
}
function NomesPorNome($ls) { (@($ls | Sort-Object { $U[$_].Nome }) | ForEach-Object { NM $_ }) -join ', ' }
function PrincipalHtml($r) {
    $p = Resolve-Principal $r.Sid
    if ($p.Todos) { return "<strong>$(HE $p.Curto)</strong> <span class=""muted"">(todos os usuários)</span>" }
    if ($p.Tipo -eq 'usuario') {
        if ($SidToSam.ContainsKey($p.Sid)) { $l = $SidToSam[$p.Sid]; return "$(NM $l) <span class=""muted"">· $(HE $l)</span>" }
        return "$(HE $p.Curto) <span class=""muted"">(conta desabilitada)</span>"
    }
    if ($p.Tipo -eq 'grupo') {
        $mem = if ($p.Users.Count) { NomesPorNome $p.Users } else { 'sem membros ativos' }
        return "<span class=""grp"">Grupo</span> <strong>$(HE $p.Curto)</strong> <span class=""muted"">($($p.Users.Count))</span><div class=""mem"">$mem</div>"
    }
    return "$(HE $p.Nome) <span class=""muted"">(conta local / fora do AD)</span>"
}

$css = @'
:root{
  --ink:#16202c; --ink2:#4a5666; --mute:#7b8594; --line:#e3e7ec; --bg:#f6f7f9; --paper:#fff;
  --brand:#123a5c; --brand2:#1e5a8a; --accent:#e8a33d;
  --l3:#1f6f4a; --l3bg:#d8efe2; --l2:#1d5f99; --l2bg:#dbe9f7; --l1:#8a6312; --l1bg:#f6ecd2;
  --red:#b3261e; --redbg:#fbe4e1; --amb:#9a5b00; --ambbg:#fdf0d7; --blu:#35567a; --blubg:#e6eef7;
}
*{box-sizing:border-box}
html{scroll-behavior:smooth}
body{margin:0;background:var(--bg);color:var(--ink);font:15px/1.55 "IBM Plex Sans",system-ui,-apple-system,"Segoe UI",Roboto,sans-serif;-webkit-font-smoothing:antialiased}
code,.mono{font-family:"IBM Plex Mono",ui-monospace,Consolas,monospace;font-size:.86em}
.wrap{max-width:1180px;margin:0 auto;padding:0 28px}
.muted{color:var(--mute);font-weight:400}

.cover{background:linear-gradient(135deg,#0d2c47 0%,var(--brand) 55%,var(--brand2) 100%);color:#fff;padding:64px 0 52px}
.kicker{text-transform:uppercase;letter-spacing:.12em;font-size:12px;font-weight:600;color:var(--accent)}
.cover h1{font-size:40px;line-height:1.15;margin:12px 0 32px;font-weight:700;max-width:820px;letter-spacing:-.01em}
.meta{display:grid;grid-template-columns:repeat(4,1fr);gap:18px;border-top:1px solid rgba(255,255,255,.2);padding-top:22px}
.meta div{font-size:14px}
.meta span{display:block;font-size:11px;text-transform:uppercase;letter-spacing:.1em;opacity:.65;margin-bottom:3px}
.meta code{background:rgba(255,255,255,.12);padding:1px 6px;border-radius:4px;font-size:12.5px}

.toc{position:sticky;top:0;z-index:20;background:rgba(255,255,255,.96);backdrop-filter:blur(6px);border-bottom:1px solid var(--line)}
.toc .wrap{display:flex;gap:4px;overflow-x:auto}
.toc a{padding:13px 12px;color:var(--ink2);text-decoration:none;font-size:13.5px;font-weight:500;white-space:nowrap;border-bottom:2px solid transparent}
.toc a:hover{color:var(--brand);border-color:var(--accent)}

section{padding:52px 0 8px}
h2{font-size:26px;margin:0 0 18px;display:flex;align-items:center;gap:14px;letter-spacing:-.01em}
h2 span{display:inline-grid;place-items:center;width:36px;height:36px;border-radius:9px;background:var(--brand);color:#fff;font-size:16px}
h3{margin:0 0 8px;font-size:17px}
h4{margin:0 0 8px;font-size:12px;text-transform:uppercase;letter-spacing:.08em;color:var(--mute)}
.lead{color:var(--ink2);max-width:880px;margin:0 0 14px}

.kpis{display:grid;grid-template-columns:repeat(4,1fr);gap:16px;margin:6px 0 24px}
.kpi{background:var(--paper);border:1px solid var(--line);border-radius:12px;padding:20px 20px 18px;border-top:4px solid var(--brand2)}
.kpi.warn{border-top-color:var(--accent)}
.kpi b{display:block;font-size:40px;line-height:1;font-weight:700;color:var(--brand);margin-bottom:8px}
.kpi.warn b{color:var(--amb)}
.kpi span{color:var(--ink2);font-size:13.5px}

.cards{display:grid;grid-template-columns:1fr 1fr;gap:18px}
.card{background:var(--paper);border:1px solid var(--line);border-radius:12px;padding:20px 22px;border-left:5px solid var(--blu)}
.card p{margin:0 0 10px;color:var(--ink2);font-size:14.5px}
.card ul{margin:0 0 10px;padding-left:20px;font-size:14px}
.card li{margin:2px 0}
.card ul.cols{columns:1}
.sev{font-size:11px;font-weight:700;text-transform:uppercase;letter-spacing:.1em;margin-bottom:6px}
.sev-alta{border-left-color:var(--red)} .sev-alta .sev{color:var(--red)}
.sev-media{border-left-color:var(--accent)} .sev-media .sev{color:var(--amb)}
.sev-baixa .sev{color:var(--blu)}
.sev-alta:first-child{grid-column:1/-1}
.pills{display:flex;flex-wrap:wrap;gap:6px}
.pill{background:var(--redbg);color:var(--red);padding:4px 10px;border-radius:99px;font-size:13px;font-weight:500}
.scroll{max-height:300px;overflow:auto;border:1px solid var(--line);border-radius:8px}
.scroll table.mini{margin:0}

table{border-collapse:collapse;width:100%}
table.mini{font-size:13.5px;margin:4px 0}
table.mini th{text-align:left;font-size:11px;text-transform:uppercase;letter-spacing:.06em;color:var(--mute);font-weight:600;padding:7px 8px;border-bottom:1px solid var(--line);background:#fafbfc;position:sticky;top:0}
table.mini td{padding:6px 8px;border-bottom:1px solid var(--line);vertical-align:top}
.num{text-align:right;font-variant-numeric:tabular-nums;white-space:nowrap}
.nw{white-space:nowrap}

.tag{display:inline-block;font-size:10.5px;font-weight:600;padding:1px 7px;border-radius:5px;background:#eef1f4;color:var(--ink2);margin-left:3px;vertical-align:1px;white-space:nowrap}
.t-adm{background:var(--redbg);color:var(--red)}
.t-terc{background:#ece6f7;color:#5b3c99}
.t-warn{background:var(--ambbg);color:var(--amb)}

.legend{display:flex;align-items:center;gap:10px;flex-wrap:wrap;margin:6px 0 12px;font-size:13px}
.legend label{margin-left:auto;display:flex;gap:6px;align-items:center;color:var(--ink2)}
.legend input#fMat,.search{border:1px solid var(--line);border-radius:8px;padding:7px 11px;font:inherit;font-size:13.5px;min-width:220px;background:#fff}
.search{width:100%;max-width:420px;margin-bottom:12px}
.lv{display:inline-block;padding:2px 9px;border-radius:5px;font-size:12px;font-weight:600}
.lv.l3,.c.l3,.dot.l3{background:var(--l3bg);color:var(--l3)}
.lv.l2,.c.l2,.dot.l2{background:var(--l2bg);color:var(--l2)}
.lv.l1,.c.l1,.dot.l1{background:var(--l1bg);color:var(--l1)}
.lv.l0{background:#f1f3f5;color:var(--mute)}

.matwrap{background:var(--paper);border:1px solid var(--line);border-radius:12px;overflow:auto;max-height:78vh}
table.mat{font-size:12.5px;width:auto;min-width:100%}
table.mat thead th{position:sticky;top:0;background:#fff;z-index:2;border-bottom:1px solid var(--line)}
table.mat th.uh,table.mat td.uh{position:sticky;left:0;background:#fff;z-index:1;text-align:left;min-width:230px;max-width:260px;padding:5px 12px;border-right:1px solid var(--line)}
table.mat thead th.uh{z-index:3;vertical-align:bottom;font-size:11px;text-transform:uppercase;letter-spacing:.06em;color:var(--mute);padding-bottom:10px}
th.rot{vertical-align:bottom;padding:12px 0 10px;width:34px;min-width:34px}
th.rot div{display:flex;justify-content:center}
th.rot span{writing-mode:vertical-rl;transform:rotate(180deg);white-space:nowrap;font-weight:600;font-size:12px;color:var(--ink)}
table.mat thead th.num{vertical-align:bottom;padding:0 10px 10px;font-size:11px;text-transform:uppercase;color:var(--mute)}
.un{display:block;font-weight:500;white-space:nowrap;overflow:hidden;text-overflow:ellipsis}
.ul{font-family:"IBM Plex Mono",monospace;font-size:11px;color:var(--mute)}
td.uh .tag{font-size:9.5px}
td.c{text-align:center;font-weight:700;font-size:11.5px;border:2px solid #fff;padding:4px 0;min-width:34px;cursor:default}
td.c.l0{background:#fafbfc}
table.mat tbody tr:hover td.uh{background:#f3f6fa}
table.mat tbody tr:hover td.c.l0{background:#eef2f6}
table.mat td.num{padding:0 12px;color:var(--ink2)}
tr.grp-row td{background:#eef2f6!important;font-weight:600;font-size:11.5px;text-transform:uppercase;letter-spacing:.07em;color:var(--brand);padding:7px 12px;position:sticky;left:0}
table.mat.noadm tr.adm{display:none}
.note{font-size:12.5px;color:var(--mute);margin-top:10px}

.folders{display:grid;grid-template-columns:1fr;gap:16px}
.folder{background:var(--paper);border:1px solid var(--line);border-radius:12px;padding:18px 22px;break-inside:avoid}
.folder header{display:flex;align-items:baseline;gap:12px;flex-wrap:wrap;margin-bottom:10px}
.folder header h3{margin:0;font-size:19px}
.folder header code{color:var(--mute)}
.fstats{display:flex;gap:18px;font-size:13px;color:var(--ink2);padding:10px 0 12px;border-top:1px solid var(--line);border-bottom:1px solid var(--line);margin-bottom:14px}
.fstats b{font-size:16px;color:var(--ink)}
.dot{display:inline-block;width:10px;height:10px;border-radius:3px;margin-right:6px;vertical-align:-1px}
.fgrid{display:grid;grid-template-columns:1fr 2fr;gap:24px}
.ntfs td:first-child{width:78%}
.grp{font-size:10.5px;font-weight:700;color:var(--brand2);background:var(--blubg);padding:1px 6px;border-radius:4px;margin-right:2px}
.mem{font-size:12.5px;color:var(--ink2);margin-top:3px;line-height:1.45}
.exc{margin-top:12px;padding:10px 14px;background:var(--ambbg);border-radius:8px;font-size:13px}
.exc ul{margin:4px 0 0;padding-left:18px}

table.users{background:var(--paper);border:1px solid var(--line);border-radius:12px;overflow:hidden;font-size:13px}
table.users th{text-align:left;font-size:11px;text-transform:uppercase;letter-spacing:.06em;color:var(--mute);padding:10px;border-bottom:1px solid var(--line);background:#fafbfc}
table.users td{padding:8px 10px;border-bottom:1px solid var(--line);vertical-align:top}
table.users td.grps{color:var(--ink2);font-size:12.5px;max-width:280px}
table.users tbody tr:hover{background:#f7f9fb}

.recs{padding-left:22px;max-width:900px}
.recs li{margin:0 0 12px;padding-left:4px;color:var(--ink2)}
.recs li strong{color:var(--ink)}
.method{margin-top:26px;background:var(--paper);border:1px solid var(--line);border-radius:12px;padding:16px 20px;font-size:13.5px;color:var(--ink2);max-width:900px}
.method p{margin:0}
.foot{padding:40px 28px 50px;color:var(--mute);font-size:12.5px;border-top:1px solid var(--line);margin-top:50px}

@media (max-width:860px){
  .meta,.kpis{grid-template-columns:1fr 1fr}.cards,.fgrid{grid-template-columns:1fr}
  .cover h1{font-size:30px}.wrap{padding:0 16px}
}
@media print{
  @page{size:A4 landscape;margin:12mm}
  body{background:#fff;font-size:12px}
  .toc,.noprint{display:none!important}
  .cover{-webkit-print-color-adjust:exact;print-color-adjust:exact;padding:36px 0}
  *{-webkit-print-color-adjust:exact;print-color-adjust:exact}
  section{padding-top:24px;break-before:page}
  #resumo{break-before:auto}
  .matwrap,.scroll{max-height:none;overflow:visible}
  .card,.folder,.kpi{break-inside:avoid}
  table.mat thead{display:table-header-group}
  tr{break-inside:avoid}
}
'@

$sb = New-Object System.Text.StringBuilder
function A([string]$s) { [void]$sb.Append($s) }

$dataTxt = $hoje.ToString('dd/MM/yyyy')
A "<!doctype html><html lang=""pt-BR""><head><meta charset=""utf-8""><meta name=""viewport"" content=""width=device-width,initial-scale=1"">"
A "<title>Relatório de Acessos — $(HE $Cliente)</title>"
A '<link rel="preconnect" href="https://fonts.googleapis.com"><link href="https://fonts.googleapis.com/css2?family=IBM+Plex+Sans:wght@400;500;600;700&family=IBM+Plex+Mono:wght@400;500&display=swap" rel="stylesheet">'
A "<style>$css</style></head><body>"

# ---- capa ----
$escopo = if ($Caminhos) { "Pastas selecionadas em <code>$(HE $env:COMPUTERNAME)</code>" } else { "$nSh pastas compartilhadas em <code>$(HE $env:COMPUTERNAME)</code>" }
A @"
<header class="cover"><div class="wrap">
<div class="kicker">Auditoria de acessos · Active Directory e servidor de arquivos</div>
<h1>Usuários e permissões de acesso às pastas da rede</h1>
<div class="meta"><div><span>Cliente</span>$(HE $Cliente) · <code>$(HE $dom.DNSRoot)</code></div>
<div><span>Escopo</span>$escopo</div>
<div><span>Levantamento</span>$dataTxt</div>
<div><span>Responsável</span>$(HE $Responsavel)</div></div>
</div></header>
<nav class="toc"><div class="wrap">
<a href="#resumo">1. Resumo</a><a href="#atencao">2. Pontos de atenção</a><a href="#matriz">3. Matriz de acessos</a><a href="#pastas">4. Detalhe por pasta</a><a href="#usuarios">5. Usuários ativos</a>
</div></nav><main class="wrap">
"@

# ---- 1. resumo ----
$nTercAdm = @($ADM | Where-Object { $U[$_].Terceiro }).Count
$destaques = @()
if ($ADM.Count)      { $destaques += "o número de contas com privilégio administrativo" + $(if ($nTercAdm) { " (incluindo $nTercAdm de terceiros)" } else { '' }) }
if ($nDiretas)       { $destaques += 'permissões atribuídas diretamente a usuários em vez de grupos' }
if ($inativos.Count) { $destaques += 'contas sem uso que seguem habilitadas' }
if ($todosFull.Count){ $destaques += 'compartilhamentos liberados para todos' }
$frase = if ($destaques.Count -gt 1) { ($destaques[0..($destaques.Count - 2)] -join ', ') + ' e ' + $destaques[-1] } elseif ($destaques.Count) { $destaques[0] } else { '' }
A @"
<section id="resumo"><h2><span>1</span>Resumo executivo</h2>
<div class="kpis">
<div class="kpi"><b>$nUsers</b><span>contas de usuário ativas no AD</span></div>
<div class="kpi"><b>$nSh</b><span>pastas compartilhadas analisadas</span></div>
<div class="kpi warn"><b>$($ADM.Count)</b><span>contas com privilégio de administrador no servidor</span></div>
<div class="kpi warn"><b>$($inativos.Count)</b><span>contas ativas sem logon há mais de $DiasInativo dias</span></div>
</div>
<p class="lead">O levantamento cruzou três fontes: as contas ativas do Active Directory, as permissões de cada compartilhamento e as permissões NTFS das pastas (com os grupos expandidos até o usuário final). O <strong>acesso efetivo</strong> de cada pessoa é o mais restritivo entre a permissão do compartilhamento e a permissão NTFS — é ele que aparece na matriz da seção 3.</p>
$(if ($frase) { "<p class=""lead"">Os principais pontos a tratar são $frase. Os detalhes estão na seção 2.</p>" })
</section>
"@

# ---- 2. pontos de atenção ----
A '<section id="atencao"><h2><span>2</span>Pontos de atenção</h2><div class="cards">'
$nCards = 0
if ($ADM.Count) {
    $linhas = ($ADM | Sort-Object @{e={ -not $U[$_].Terceiro }}, @{e={ $U[$_].Nome }} | ForEach-Object {
        $l = $_; $n = @($alvos | Where-Object { $Calc[$_.Nome].EF[$l] -eq 3 }).Count
        "<tr><td>$(NM $l)</td><td class=""mono"">$(HE $l)</td><td>$(HE $U[$l].Setor)$(if ($U[$l].Terceiro) { ' <span class="tag t-terc">terceiro</span>' })</td><td class=""num"">$n de $nSh</td></tr>"
    }) -join ''
    $txtTerc = if ($nTercAdm) { " $nTercAdm dessas contas estão em OUs de terceiros." } else { '' }
    A (Card 'alta' "$($ADM.Count) contas são administradoras do servidor" ("<p>Membros do grupo <em>Administradores</em> costumam ter controle total nas pastas e podem alterar permissões e assumir a posse de qualquer arquivo.$txtTerc</p><table class=""mini""><thead><tr><th>Conta</th><th>Login</th><th>Setor</th><th class=""num"">Pastas c/ controle total</th></tr></thead><tbody>$linhas</tbody></table>"))
    $nCards++
}
if ($todosFull.Count) {
    $pills = ($todosFull | ForEach-Object { "<span class=""pill"">$(HE $_)</span>" }) -join ''
    A (Card 'alta' "$($todosFull.Count) compartilhamentos liberam controle total para todos" "<p>Nesses compartilhamentos a única proteção é a permissão NTFS da pasta. Qualquer ajuste errado no NTFS passa a valer para toda a rede.</p><div class=""pills"">$pills</div>")
    $nCards++
}
foreach ($s in $sensiveis) {
    $quem = @($Calc[$s].EF.Keys | Where-Object { -not $admSet.ContainsKey($_) })
    if (-not $quem.Count) { continue }
    $nAdm = @($Calc[$s].EF.Keys | Where-Object { $admSet.ContainsKey($_) }).Count
    $li = ($quem | Sort-Object { $U[$_].Nome } | ForEach-Object { "<li>$(NM $_) <span class=""muted"">· $(HE $_) · $(HE $U[$_].Setor) · $($TXT[$Calc[$s].EF[$_]].ToLower())</span></li>" }) -join ''
    $path = ($alvos | Where-Object Nome -eq $s).Path
    A (Card 'media' "Quem acessa a pasta $(HE $s)" "<p>Pasta com conteúdo potencialmente sensível (<code>$(HE $path)</code>). Além de $nAdm administrador(es), têm acesso:</p><ul>$li</ul><p>Validar com a gestão se todos esses acessos são intencionais.</p>")
    $nCards++
}
if ($nDiretas) {
    $li = ($diretas.Keys | ForEach-Object { "<li><strong>$(HE $_)</strong>: $(NomesPorNome $diretas[$_])</li>" }) -join ''
    A (Card 'media' "$nDiretas permissões dadas diretamente a usuários" "<p>Em vez de usar grupos por setor, várias pastas recebem usuários um a um. Isso dificulta revisar quem tem acesso e faz com que permissões fiquem para trás quando alguém muda de função.</p><ul>$li</ul>")
    $nCards++
}
if ($inativos.Count) {
    $tr = ($inativos | ForEach-Object { "<tr><td>$(NM $_)</td><td class=""mono"">$(HE $_)</td><td>$(HE $U[$_].Setor)</td><td class=""num"">$(DT $U[$_].Logon)</td></tr>" }) -join ''
    A (Card 'media' "$($inativos.Count) contas ativas sem uso há mais de $DiasInativo dias" "<p>Contas habilitadas sem uso ampliam a superfície de ataque, principalmente as de terceiros e as com privilégio administrativo. Candidatas a desabilitar após confirmação com o cliente.</p><div class=""scroll""><table class=""mini""><thead><tr><th>Conta</th><th>Login</th><th>Setor</th><th class=""num"">Último logon</th></tr></thead><tbody>$tr</tbody></table></div>")
    $nCards++
}
if ($nunca.Count) {
    $pct = [Math]::Round(100 * $nunca.Count / [Math]::Max(1, $nUsers))
    A (Card 'baixa' "$($nunca.Count) contas com senha que nunca expira" "<p>Isso representa $pct% das contas ativas. Para contas pessoais, o ideal é seguir a política de expiração do domínio; a exceção se justifica apenas para contas de serviço.</p>")
    $nCards++
}
if ($dups.Count) {
    $li = ($dups | ForEach-Object { "<li>$(NM $_[0]) <span class=""muted"">($(HE $_[0]))</span> e $(NM $_[1]) <span class=""muted"">($(HE $_[1]))</span></li>" }) -join ''
    A (Card 'baixa' 'Possíveis contas duplicadas' "<p>Contas diferentes com o mesmo nome de exibição. Pode ser uma conta antiga que ficou ativa:</p><ul>$li</ul>")
    $nCards++
}
if ($abertas.Count) {
    $li = ($abertas | ForEach-Object { $mx = ($Calc[$_].EF.Values | Measure-Object -Maximum).Maximum; if (-not $mx) { $mx = 0 }; "<li><strong>$(HE $_)</strong> — todos os usuários, com $($TXT[[int]$mx].ToLower()) efetivo</li>" }) -join ''
    A (Card 'baixa' 'Pastas abertas para todos os usuários' "<p>Confirmar que o conteúdo dessas pastas pode ser visto por qualquer pessoa da empresa:</p><ul>$li</ul>")
    $nCards++
}
if (-not $nCards) { A '<p class="lead">Nenhum ponto de atenção identificado pelas regras automáticas.</p>' }
A '</div></section>'

# ---- 3. matriz ----
A @"
<section id="matriz"><h2><span>3</span>Matriz de acessos efetivos</h2>
<p class="lead">Cada linha é um usuário ativo e cada coluna uma pasta compartilhada. O nível mostrado já considera a combinação entre compartilhamento e NTFS na raiz da pasta.</p>
<div class="legend"><span class="lv l3">Controle total</span><span class="lv l2">Modificar</span><span class="lv l1">Leitura</span><span class="lv l0">—</span><span class="muted">Sem acesso</span>
<label class="noprint"><input type="checkbox" id="hideAdm"> Ocultar administradores</label>
<input class="noprint" id="fMat" placeholder="Filtrar usuário…"></div>
<div class="matwrap"><table class="mat" id="mat"><thead><tr><th class="uh">Usuário</th>
"@
foreach ($a in $alvos) { A "<th class=""rot""><div><span>$(HE $a.Nome)</span></div></th>" }
A '<th class="num">Pastas</th></tr></thead><tbody>'
foreach ($st in $setores.Keys) {
    $ls = $setores[$st]
    A "<tr class=""grp-row""><td colspan=""$($nSh + 2)"">$(HE $st) <span class=""muted"">· $($ls.Count)</span></td></tr>"
    foreach ($l in $ls) {
        $isAdm = $admSet.ContainsKey($l)
        $tags = $(if ($isAdm) { ' <span class="tag t-adm">adm</span>' }) + $(if ($U[$l].Terceiro) { ' <span class="tag t-terc">terceiro</span>' })
        A "<tr data-u=""$(HE ($U[$l].Nome + ' ' + $l).ToLower())"" class=""$(if ($isAdm) { 'adm' })""><td class=""uh""><span class=""un"">$(NM $l)</span><span class=""ul"">$(HE $l)</span>$tags</td>"
        $cnt = 0
        foreach ($a in $alvos) {
            $d = $Calc[$a.Nome]
            if ($d.EF.ContainsKey($l)) {
                $lv = $d.EF[$l]; $cnt++
                $o = if ($d.Orig.ContainsKey($l)) { @($d.Orig[$l]) -join ', ' } else { '' }
                $spl = if ($d.SP.ContainsKey($l)) { $d.SP[$l] } else { 0 }
                $tip = "$($a.Nome): $($TXT[$lv]) · NTFS via $o · compartilhamento: $($TXT[$spl].ToLower())"
                A "<td class=""c l$lv"" title=""$(HE $tip)"">$('TML'[3 - $lv])</td>"
            } else { A '<td class="c l0"></td>' }
        }
        A "<td class=""num"">$cnt</td></tr>"
    }
}
A '</tbody></table></div><p class="note">T = controle total · M = modificar (ler, gravar e excluir) · L = somente leitura. Passe o mouse sobre uma célula para ver de onde vem a permissão.</p></section>'

# ---- 4. detalhe por pasta ----
A '<section id="pastas"><h2><span>4</span>Detalhe por pasta compartilhada</h2><p class="lead">Para cada pasta: quem está liberado no compartilhamento, quem está liberado no NTFS (com os membros de cada grupo) e o total de usuários com acesso efetivo.</p><div class="folders">'
foreach ($a in $alvos) {
    $d = $Calc[$a.Nome]
    $vals = @($d.EF.Values)
    $c3 = @($vals | Where-Object { $_ -eq 3 }).Count; $c2 = @($vals | Where-Object { $_ -eq 2 }).Count; $c1 = @($vals | Where-Object { $_ -eq 1 }).Count
    if ($a.Manual) { $shl = '<tr><td colspan="2" class="muted">Caminho informado manualmente (sem compartilhamento)</td></tr>' }
    else {
        $shl = ($a.Acl | Sort-Object @{e={ -$SLV[$_.Perm] }} | ForEach-Object {
            $neg = if ($_.Tipo -eq 'Deny') { ' <span class="tag t-adm">NEGAR</span>' } else { '' }
            "<tr><td>$(HE (Curto $_.Conta))$neg</td><td>$($SPT[$_.Perm])</td></tr>" }) -join ''
    }
    $nl = (@($d.NP.Values) | Sort-Object @{e={ $_.Permitir }}, @{e={ -$_.Nivel }}, @{e={ $_.Sid -ne $SidAdmins }}, @{e={ (Resolve-Principal $_.Sid).Curto }} | ForEach-Object {
        $neg = if (-not $_.Permitir) { ' <span class="tag t-adm">NEGAR</span>' } else { '' }
        "<tr><td>$(PrincipalHtml $_)$neg</td><td class=""nw"">$(HE $_.Rotulo)</td></tr>" }) -join ''
    $flag = if ($todosFull -contains $a.Nome) { '<span class="tag t-warn">Todos: controle total no compartilhamento</span>' } else { '' }
    $exc = ''
    if ($d.Exc.Count) {
        $exc = '<div class="exc"><b>Subpastas com permissão própria:</b><ul>' + (($d.Exc | ForEach-Object {
            $p = Resolve-Principal $_.Sid
            "<li><code>$(HE $_.Pasta)</code> — $(HE $p.Curto): $(HE $_.Rotulo)$(if (-not $_.Permitir) { ' (NEGAR)' })</li>" }) -join '') + '</ul></div>'
    }
    A @"
<article class="folder"><header><h3>$(HE $a.Nome)</h3><code>$(HE $a.Path)</code>$flag</header>
<div class="fstats"><div><b>$($d.EF.Count)</b> com acesso</div><div><span class="dot l3"></span>$c3 total</div><div><span class="dot l2"></span>$c2 modificar</div><div><span class="dot l1"></span>$c1 leitura</div></div>
<div class="fgrid"><div><h4>Compartilhamento</h4><table class="mini"><tbody>$shl</tbody></table></div>
<div><h4>Permissões NTFS da pasta</h4><table class="mini ntfs"><tbody>$nl</tbody></table></div></div>$exc</article>
"@
}
A '</div></section>'

# ---- 5. usuários ----
A @"
<section id="usuarios"><h2><span>5</span>Usuários ativos no AD</h2>
<p class="lead">$nUsers contas habilitadas, agrupadas pela unidade organizacional (setor) em que estão no AD.</p>
<input class="noprint search" id="fUsr" placeholder="Filtrar por nome, login, setor ou grupo…">
<table class="users" id="usr"><thead><tr><th>Nome</th><th>Login</th><th>Setor (OU)</th><th>Grupos</th><th class="num">Último logon</th><th class="num">Pastas</th><th>Observações</th></tr></thead><tbody>
"@
foreach ($l in $ordUsers) {
    $usr = $U[$l]
    $obs = @()
    if ($admSet.ContainsKey($l)) { $obs += '<span class="tag t-adm">administrador</span>' }
    if ($usr.Terceiro)             { $obs += '<span class="tag t-terc">terceiro</span>' }
    if ($inaSet.ContainsKey($l)) { $obs += "<span class=""tag t-warn"">sem uso &gt;$($DiasInativo)d</span>" }
    if ($usr.Nunca)                { $obs += '<span class="tag">senha não expira</span>' }
    $np = @($alvos | Where-Object { $Calc[$_.Nome].EF.ContainsKey($l) }).Count
    $grp = if ($usr.Grupos.Count) { ($usr.Grupos | ForEach-Object { HE $_ }) -join ', ' } else { '<span class="muted">—</span>' }
    A "<tr><td>$(HE $usr.Nome)</td><td class=""mono"">$(HE $l)</td><td>$(HE $usr.Setor)</td><td class=""grps"">$grp</td><td class=""num"">$(DT $usr.Logon)</td><td class=""num"">$np</td><td>$($obs -join ' ')</td></tr>"
}
A '</tbody></table></section>'

# ---- metodologia ----
A @"
<section id="recom"><div class="method"><h4>Como o levantamento foi feito</h4><p>Script PowerShell executado no servidor <code>$(HE $env:COMPUTERNAME)</code>$(if ($VersaoSO) { " (Windows Server $(HE $VersaoSO))" }) em $dataTxt, lendo o AD (contas habilitadas e grupos, com expansão recursiva) e as ACLs de $($pastasComAcl.Count) pastas com permissões próprias, até $Profundidade nível(is) abaixo de cada compartilhamento. Contas de sistema (SYSTEM, CREATOR OWNER) foram desconsideradas. O “último logon” vem do atributo replicado do AD e pode ter alguns dias de atraso.$(if ($erros.Count) { " $($erros.Count) pasta(s) não puderam ser lidas — ver arquivo 99_Erros_Acesso." })</p></div>
</section></main>
<footer class="wrap foot">$(HE $Responsavel) · Relatório gerado em $dataTxt · Documento de uso interno do cliente</footer>
"@
A @'
<script>
function f(id,sel,attr){const i=document.getElementById(id);if(!i)return;i.addEventListener('input',()=>{const q=i.value.toLowerCase();document.querySelectorAll(sel).forEach(r=>{const t=attr?r.dataset.u:r.textContent.toLowerCase();r.style.display=(!q||t.includes(q))?'':'none'})})}
f('fMat','#mat tbody tr[data-u]',true);f('fUsr','#usr tbody tr',false);
document.getElementById('hideAdm').addEventListener('change',e=>document.getElementById('mat').classList.toggle('noadm',e.target.checked));
</script></body></html>
'@

$arqHtml = Join-Path $Saida ("Relatorio_Acessos_" + ($Cliente -replace '[^\w\-]', '_') + "_$stamp.html")
[System.IO.File]::WriteAllText($arqHtml, $sb.ToString(), (New-Object System.Text.UTF8Encoding($true)))

Write-Host ''
Write-Host "Concluído. Usuários ativos: $nUsers | Pastas: $nSh | Administradores: $($ADM.Count) | Sem uso: $($inativos.Count)" -ForegroundColor Green
if ($erros.Count) { Write-Host "Atenção: $($erros.Count) pasta(s) sem acesso - ver 99_Erros_Acesso_$stamp.csv" -ForegroundColor Yellow }
Write-Host "Relatório: $arqHtml" -ForegroundColor Green
if ($Abrir) { Start-Process $arqHtml }
