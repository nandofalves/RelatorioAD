#Requires -Version 2.0
<#
  Itens adicionais (opcionais) do RelatorioAD: comparação com o levantamento anterior,
  contas privilegiadas, política de senha, GPOs e computadores do domínio.

  O build.ps1 embute este arquivo na região MODULOS das 3 bases. Aqui só ficam funções e
  tabelas constantes (nenhum efeito colateral), para o tools\Test-Modulos.ps1 fazer
  dot-source sem AD. As funções leem as variáveis do script ($dom, $U, $ADM, $Calc...)
  e usam o HE, o DT e o Card da base.

  Regras PS 2.0: foreach só sobre @(...); $x = @(if ...); .Count só sobre @(...);
  sem [pscustomobject], [ordered], HashSet, -in/-notin nem Where-Object sem bloco.
#>

# Módulos na ordem de execução, da navegação e dos cards.
#   Id = nome do switch · Chave = campo em $MOD · Cards/Secao = funções de HTML (fases 2 a 6)
$ModulosDef = @(
    @{ Id = 'Comparar';      Chave = 'Mudancas'; Ancora = 'mod-mudancas';     Titulo = 'Mudanças';             Metodo = 'as mudanças desde o levantamento anterior'; Cards = 'CardsMudancas';     Secao = 'SecaoMudancas' }
    @{ Id = 'ContasDominio'; Chave = 'Contas';   Ancora = 'mod-contas';       Titulo = 'Contas privilegiadas'; Metodo = 'as contas privilegiadas e de risco do domínio'; Cards = 'CardsContas';   Secao = 'SecaoContas' }
    @{ Id = 'PoliticaSenha'; Chave = 'Senha';    Ancora = 'mod-senha';        Titulo = 'Política de senha';    Metodo = 'a política de senha e de bloqueio';      Cards = 'CardsSenha';        Secao = 'SecaoSenha' }
    @{ Id = 'GPOs';          Chave = 'Gpo';      Ancora = 'mod-gpos';         Titulo = 'GPOs';                 Metodo = 'as GPOs (via LDAP e SYSVOL)';            Cards = 'CardsGpo';          Secao = 'SecaoGpo' }
    @{ Id = 'Computadores';  Chave = 'Comp';     Ancora = 'mod-computadores'; Titulo = 'Computadores';         Metodo = 'os computadores do domínio';             Cards = 'CardsComputadores'; Secao = 'SecaoComputadores' }
)

function Get-ModuloDef([string]$id) {
    foreach ($d in @($ModulosDef)) { if ($d.Id -eq $id) { return $d } }
    $null
}

# Escapa um valor para uso em filtro LDAP (RFC 4515)
function ConvertTo-LdapEscape([string]$s) {
    $s.Replace('\', '\5c').Replace('*', '\2a').Replace('(', '\28').Replace(')', '\29').Replace([string][char]0, '\00')
}

# Executa um módulo; em caso de erro, avisa no log e devolve Ok=$false (não derruba o relatório)
function Invoke-Modulo([string]$titulo, [scriptblock]$bloco) {
    Write-Host "      · $titulo..."
    try {
        & $bloco
    } catch {
        $msg = $_.Exception.Message
        Write-Host "[AVISO] Módulo ${titulo}: $msg" -ForegroundColor Yellow
        New-Object PSObject -Property @{ Ok = $false; Erro = $msg; Avisos = @() }
    }
}

# Monta $MOD a partir dos switches do script. Ordem: módulos 2 a 5 → snapshot → comparação.
function Invoke-Modulos {
    $sel = @{ Comparar = $Comparar; ContasDominio = $ContasDominio; PoliticaSenha = $PoliticaSenha; GPOs = $GPOs; Computadores = $Computadores }
    $marcados = @()
    foreach ($d in @($ModulosDef)) { if ($sel[$d.Id]) { $marcados += $d.Id } }
    $MOD = @{ Marcados = $marcados; Mudancas = $null; Contas = $null; Senha = $null; Gpo = $null; Comp = $null; Snapshot = $null }

    # 1. Módulos 2 a 5 (fases 3 a 6)
    if ($ContasDominio) {
        $MOD.Contas = Invoke-Modulo 'Contas privilegiadas e de risco' {
            $r = Get-AchadosContas (Read-ModContas $dom) $hoje $DiasInativo $PadraoServico $DiasSenhaServico $PadraoTerceiros
            Write-CsvContas $r "$Saida\05_Contas_Privilegiadas_$stamp.csv"
            $r
        }
    }
    if ($PoliticaSenha) {
        $MOD.Senha = Invoke-Modulo 'Política de senha e bloqueio' { Get-ResultadoSenha (Read-ModSenha) }
    }
    if ($GPOs) {
        $MOD.Gpo = Invoke-Modulo 'GPOs' {
            $b = Read-ModGpo $dom
            $r = Get-AchadosGpo $b.Gpos $b.Alvos $dom.DistinguishedName $b.SysvolOk
            $r.Avisos = @(@($b.Avisos) | Where-Object { $_ })
            Write-CsvGpo $r "$Saida\06_GPOs_$stamp.csv"
            Write-CsvGpoConfig $r "$Saida\06_GPOs_Configuracoes_$stamp.csv"
            $r
        }
    }
    if ($Computadores) {
        $MOD.Comp = Invoke-Modulo 'Computadores do domínio' {
            $r = Get-AchadosComputadores (Read-ModComputadores) $hoje $DiasInativo
            Write-CsvComputadores $r "$Saida\07_Computadores_$stamp.csv"
            $r
        }
    }

    # 2. Levantamento anterior: localizado e lido antes de gravar o atual
    $arqSnap = Join-Path $Saida "00_Snapshot_$stamp.csv"
    $antes = $null; $arqAntes = $null; $erroAntes = $null
    if ($Comparar) {
        try {
            $arqAntes = if ($CompararCom) { $CompararCom } else { Find-SnapshotAnterior $Saida $dom.DNSRoot $arqSnap }
            if ($arqAntes) { $antes = Read-Snapshot $arqAntes }
        } catch { $erroAntes = $_.Exception.Message }
    }

    # 3. Snapshot atual: gravado sempre
    $linhas = @(Get-LinhasSnapshot $MOD)
    $MOD.Snapshot = Write-Snapshot $linhas $arqSnap

    # 4. Comparação
    if ($Comparar) {
        $MOD.Mudancas = Invoke-Modulo 'Comparação com o levantamento anterior' {
            if ($erroAntes) { throw $erroAntes }
            $r = Read-ModComparar $antes (ConvertTo-Snapshot $linhas) $arqAntes
            if (@($r.Mudancas).Count) {
                $arq04 = "$Saida\04_Mudancas_$stamp.csv"
                try {
                    @($r.Mudancas) | Select-Object Categoria, Item, Mudanca, Antes, Depois | Export-Csv $arq04 @csvOpts
                    $r.Csv = $arq04
                } catch { $r.Avisos += "não foi possível gravar o arquivo 04_Mudancas: $($_.Exception.Message)" }
            }
            $r
        }
    }
    $MOD
}

# ---------------------------------------------------------------------
# Snapshot (00_Snapshot_<stamp>.csv) e comparação com o anterior (módulo 1)
# ---------------------------------------------------------------------
$SnapNivel = @{ '1' = 'Leitura'; '2' = 'Modificar'; '3' = 'Controle total' }
$SnapGpoStatus = @{ '0' = 'ativa'; '1' = 'usuário desativado'; '2' = 'computador desativado'; '3' = 'toda desativada' }
# Módulos que gravam linhas no snapshot: Id → nome usado em "não comparado: ..."
$SnapModulos = @(
    @{ Id = 'ContasDominio'; Nome = 'os grupos privilegiados' }
    @{ Id = 'GPOs';          Nome = 'as GPOs' }
    @{ Id = 'Computadores';  Nome = 'os computadores' }
)
$MudCategorias = @('Contas', 'Administradores do servidor', 'Pastas', 'Acesso', 'Grupos privilegiados', 'GPOs', 'Computadores')

function NovaLinhaSnap([string]$tipo, [string]$chave, [string]$valor) {
    New-Object PSObject -Property @{ Tipo = $tipo; Chave = $chave; Valor = $valor }
}

# Linhas Tipo/Chave/Valor do levantamento atual, a partir das variáveis da base
function Get-LinhasSnapshot($MOD) {
    $mods = @()
    foreach ($m in @($SnapModulos)) {
        $r = $MOD[(Get-ModuloDef $m.Id).Chave]
        if ($r -and $r.Ok) { $mods += $m.Id }
    }
    $escopo = if ($Caminhos) { @($Caminhos) -join '|' } else { 'compartilhamentos' }
    NovaLinhaSnap '_meta' 'Formato' '1'
    NovaLinhaSnap '_meta' 'Dominio' $dom.DNSRoot
    NovaLinhaSnap '_meta' 'Data' $hoje.ToString('yyyy-MM-dd HH:mm', [Globalization.CultureInfo]::InvariantCulture)
    NovaLinhaSnap '_meta' 'Servidor' $env:COMPUTERNAME
    NovaLinhaSnap '_meta' 'Escopo' $escopo
    NovaLinhaSnap '_meta' 'Modulos' ($mods -join ',')
    foreach ($l in @($U.Keys | Sort-Object)) { NovaLinhaSnap 'Usuario' $l $U[$l].Nome }
    foreach ($l in @($ADM)) { if ($l) { NovaLinhaSnap 'Admin' $l '' } }
    foreach ($a in @($alvos)) {
        if (-not $a) { continue }
        NovaLinhaSnap 'Pasta' $a.Nome $a.Path
        $c = $Calc[$a.Nome]
        if (-not $c -or -not $c.EF) { continue }
        foreach ($l in @($c.EF.Keys | Sort-Object)) { NovaLinhaSnap 'Acesso' "$($a.Nome)|$l" ([string]$c.EF[$l]) }
    }
    # Membros ativos dos grupos privilegiados
    $ct = $MOD['Contas']
    if ($ct -and $ct.Ok) {
        foreach ($x in @($ct.Privilegiadas)) {
            if (-not $x -or -not $x.Ativa) { continue }
            foreach ($g in @($x.Grupos)) { if ($g) { NovaLinhaSnap 'Privilegiado' "$g|$($x.Login)" '' } }
        }
    }
    # GPOs: nome|flags|versao
    $gp = $MOD['Gpo']
    if ($gp -and $gp.Ok) {
        foreach ($g in @($gp.Gpos)) { if ($g) { NovaLinhaSnap 'GPO' $g.Guid "$($g.Nome)|$($g.Flags)|$($g.Versao)" } }
    }
    # Computadores: ativo(1/0)|SO
    $cp = $MOD['Comp']
    if ($cp -and $cp.Ok) {
        foreach ($c in @($cp.Computadores)) {
            if ($c) { NovaLinhaSnap 'Computador' $c.Nome "$(if ($c.Ativo) { '1' } else { '0' })|$($c.SO)" }
        }
    }
}

# Grava o snapshot; se falhar, avisa no log e devolve $null (não derruba o relatório)
function Write-Snapshot($linhas, [string]$arquivo) {
    try {
        @($linhas) | Select-Object Tipo, Chave, Valor | Export-Csv $arquivo @csvOpts
        $arquivo
    } catch {
        Write-Host "[AVISO] Não foi possível gravar o levantamento ${arquivo}: $($_.Exception.Message)" -ForegroundColor Yellow
        $null
    }
}

# Linhas → @{ Meta = @{ chave = valor }; Dados = @{ Tipo = @{ Chave = Valor } } }
function ConvertTo-Snapshot($linhas) {
    $meta = @{}; $dados = @{}
    foreach ($x in @($linhas)) {
        if (-not $x) { continue }
        $t = [string]$x.Tipo
        if ($t -eq '_meta') { $meta[[string]$x.Chave] = [string]$x.Valor; continue }
        if (-not $dados.ContainsKey($t)) { $dados[$t] = @{} }
        $dados[$t][[string]$x.Chave] = [string]$x.Valor
    }
    @{ Meta = $meta; Dados = $dados }
}

# Lê um 00_Snapshot. -SoMeta lê só as primeiras linhas (as _meta), para a busca do anterior.
function Read-Snapshot([string]$arquivo, [switch]$SoMeta) {
    if ($SoMeta) { $linhas = @(Get-Content -LiteralPath $arquivo -TotalCount 8 | ConvertFrom-Csv -Delimiter ';') }
    else         { $linhas = @(Import-Csv -LiteralPath $arquivo -Delimiter ';') }
    $s = ConvertTo-Snapshot $linhas
    if ($s.Meta['Formato'] -ne '1' -or -not $s.Meta['Dominio']) { throw 'arquivo não é um levantamento do RelatorioAD' }
    $s
}

# Snapshot mais recente do mesmo domínio na pasta (ignora o atual e os ilegíveis)
function Find-SnapshotAnterior([string]$pasta, [string]$dominio, [string]$atual) {
    if (-not (Test-Path -LiteralPath $pasta)) { return $null }
    if ($atual) { $atual = [IO.Path]::GetFullPath($atual) }
    $melhor = $null; $data = ''
    foreach ($f in @(Get-ChildItem -LiteralPath $pasta -Filter '00_Snapshot_*.csv')) {
        if (-not $f -or $f.PSIsContainer) { continue }
        if ($atual -and $f.FullName -eq $atual) { continue }
        try { $m = (Read-Snapshot $f.FullName -SoMeta).Meta } catch { continue }
        if ($m['Dominio'] -ne $dominio) { continue }
        if ([string]$m['Data'] -gt $data) { $data = [string]$m['Data']; $melhor = $f.FullName }
    }
    $melhor
}

function Get-SnapData([string]$s) {
    try { [datetime]::ParseExact($s, 'yyyy-MM-dd HH:mm', [Globalization.CultureInfo]::InvariantCulture) } catch { $null }
}

# Lê o que falta no AD e compara. $antes = $null → sem levantamento anterior.
function Read-ModComparar($antes, $depois, [string]$arquivo) {
    $r = New-Object PSObject -Property @{ Ok = $true; Erro = $null; Avisos = @(); Arquivo = $arquivo; DataAnterior = $null
        OutroDominio = $null; Mudancas = @(); NaoComparado = @(); EscopoMudou = $false; Csv = $null }
    if (-not $antes) { return $r }
    $r.DataAnterior = Get-SnapData $antes.Meta['Data']
    if ($antes.Meta['Dominio'] -ne $depois.Meta['Dominio']) { $r.OutroDominio = $antes.Meta['Dominio']; return $r }
    # Contas que saíram: desativada ou excluída (uma consulta só)
    $desativados = @{}
    $ua = Get-SnapTabela $antes 'Usuario'; $ud = Get-SnapTabela $depois 'Usuario'
    $saiu = $false
    foreach ($l in @($ua.Keys)) { if (-not $ud.ContainsKey($l)) { $saiu = $true; break } }
    if ($saiu) { foreach ($u in @(Get-ADUser -Filter 'Enabled -eq $false')) { if ($u) { $desativados[$u.SamAccountName] = 1 } } }
    $c = Compare-Snapshot $antes $depois $desativados
    $r.Mudancas = @($c.Mudancas); $r.NaoComparado = @($c.NaoComparado); $r.EscopoMudou = $c.EscopoMudou
    $r
}

function Get-SnapTabela($s, [string]$tipo) {
    if ($s.Dados.ContainsKey($tipo)) { $s.Dados[$tipo] } else { @{} }
}

function Get-SnapModulos($s) {
    $h = @{}
    foreach ($m in @(([string]$s.Meta['Modulos']) -split ',')) { if ($m) { $h[$m.Trim()] = 1 } }
    $h
}

function NovaMudanca([string]$cat, [string]$item, [string]$mud, [string]$antes, [string]$depois) {
    New-Object PSObject -Property @{ Categoria = $cat; Item = $item; Mudanca = $mud; Antes = $antes; Depois = $depois; Pasta = $null }
}

# Une as chaves de duas tabelas, em ordem
function Get-ChavesUnidas($a, $b) {
    $h = @{}
    foreach ($k in @($a.Keys)) { $h[$k] = 1 }
    foreach ($k in @($b.Keys)) { $h[$k] = 1 }
    @($h.Keys | Sort-Object)
}

# Compara dois snapshots (pura). $desativados = @{ login = 1 } das contas desativadas hoje no AD.
function Compare-Snapshot($antes, $depois, $desativados) {
    $m = New-Object System.Collections.ArrayList
    $nao = @()
    $ua = Get-SnapTabela $antes 'Usuario'; $ud = Get-SnapTabela $depois 'Usuario'
    $nome = {
        param($l)
        $n = if ($ud.ContainsKey($l)) { $ud[$l] } elseif ($ua.ContainsKey($l)) { $ua[$l] } else { '' }
        if ($n -and $n -cne $l) { "$n ($l)" } else { $l }
    }

    # Contas
    $acD = Get-SnapTabela $depois 'Acesso'; $acA = Get-SnapTabela $antes 'Acesso'
    $nPastas = @{}
    foreach ($k in @($acD.Keys)) { $l = ($k -split '\|', 2)[1]; $nPastas[$l] = 1 + [int]$nPastas[$l] }
    foreach ($l in @(Get-ChavesUnidas $ua $ud)) {
        if (-not $ua.ContainsKey($l)) {
            $n = [int]$nPastas[$l]
            [void]$m.Add((NovaMudanca 'Contas' (& $nome $l) 'nova' '' "acesso a $n $(if ($n -eq 1) { 'pasta' } else { 'pastas' })"))
        } elseif (-not $ud.ContainsKey($l)) {
            $tipo = if ($desativados -and $desativados.ContainsKey($l)) { 'desativada' } else { 'excluída' }
            [void]$m.Add((NovaMudanca 'Contas' (& $nome $l) $tipo 'ativa' $tipo))
        }
    }

    # Administradores do servidor
    $aa = Get-SnapTabela $antes 'Admin'; $ad = Get-SnapTabela $depois 'Admin'
    foreach ($l in @(Get-ChavesUnidas $aa $ad)) {
        if (-not $aa.ContainsKey($l))     { [void]$m.Add((NovaMudanca 'Administradores do servidor' (& $nome $l) 'entrou' '' '')) }
        elseif (-not $ad.ContainsKey($l)) { [void]$m.Add((NovaMudanca 'Administradores do servidor' (& $nome $l) 'saiu' '' '')) }
    }

    # Pastas (só se o escopo for o mesmo)
    $pa = Get-SnapTabela $antes 'Pasta'; $pd = Get-SnapTabela $depois 'Pasta'
    $escopoMudou = ([string]$antes.Meta['Escopo'] -ne [string]$depois.Meta['Escopo'])
    if (-not $escopoMudou) {
        foreach ($p in @(Get-ChavesUnidas $pa $pd)) {
            if (-not $pa.ContainsKey($p))     { [void]$m.Add((NovaMudanca 'Pastas' $p 'nova' '' $pd[$p])) }
            elseif (-not $pd.ContainsKey($p)) { [void]$m.Add((NovaMudanca 'Pastas' $p 'removida' $pa[$p] '')) }
        }
    }

    # Acesso: só pastas e logins presentes nos dois levantamentos
    foreach ($k in @(Get-ChavesUnidas $acA $acD)) {
        $partes = $k -split '\|', 2
        $p = $partes[0]; $l = $partes[1]
        if (-not ($pa.ContainsKey($p) -and $pd.ContainsKey($p) -and $ua.ContainsKey($l) -and $ud.ContainsKey($l))) { continue }
        $va = [string]$acA[$k]; $vd = [string]$acD[$k]
        if ($va -eq $vd) { continue }
        $mud = if (-not $va) { 'ganhou' } elseif (-not $vd) { 'perdeu' } else { 'mudou' }
        $ta = if ($va) { $SnapNivel[$va] } else { 'Sem acesso' }
        $td = if ($vd) { $SnapNivel[$vd] } else { 'Sem acesso' }
        $x = NovaMudanca 'Acesso' "$p · $(& $nome $l)" $mud $ta $td
        $x.Pasta = $p
        [void]$m.Add($x)
    }

    # Módulos opcionais: só se o módulo rodou nos dois levantamentos
    $ma = Get-SnapModulos $antes; $md = Get-SnapModulos $depois
    $ambos = @{}
    foreach ($s in @($SnapModulos)) {
        if ($ma.ContainsKey($s.Id) -and $md.ContainsKey($s.Id)) { $ambos[$s.Id] = 1 }
        elseif ($md.ContainsKey($s.Id)) { $nao += "o levantamento anterior não incluiu $($s.Nome)" }
        elseif ($ma.ContainsKey($s.Id)) { $nao += "este levantamento não incluiu $($s.Nome)" }
    }

    if ($ambos['ContasDominio']) {
        $ga = Get-SnapTabela $antes 'Privilegiado'; $gd = Get-SnapTabela $depois 'Privilegiado'
        foreach ($k in @(Get-ChavesUnidas $ga $gd)) {
            $partes = $k -split '\|', 2
            $item = "$($partes[0]) · $(& $nome $partes[1])"
            if (-not $ga.ContainsKey($k))     { [void]$m.Add((NovaMudanca 'Grupos privilegiados' $item 'entrou' '' '')) }
            elseif (-not $gd.ContainsKey($k)) { [void]$m.Add((NovaMudanca 'Grupos privilegiados' $item 'saiu' '' '')) }
        }
    }

    if ($ambos['GPOs']) {
        $oa = Get-SnapTabela $antes 'GPO'; $od = Get-SnapTabela $depois 'GPO'
        foreach ($g in @(Get-ChavesUnidas $oa $od)) {
            $xa = if ($oa.ContainsKey($g)) { Split-SnapGpo $oa[$g] } else { $null }
            $xd = if ($od.ContainsKey($g)) { Split-SnapGpo $od[$g] } else { $null }
            if (-not $xa)     { [void]$m.Add((NovaMudanca 'GPOs' $xd.Nome 'nova' '' $SnapGpoStatus[$xd.Flags])); continue }
            if (-not $xd)     { [void]$m.Add((NovaMudanca 'GPOs' $xa.Nome 'removida' $SnapGpoStatus[$xa.Flags] '')); continue }
            if ($xa.Versao -ne $xd.Versao) { [void]$m.Add((NovaMudanca 'GPOs' $xd.Nome 'alterada' "versão $($xa.Versao)" "versão $($xd.Versao)")) }
            if ($xa.Flags -ne $xd.Flags)   { [void]$m.Add((NovaMudanca 'GPOs' $xd.Nome 'status' $SnapGpoStatus[$xa.Flags] $SnapGpoStatus[$xd.Flags])) }
        }
    }

    if ($ambos['Computadores']) {
        $ca = Get-SnapTabela $antes 'Computador'; $cd = Get-SnapTabela $depois 'Computador'
        foreach ($n in @(Get-ChavesUnidas $ca $cd)) {
            $xa = if ($ca.ContainsKey($n)) { ([string]$ca[$n]) -split '\|', 2 } else { $null }
            $xd = if ($cd.ContainsKey($n)) { ([string]$cd[$n]) -split '\|', 2 } else { $null }
            if (-not $xa) { [void]$m.Add((NovaMudanca 'Computadores' $n 'novo' '' $xd[1])); continue }
            if (-not $xd) { [void]$m.Add((NovaMudanca 'Computadores' $n 'removido' $xa[1] '')); continue }
            if ($xa[0] -eq '1' -and $xd[0] -ne '1') { [void]$m.Add((NovaMudanca 'Computadores' $n 'desativado' 'ativo' 'desativado')) }
            elseif ($xa[0] -ne '1' -and $xd[0] -eq '1') { [void]$m.Add((NovaMudanca 'Computadores' $n 'reativado' 'desativado' 'ativo')) }
            if ($xa[1] -ne $xd[1]) { [void]$m.Add((NovaMudanca 'Computadores' $n 'SO mudou' $xa[1] $xd[1])) }
        }
    }

    New-Object PSObject -Property @{ Mudancas = @($m); NaoComparado = $nao; EscopoMudou = $escopoMudou }
}

# "nome|flags|versao" (o nome pode ter |: flags e versão são os 2 últimos campos)
function Split-SnapGpo([string]$v) {
    $i2 = $v.LastIndexOf('|')
    $i1 = if ($i2 -gt 0) { $v.LastIndexOf('|', $i2 - 1) } else { -1 }
    if ($i1 -lt 0) { return New-Object PSObject -Property @{ Nome = $v; Flags = ''; Versao = '' } }
    New-Object PSObject -Property @{ Nome = $v.Substring(0, $i1); Flags = $v.Substring($i1 + 1, $i2 - $i1 - 1); Versao = $v.Substring($i2 + 1) }
}

# Contagem por categoria: @{ Categoria = @{ Mudanca = n } }, só categorias com mudança
function Get-ResumoMudancas($r) {
    $h = @{}
    foreach ($x in @($r.Mudancas)) {
        if (-not $x) { continue }
        if (-not $h.ContainsKey($x.Categoria)) { $h[$x.Categoria] = @{} }
        $h[$x.Categoria][$x.Mudanca] = 1 + [int]$h[$x.Categoria][$x.Mudanca]
    }
    $h
}

function TextoResumo($cont) {
    (@($cont.Keys | Sort-Object) | ForEach-Object { "$($cont[$_]) $_" }) -join ', '
}

function CardsMudancas($r) {
    if (-not $r.DataAnterior -or $r.OutroDominio) { return }
    $lista = @($r.Mudancas)
    if (-not $lista.Count) { return }
    $motivos = @()
    foreach ($x in @($lista)) {
        if ($x.Categoria -eq 'Administradores do servidor' -and $x.Mudanca -eq 'entrou') { $motivos += 'entrou alguém em Administradores do servidor' }
        elseif ($x.Categoria -eq 'Grupos privilegiados' -and $x.Mudanca -eq 'entrou') { $motivos += 'entrou alguém em grupo privilegiado do domínio' }
        elseif ($x.Categoria -eq 'Acesso' -and $x.Mudanca -eq 'ganhou' -and $x.Pasta -match $PastasSensiveis) { $motivos += 'alguém ganhou acesso a pasta sensível' }
    }
    $motivos = @($motivos | Sort-Object -Unique)
    $sev = if ($motivos.Count) { 'media' } else { 'baixa' }
    $res = Get-ResumoMudancas $r
    $tr = ''
    foreach ($cat in @($MudCategorias)) { if ($res.ContainsKey($cat)) { $tr += "<tr><td>$(HE $cat)</td><td>$(HE (TextoResumo $res[$cat]))</td></tr>" } }
    $txt = if ($motivos.Count) { " Atenção: $($motivos -join '; ')." } else { '' }
    Card $sev "$($lista.Count) $(if ($lista.Count -eq 1) { 'mudança' } else { 'mudanças' }) desde $(DT $r.DataAnterior)" ("<p>Comparação com o levantamento anterior.$txt Detalhes na seção <a href=""#mod-mudancas"">Mudanças</a>.</p><table class=""mini""><thead><tr><th>Categoria</th><th>Mudanças</th></tr></thead><tbody>$tr</tbody></table>")
}

function SecaoMudancas($r, $num) {
    $abre = "<section id=""mod-mudancas""><h2><span>$num</span>"
    if ($r.OutroDominio) {
        return $abre + "Mudanças</h2><p class=""lead"">O arquivo escolhido é do domínio $(HE $r.OutroDominio); a comparação só é feita entre levantamentos do mesmo domínio.</p></section>"
    }
    if (-not $r.DataAnterior) {
        return $abre + 'Mudanças</h2><p class="lead">Nenhum levantamento anterior encontrado nesta pasta. Este levantamento foi salvo e servirá de base para a próxima comparação.</p></section>'
    }
    $h = $abre + "Mudanças desde $(DT $r.DataAnterior)</h2>"
    $h += "<p class=""lead"">Comparação com o levantamento de $($r.DataAnterior.ToString('dd/MM/yyyy HH:mm')) (<code>$(HE (Split-Path $r.Arquivo -Leaf))</code>).</p>"
    $notas = @()
    if ($r.EscopoMudou) { $notas += 'o escopo de pastas mudou entre os dois levantamentos, por isso pastas novas e removidas não foram listadas' }
    foreach ($n in @($r.NaoComparado)) { if ($n) { $notas += "não comparado: $n" } }
    foreach ($a in @($r.Avisos)) { if ($a) { $notas += $a } }
    foreach ($n in @($notas)) { $h += "<p class=""lead"">Observação: $(HE $n).</p>" }
    $lista = @($r.Mudancas)
    if (-not $lista.Count) { return $h + '<p class="lead">Nenhuma mudança encontrada.</p></section>' }
    $res = Get-ResumoMudancas $r
    $h += '<table class="mini"><thead><tr><th>Categoria</th><th>Mudanças</th></tr></thead><tbody>'
    foreach ($cat in @($MudCategorias)) { if ($res.ContainsKey($cat)) { $h += "<tr><td>$(HE $cat)</td><td>$(HE (TextoResumo $res[$cat]))</td></tr>" } }
    $h += '</tbody></table>'
    foreach ($cat in @($MudCategorias)) {
        if (-not $res.ContainsKey($cat)) { continue }
        $h += "<h3 class=""mod-sub"">$(HE $cat)</h3><div class=""scroll""><table class=""mini""><thead><tr><th>Item</th><th>Mudança</th><th>Antes</th><th>Depois</th></tr></thead><tbody>"
        foreach ($x in @($lista)) {
            if ($x.Categoria -ne $cat) { continue }
            $h += "<tr><td>$(HE $x.Item)</td><td>$(HE $x.Mudanca)</td><td>$(HE $x.Antes)</td><td>$(HE $x.Depois)</td></tr>"
        }
        $h += '</tbody></table></div>'
    }
    $h + '</section>'
}

# ---------------------------------------------------------------------
# Contas privilegiadas e de risco no domínio (módulo 2)
# ---------------------------------------------------------------------
# Grupos privilegiados por SID (sufixo do domínio ou SID completo do Builtin); Raiz = lido no domínio raiz da floresta
$ContasGrupos = @(
    @{ Sid = '-512';         Nome = 'Domain Admins' }
    @{ Sid = 'S-1-5-32-544'; Nome = 'Administrators' }
    @{ Sid = 'S-1-5-32-548'; Nome = 'Account Operators' }
    @{ Sid = 'S-1-5-32-549'; Nome = 'Server Operators' }
    @{ Sid = 'S-1-5-32-550'; Nome = 'Print Operators' }
    @{ Sid = 'S-1-5-32-551'; Nome = 'Backup Operators' }
    @{ Sid = '-519';         Nome = 'Enterprise Admins'; Raiz = $true }
    @{ Sid = '-518';         Nome = 'Schema Admins';     Raiz = $true }
)
# Achados na ordem dos cards. Titulo = { param($n, $r) }; Valor = { param($c) } preenche a 3ª coluna da tabela do card.
$ContasAchadosDef = @(
    @{ Id = 'SenhaNaoExigida'; Nome = 'Senha não exigida'; Sev = 'alta'; Coluna = 'Grupos privilegiados'
       Titulo = { param($n, $r) "$n $(Plural $n 'conta ativa' 'contas ativas') sem exigência de senha" }
       Texto = 'Com o atributo "senha não exigida" ligado, a conta pode ficar com senha em branco. Desligar o atributo e confirmar que a conta tem senha.'
       Valor = { param($c) TextoGrupos $c } }
    @{ Id = 'DelegacaoIrrestrita'; Nome = 'Delegação irrestrita'; Sev = 'alta'; Coluna = 'Tipo'
       Titulo = { param($n, $r) "$n $(Plural $n 'conta' 'contas') com delegação irrestrita" }
       Texto = 'Na delegação irrestrita, quem controla a conta pode se passar por qualquer usuário que se autentique nela, inclusive administradores. Trocar por delegação restrita ou remover, se não for usada. Os controladores de domínio não entram nesta lista.'
       Valor = { param($c) $c.Tipo } }
    @{ Id = 'SemPreAuth'; Nome = 'Pré-autenticação desativada'; Sev = 'alta'; Coluna = 'Senha alterada'
       Titulo = { param($n, $r) "$n $(Plural $n 'conta' 'contas') sem pré-autenticação Kerberos" }
       Texto = 'Sem a pré-autenticação Kerberos, qualquer pessoa na rede pode pedir dados da conta e tentar descobrir a senha offline (AS-REP roasting). Religar a pré-autenticação.'
       Valor = { param($c) DT $c.SenhaEm } }
    @{ Id = 'Reversivel'; Nome = 'Criptografia reversível'; Sev = 'alta'; Coluna = 'Senha alterada'
       Titulo = { param($n, $r) "$n $(Plural $n 'conta' 'contas') com senha em criptografia reversível" }
       Texto = 'Com a criptografia reversível, a senha fica armazenada de forma recuperável. Desligar a opção e trocar a senha.'
       Valor = { param($c) DT $c.SenhaEm } }
    @{ Id = 'PrivTerceiro'; Nome = 'Terceiro com privilégio no domínio'; Sev = 'alta'; Coluna = 'Grupos privilegiados'
       Titulo = { param($n, $r) "$n $(Plural $n 'conta' 'contas') de terceiros com privilégio no domínio" }
       Texto = 'Contas em OUs de terceiros/fornecedores com privilégio administrativo no domínio inteiro. Confirmar a necessidade com o cliente; preferir acesso temporário ou com privilégio menor.'
       Valor = { param($c) TextoGrupos $c } }
    @{ Id = 'PrivInativo'; Nome = 'Privilegiada sem uso'; Sev = 'media'; Coluna = 'Último logon'
       Titulo = { param($n, $r) "$n $(Plural $n 'conta privilegiada' 'contas privilegiadas') sem uso há mais de $($r.DiasInativo) dias" }
       Texto = 'Contas administrativas paradas são alvo fácil: ninguém percebe se forem usadas indevidamente. Candidatas a desabilitar após confirmação com o cliente.'
       Valor = { param($c) DT $c.UltimoLogon } }
    @{ Id = 'Servico'; Nome = 'Conta de serviço com senha antiga'; Sev = 'media'; Coluna = 'Senha alterada'
       Titulo = { param($n, $r) "$n $(Plural $n 'conta' 'contas') de serviço com senha de mais de $($r.DiasSenhaServico) dias" }
       Texto = 'Contas de serviço (identificadas por SPN ou pelo padrão de nome) costumam ter privilégio alto e senha que nunca muda; com SPN, a senha pode ser atacada offline (Kerberoasting). Trocar a senha ou migrar para conta de serviço gerenciada (gMSA).'
       Valor = { param($c) DT $c.SenhaEm } }
    @{ Id = 'Admin500'; Nome = 'Administrador interno com nome padrão'; Sev = 'baixa'; Coluna = 'Senha alterada'
       Titulo = { param($n, $r) 'Administrador interno do domínio ativo e com o nome padrão' }
       Texto = 'O Administrador interno (RID 500) está ativo e com o nome padrão, conhecido por qualquer atacante. Usar contas administrativas nominais e renomear ou desabilitar esta, mantendo a senha guardada em local seguro.'
       Valor = { param($c) DT $c.SenhaEm } }
    @{ Id = 'NuncaLogou'; Nome = 'Nunca fez logon'; Sev = 'baixa'; Coluna = 'Criada em'
       Titulo = { param($n, $r) "$n $(Plural $n 'conta ativa que nunca fez logon' 'contas ativas que nunca fizeram logon')" }
       Texto = 'Contas habilitadas, criadas antes do limite de dias sem uso e que nunca foram usadas. Candidatas a desabilitar após confirmação com o cliente.'
       Valor = { param($c) DT $c.CriadaEm } }
)

function Plural($n, [string]$um, [string]$varios) { if ($n -eq 1) { $um } else { $varios } }
function TextoGrupos($c) { if (@($c.Grupos).Count) { @($c.Grupos) -join ', ' } else { '—' } }
function Get-OUConta([string]$dn) { ($dn -split '(?<!\\),', 2)[1] }

# lastLogonTimestamp / pwdLastSet → data (0, vazio ou inválido = $null)
function ConvertFrom-FileTimeAD($v) {
    if (-not $v) { return $null }
    try { $n = [int64]$v } catch { return $null }
    if ($n -le 0 -or $n -ge [datetime]::MaxValue.ToFileTimeUtc()) { return $null }
    [datetime]::FromFileTime($n)
}

# Formato comum das contas: Dn, Login, Nome, Tipo, Ativa, UltimoLogon, SenhaEm, CriadaEm, Spn
function NovaContaAD($dn, $login, $nome, $tipo, $ativa, $logon, $senha, $criada, $spn) {
    New-Object PSObject -Property @{ Dn = [string]$dn; Login = [string]$login; Nome = $(if ($nome) { [string]$nome } else { [string]$login })
        Tipo = $tipo; Ativa = [bool]$ativa; UltimoLogon = $logon; SenhaEm = $senha; CriadaEm = $criada; Spn = @(@($spn) | Where-Object { $_ }) }
}

# Membro devolvido pelo Get-ADObject (atributos LDAP)
function ConvertFrom-MembroAD($o) {
    $cls = [string]$o.ObjectClass
    $tipo = if ($cls -eq 'computer') { 'computador' } elseif ($cls -like '*ManagedServiceAccount') { 'conta gerenciada' } else { 'usuario' }
    NovaContaAD $o.DistinguishedName $o.sAMAccountName $o.displayName $tipo (-not ([int64]$o.userAccountControl -band 2)) `
        (ConvertFrom-FileTimeAD $o.lastLogonTimestamp) (ConvertFrom-FileTimeAD $o.pwdLastSet) $o.whenCreated $o.servicePrincipalName
}

# Lê o AD. Grupos = { Nome; Sid; Membros (formato comum) }; Usuarios/Computadores = objetos do módulo ActiveDirectory.
function Read-ModContas($dom) {
    $avisos = @()
    $sidDom = $dom.DomainSID.Value
    $raiz = $dom.DNSRoot; $srvRaiz = $null; $sidRaiz = $null
    try {
        $raiz = (Get-ADForest).RootDomain
        if ($raiz -ne $dom.DNSRoot) { $sidRaiz = (Get-ADDomain -Server $raiz).DomainSID.Value; $srvRaiz = $raiz }
        else { $sidRaiz = $sidDom }
    } catch { $avisos += "Enterprise/Schema Admins não puderam ser lidos no domínio raiz ${raiz}: $($_.Exception.Message)" }

    $props = @('sAMAccountName', 'displayName', 'userAccountControl', 'lastLogonTimestamp', 'pwdLastSet', 'objectClass',
        'distinguishedName', 'whenCreated', 'servicePrincipalName')
    $grupos = @()
    foreach ($d in @($ContasGrupos)) {
        $srv = @{}
        if ($d.Raiz) {
            if (-not $sidRaiz) { continue }
            $sid = "$sidRaiz$($d.Sid)"
            if ($srvRaiz) { $srv.Server = $srvRaiz }
        } elseif ($d.Sid -like 'S-*') { $sid = $d.Sid }
        else { $sid = "$sidDom$($d.Sid)" }
        try {
            $g = Get-ADGroup -Identity $sid @srv
            $filtro = "(&(objectClass=user)(memberOf:1.2.840.113556.1.4.1941:=$(ConvertTo-LdapEscape $g.DistinguishedName)))"
            $membros = @()
            foreach ($o in @(Get-ADObject -LDAPFilter $filtro -Properties $props @srv)) { if ($o) { $membros += ConvertFrom-MembroAD $o } }
            $grupos += New-Object PSObject -Property @{ Nome = [string]$g.Name; Sid = $sid; Membros = $membros }
        } catch {
            if ($d.Raiz) { $avisos += "$($d.Nome) não pôde ser lido no domínio raiz ${raiz}: $($_.Exception.Message)" }
            else { $avisos += "o grupo $($d.Nome) ($sid) não pôde ser lido: $($_.Exception.Message)" }
        }
    }

    $usuarios = @(Get-ADUser -Filter 'Enabled -eq $true' -Properties PasswordNotRequired, TrustedForDelegation, DoesNotRequirePreAuth,
        AllowReversiblePasswordEncryption, ServicePrincipalNames, PasswordLastSet, LastLogonDate, whenCreated, DisplayName)
    $comps = @()
    try { $comps = @(Get-ADComputer -Filter 'TrustedForDelegation -eq $true' -Properties PrimaryGroupID, LastLogonDate, OperatingSystem, PasswordLastSet, whenCreated) }
    catch { $avisos += "os computadores com delegação irrestrita não puderam ser lidos: $($_.Exception.Message)" }
    $adm = $null
    try { $adm = Get-ADUser -Identity "$sidDom-500" -Properties PasswordLastSet, LastLogonDate, whenCreated }
    catch { $avisos += "o Administrador interno não pôde ser lido: $($_.Exception.Message)" }

    New-Object PSObject -Property @{ Grupos = $grupos; Usuarios = $usuarios; Computadores = $comps; Admin500 = $adm; Avisos = $avisos }
}

function ConvertFrom-UsuarioAD($u) {
    NovaContaAD $u.DistinguishedName $u.SamAccountName $u.DisplayName 'usuario' ($u.Enabled -ne $false) $u.LastLogonDate $u.PasswordLastSet $u.whenCreated $u.ServicePrincipalNames
}

# Linha da conta em $porDn (criada na primeira vez)
function Get-ContaLinha($porDn, $o) {
    if (-not $porDn.ContainsKey($o.Dn)) {
        $ou = Get-OUConta $o.Dn
        $porDn[$o.Dn] = New-Object PSObject -Property @{ Dn = $o.Dn; Login = $o.Login; Nome = $o.Nome; Tipo = $o.Tipo; Ativa = $o.Ativa
            UltimoLogon = $o.UltimoLogon; SenhaEm = $o.SenhaEm; CriadaEm = $o.CriadaEm; Spn = @($o.Spn); OU = $ou
            Grupos = @(); Achados = @(); Servico = $false; Terceiro = $false; SemUso = $false; Avaliada = $false }
    }
    $porDn[$o.Dn]
}

function Add-Achado($c, [string]$id) { if ($c.Achados -notcontains $id) { $c.Achados += $id } }

# Conta de serviço (SPN ou padrão no login/OU) e senha antiga
function Set-Servico($c, $PadraoServico, $limSenha) {
    $c.Avaliada = $true
    $c.Servico = (@($c.Spn).Count -gt 0) -or ($PadraoServico -and ($c.Login -match $PadraoServico -or $c.OU -match $PadraoServico))
    if ($c.Servico -and $c.Ativa -and (-not $c.SenhaEm -or $c.SenhaEm -lt $limSenha)) { Add-Achado $c 'Servico' }
}

# Análise (pura). Devolve o resultado do módulo: Grupos {Nome; Sid; Ativos; Total}, Contas, Privilegiadas, Achados @{ Id = contas }.
function Get-AchadosContas($bruto, $hoje, $DiasInativo, $PadraoServico, $DiasSenhaServico, $PadraoTerceiros) {
    $limite = $hoje.AddDays(-$DiasInativo)
    $limSenha = $hoje.AddDays(-$DiasSenhaServico)
    $porDn = @{}

    # Membros dos grupos privilegiados
    $grupos = @()
    foreach ($g in @($bruto.Grupos)) {
        if (-not $g) { continue }
        $ativos = 0; $total = 0
        foreach ($m in @($g.Membros)) {
            if (-not $m) { continue }
            $c = Get-ContaLinha $porDn $m
            if ($c.Grupos -notcontains $g.Nome) { $c.Grupos += $g.Nome }
            $total++
            if ($m.Ativa) { $ativos++ }
        }
        $grupos += New-Object PSObject -Property @{ Nome = $g.Nome; Sid = $g.Sid; Ativos = $ativos; Total = $total }
    }

    # Usuários ativos: atributos de risco, nunca logou e contas de serviço
    foreach ($u in @($bruto.Usuarios)) {
        if (-not $u) { continue }
        $o = ConvertFrom-UsuarioAD $u
        $ach = @()
        if ($u.PasswordNotRequired) { $ach += 'SenhaNaoExigida' }
        if ($u.TrustedForDelegation) { $ach += 'DelegacaoIrrestrita' }
        if ($u.DoesNotRequirePreAuth) { $ach += 'SemPreAuth' }
        if ($u.AllowReversiblePasswordEncryption) { $ach += 'Reversivel' }
        if (-not $o.UltimoLogon -and $o.CriadaEm -and $o.CriadaEm -lt $limite) { $ach += 'NuncaLogou' }
        $existe = $porDn.ContainsKey($o.Dn)
        $tmp = Get-ContaLinha $porDn $o
        foreach ($a in @($ach)) { Add-Achado $tmp $a }
        Set-Servico $tmp $PadraoServico $limSenha
        if (-not $existe -and -not @($tmp.Achados).Count) { $porDn.Remove($o.Dn) }
    }

    # Computadores com delegação irrestrita (fora DCs e RODCs)
    foreach ($pc in @($bruto.Computadores)) {
        if (-not $pc -or $pc.Enabled -eq $false) { continue }
        if (@(516, 521) -contains [int]$pc.PrimaryGroupID) { continue }
        $c = Get-ContaLinha $porDn (NovaContaAD $pc.DistinguishedName $pc.SamAccountName $pc.Name 'computador' $true $pc.LastLogonDate $pc.PasswordLastSet $pc.whenCreated $null)
        $c.Avaliada = $true
        Add-Achado $c 'DelegacaoIrrestrita'
    }

    # Administrador interno ativo com o nome padrão
    $admSev = 'baixa'
    $a5 = $bruto.Admin500
    if ($a5 -and $a5.Enabled -ne $false -and $a5.SamAccountName -match '^(Administrator|Administrador)$') {
        $c = Get-ContaLinha $porDn (ConvertFrom-UsuarioAD $a5)
        Add-Achado $c 'Admin500'
        if (-not $c.SenhaEm -or $c.SenhaEm -lt $limSenha) { $admSev = 'media' }
    }

    # Privilegiadas: sem uso, terceiros e serviço (as que não vieram na lista de usuários)
    foreach ($c in @($porDn.Values)) {
        if (-not $c.Avaliada) { Set-Servico $c $PadraoServico $limSenha }
        if (-not @($c.Grupos).Count -or -not $c.Ativa) { continue }
        $ref = if ($c.UltimoLogon) { $c.UltimoLogon } else { $c.CriadaEm }
        if (-not $ref -or $ref -lt $limite) { $c.SemUso = $true; Add-Achado $c 'PrivInativo' }
        if ($PadraoTerceiros -and $c.OU -match $PadraoTerceiros) { $c.Terceiro = $true; Add-Achado $c 'PrivTerceiro' }
    }

    $contas = @(@($porDn.Values) | Sort-Object Nome, Login)
    $achados = @{}
    foreach ($d in @($ContasAchadosDef)) {
        $lista = @($contas | Where-Object { $_.Achados -contains $d.Id })
        if ($lista.Count) { $achados[$d.Id] = $lista }
    }
    New-Object PSObject -Property @{ Ok = $true; Erro = $null; Avisos = @($bruto.Avisos | Where-Object { $_ }); Grupos = $grupos
        Contas = $contas; Privilegiadas = @($contas | Where-Object { @($_.Grupos).Count }); Achados = $achados; Admin500Sev = $admSev
        DiasInativo = $DiasInativo; DiasSenhaServico = $DiasSenhaServico; Csv = $null }
}

# CSV 05_Contas_Privilegiadas; se falhar, só avisa na seção
function Write-CsvContas($r, [string]$arquivo) {
    try {
        @($r.Contas) | Where-Object { $_ } | ForEach-Object {
            $c = $_
            New-Object PSObject -Property @{ Login = $c.Login; Nome = $c.Nome; Tipo = $c.Tipo; Ativa = $c.Ativa
                GruposPrivilegiados = (@($c.Grupos) -join ', '); ContaServico = $c.Servico; SPN = (@($c.Spn) -join ', ')
                Achados = ((@($ContasAchadosDef) | Where-Object { $c.Achados -contains $_.Id } | ForEach-Object { $_.Nome }) -join ', ')
                UltimoLogon = $c.UltimoLogon; SenhaAlteradaEm = $c.SenhaEm; CriadaEm = $c.CriadaEm; OU = $c.OU }
        } | Select-Object Login, Nome, Tipo, Ativa, GruposPrivilegiados, ContaServico, SPN, Achados, UltimoLogon, SenhaAlteradaEm, CriadaEm, OU |
            Export-Csv $arquivo @csvOpts
        $r.Csv = $arquivo
    } catch { $r.Avisos += "não foi possível gravar o arquivo 05_Contas_Privilegiadas: $($_.Exception.Message)" }
}

function CardsContas($r) {
    foreach ($d in @($ContasAchadosDef)) {
        if (-not $r.Achados.ContainsKey($d.Id)) { continue }
        $lista = @($r.Achados[$d.Id])
        $sev = if ($d.Id -eq 'Admin500') { $r.Admin500Sev } else { $d.Sev }
        $tr = ''
        foreach ($c in @($lista)) { $tr += "<tr><td>$(HE $c.Nome)</td><td>$(HE $c.Login)</td><td>$(HE (& $d.Valor $c))</td></tr>" }
        Card $sev (& $d.Titulo $lista.Count $r) ("<p>$($d.Texto)</p><div class=""scroll""><table class=""mini""><thead><tr><th>Conta</th><th>Login</th><th>$($d.Coluna)</th></tr></thead><tbody>$tr</tbody></table></div>")
    }
}

function SecaoContas($r, $num) {
    $h = "<section id=""mod-contas""><h2><span>$num</span>Contas privilegiadas</h2>"
    $h += '<p class="lead">Membros (diretos e por grupos aninhados) dos grupos administrativos do domínio. Membros pelo grupo primário e de outros domínios ou florestas (via confiança) não aparecem. As contas com atributos de risco estão em Pontos de atenção e no arquivo 05_Contas_Privilegiadas.</p>'
    foreach ($a in @($r.Avisos)) { if ($a) { $h += "<p class=""lead"">Observação: $(HE $a).</p>" } }
    $h += '<table class="mini"><thead><tr><th>Grupo</th><th class="num">Membros ativos</th></tr></thead><tbody>'
    foreach ($g in @($r.Grupos)) { if ($g) { $h += "<tr><td>$(HE $g.Nome)</td><td class=""num"">$($g.Ativos)</td></tr>" } }
    $h += '</tbody></table>'
    $priv = @($r.Privilegiadas)
    if (-not $priv.Count) { return $h + '<p class="lead">Nenhuma conta nos grupos privilegiados.</p></section>' }
    $h += '<h3 class="mod-sub">Contas nos grupos privilegiados</h3><div class="scroll"><table class="mini"><thead><tr><th>Conta</th><th>Login</th><th>Grupos</th><th>Ativa</th><th class="num">Último logon</th><th class="num">Senha alterada</th></tr></thead><tbody>'
    foreach ($c in @($priv)) {
        $tags = ''
        if ($c.Terceiro) { $tags += '<span class="tag t-terc">terceiro</span>' }
        if ($c.Servico) { $tags += '<span class="tag">serviço</span>' }
        if ($c.SemUso) { $tags += "<span class=""tag t-warn"">sem uso &gt;$($r.DiasInativo)d</span>" }
        if ($c.Tipo -ne 'usuario') { $tags += "<span class=""tag"">$(HE $c.Tipo)</span>" }
        $h += "<tr><td>$(HE $c.Nome)$tags</td><td>$(HE $c.Login)</td><td>$(HE (TextoGrupos $c))</td><td>$(if ($c.Ativa) { 'sim' } else { 'não' })</td><td class=""num"">$(DT $c.UltimoLogon)</td><td class=""num"">$(DT $c.SenhaEm)</td></tr>"
    }
    $h + '</tbody></table></div></section>'
}

# ---------------------------------------------------------------------
# Política de senha e bloqueio (módulo 3)
# ---------------------------------------------------------------------
$SenhaSevOrdem = @{ alta = 3; media = 2; baixa = 1 }
$SenhaBaseline = 'Referência: Microsoft Security Baseline do Windows Server 2025 (publicada em 31/01/2025). Nela o bloqueio caiu de 10 para 3 tentativas, com 15 minutos de duração e de contador.'

# Política padrão e políticas refinadas (FGPP). AplicaA = nomes curtos (CN=) dos grupos/usuários.
function Read-ModSenha {
    $avisos = @()
    $padrao = Get-ADDefaultDomainPasswordPolicy
    $refinadas = @()
    try {
        foreach ($p in @(Get-ADFineGrainedPasswordPolicy -Filter * -Properties AppliesTo, Precedence)) {
            if (-not $p) { continue }
            $aplica = @()
            foreach ($dn in @($p.AppliesTo)) { if ($dn) { $aplica += Get-NomeCN $dn } }
            $refinadas += New-Object PSObject -Property @{ Nome = [string]$p.Name; Precedencia = $p.Precedence; AplicaA = $aplica; Politica = $p }
        }
    } catch { $avisos += "as políticas refinadas (FGPP) não puderam ser lidas: $($_.Exception.Message)" }
    New-Object PSObject -Property @{ Padrao = $padrao; Refinadas = $refinadas; Avisos = $avisos }
}

# 'CN=Grupo TI,OU=...' → 'Grupo TI' (vírgula escapada com \ faz parte do nome)
function Get-NomeCN([string]$dn) {
    $p = ($dn -split '(?<!\\),', 2)[0]
    if ($p -match '^CN=(.*)$') { $matches[1].Replace('\,', ',') } else { $dn }
}

# TimeSpan → texto. Zero ou valor "infinito" (o AD grava 0x8000000000000000) = $null.
function Get-Duracao($v) {
    if ($null -eq $v) { return $null }
    $t = [TimeSpan]$v
    if ($t -eq [TimeSpan]::Zero -or [math]::Abs($t.TotalDays) -gt 36500) { return $null }
    $t
}
function TextoDuracao([TimeSpan]$t) {
    if ($t.TotalDays -ge 1 -and $t.TotalDays -eq [math]::Floor($t.TotalDays)) { $n = [int]$t.TotalDays; return "$n $(Plural $n 'dia' 'dias')" }
    if ($t.TotalHours -ge 1 -and $t.TotalHours -eq [math]::Floor($t.TotalHours)) { $n = [int]$t.TotalHours; return "$n $(Plural $n 'hora' 'horas')" }
    "$([math]::Round($t.TotalMinutes, 1).ToString([Globalization.CultureInfo]::GetCultureInfo('pt-BR'))) min"
}

function NovoItemSenha([string]$item, [string]$atual, [string]$ref, [string]$sev, [string]$expl) {
    New-Object PSObject -Property @{ Item = $item; Atual = $atual; Referencia = $ref; Sev = $sev; Explicacao = $expl }
}

# Avaliação (pura) de uma política (padrão ou FGPP): linhas {Item; Atual; Referencia; Sev; Explicacao}, Sev = ok|baixa|media|alta|info
function Get-AvaliacaoSenha($pol) {
    $n = [int]$pol.MinPasswordLength
    $sev = if ($n -lt 8) { 'alta' } elseif ($n -lt 14) { 'media' } else { 'ok' }
    NovoItemSenha 'Tamanho mínimo' "$n $(Plural $n 'caractere' 'caracteres')" '14 ou mais' $sev 'Senhas curtas são descobertas rápido por tentativa e erro.'

    $sev = if ($pol.ComplexityEnabled) { 'ok' } else { 'alta' }
    NovoItemSenha 'Complexidade' $(if ($pol.ComplexityEnabled) { 'ligada' } else { 'desligada' }) 'ligada' $sev 'Sem complexidade, senhas como "12345678" ou o nome da empresa são aceitas.'

    $n = [int]$pol.PasswordHistoryCount
    $sev = if ($n -lt 24) { 'baixa' } else { 'ok' }
    NovoItemSenha 'Histórico' "$n $(Plural $n 'senha lembrada' 'senhas lembradas')" '24 ou mais' $sev 'Com histórico curto, o usuário volta a usar as mesmas senhas de antes.'

    $sev = if ($pol.ReversibleEncryptionEnabled) { 'alta' } else { 'ok' }
    NovoItemSenha 'Criptografia reversível' $(if ($pol.ReversibleEncryptionEnabled) { 'ligada' } else { 'desligada' }) 'desligada' $sev 'Ligada, as senhas ficam armazenadas de forma recuperável.'

    $lim = [int]$pol.LockoutThreshold
    $sev = if ($lim -eq 0) { 'alta' } elseif ($lim -gt 10) { 'baixa' } else { 'ok' }
    $atual = if ($lim -eq 0) { '0 (sem bloqueio)' } else { "$lim $(Plural $lim 'tentativa' 'tentativas')" }
    NovoItemSenha 'Bloqueio por tentativas' $atual '1 a 10 tentativas' $sev 'Sem bloqueio, é possível testar senhas sem limite; a baseline do Windows Server 2025 usa 3.'

    if ($lim -gt 0) {
        $d = Get-Duracao $pol.LockoutDuration
        $sev = if ($d -and $d.TotalMinutes -lt 15) { 'baixa' } else { 'ok' }
        NovoItemSenha 'Duração do bloqueio' $(if ($d) { TextoDuracao $d } else { 'até desbloqueio manual' }) '15 min ou mais' $sev 'Bloqueio curto deixa o atacante voltar a tentar logo em seguida.'

        $d = Get-Duracao $pol.LockoutObservationWindow
        $sev = if ($d -and $d.TotalMinutes -lt 15) { 'baixa' } else { 'ok' }
        NovoItemSenha 'Zerar contador após' $(if ($d) { TextoDuracao $d } else { '—' }) '15 min ou mais' $sev 'Contador zerado rápido permite várias tentativas espaçadas sem bloquear a conta.'
    }

    $d = Get-Duracao $pol.MinPasswordAge
    $sev = if (-not $d -or $d.TotalDays -lt 1) { 'baixa' } else { 'ok' }
    NovoItemSenha 'Idade mínima' $(if ($d) { TextoDuracao $d } else { '0 (troca imediata)' }) '1 dia ou mais' $sev 'Sem idade mínima, o usuário troca a senha várias vezes seguidas e volta à antiga, driblando o histórico.'

    $d = Get-Duracao $pol.MaxPasswordAge
    NovoItemSenha 'Validade máxima' $(if ($d) { TextoDuracao $d } else { 'nunca expira' }) 'informativo' 'info' 'A Microsoft deixou de recomendar a troca periódica obrigatória; o que protege é senha longa e bloqueio.'
}

function Test-AbaixoSenha($i) { $SenhaSevOrdem.ContainsKey([string]$i.Sev) }

# Monta o resultado do módulo (pura): Padrao = itens; Refinadas = {Nome; Precedencia; AplicaA; Itens}; Sev = maior severidade abaixo da referência
function Get-ResultadoSenha($bruto) {
    $padrao = @(Get-AvaliacaoSenha $bruto.Padrao)
    $refs = @()
    foreach ($f in @(@($bruto.Refinadas) | Where-Object { $_ } | Sort-Object Precedencia, Nome)) {
        $refs += New-Object PSObject -Property @{ Nome = $f.Nome; Precedencia = $f.Precedencia; AplicaA = @($f.AplicaA); Itens = @(Get-AvaliacaoSenha $f.Politica) }
    }
    $abaixo = @()
    foreach ($i in @($padrao)) { if (Test-AbaixoSenha $i) { $abaixo += New-Object PSObject -Property @{ Politica = 'Padrão do domínio'; Linha = $i } } }
    foreach ($f in @($refs)) {
        foreach ($i in @($f.Itens)) { if (Test-AbaixoSenha $i) { $abaixo += New-Object PSObject -Property @{ Politica = "Refinada: $($f.Nome)"; Linha = $i } } }
    }
    $sev = $null
    foreach ($a in @($abaixo)) { if (-not $sev -or $SenhaSevOrdem[$a.Linha.Sev] -gt $SenhaSevOrdem[$sev]) { $sev = $a.Linha.Sev } }
    New-Object PSObject -Property @{ Ok = $true; Erro = $null; Avisos = @(@($bruto.Avisos) | Where-Object { $_ }); Padrao = $padrao; Refinadas = $refs; Abaixo = $abaixo; Sev = $sev }
}

function CardsSenha($r) {
    $lista = @($r.Abaixo)
    if (-not $lista.Count) { return }
    $tr = ''
    foreach ($a in @($lista)) {
        $i = $a.Linha
        $tr += "<tr><td>$(HE $a.Politica)</td><td>$(HE $i.Item)</td><td class=""sev-$($i.Sev)"">$(HE $i.Atual)</td><td>$(HE $i.Referencia)</td></tr>"
    }
    $tit = "$($lista.Count) $(Plural $lista.Count 'item da política de senha abaixo' 'itens da política de senha abaixo') da referência"
    Card $r.Sev $tit ("<p>A política de senha e de bloqueio está mais fraca que a referência da Microsoft. Detalhes e explicação de cada item na seção <a href=""#mod-senha"">Política de senha</a>.</p><div class=""scroll""><table class=""mini""><thead><tr><th>Política</th><th>Item</th><th>Atual</th><th>Referência</th></tr></thead><tbody>$tr</tbody></table></div>")
}

function TabelaSenha($itens) {
    $h = '<div class="scroll"><table class="mini"><thead><tr><th>Item</th><th>Atual</th><th>Referência</th><th>Situação</th><th>Por quê</th></tr></thead><tbody>'
    foreach ($i in @($itens)) {
        if (-not $i) { continue }
        $sit = if (Test-AbaixoSenha $i) { '<span class="tag t-bad">abaixo</span>' } elseif ($i.Sev -eq 'info') { '<span class="tag">informativo</span>' } else { '<span class="tag t-ok">ok</span>' }
        $h += "<tr><td>$(HE $i.Item)</td><td class=""sev-$($i.Sev)"">$(HE $i.Atual)</td><td>$(HE $i.Referencia)</td><td>$sit</td><td>$(HE $i.Explicacao)</td></tr>"
    }
    $h + '</tbody></table></div>'
}

function SecaoSenha($r, $num) {
    $h = "<section id=""mod-senha""><h2><span>$num</span>Política de senha</h2>"
    $h += '<p class="lead">Política padrão do domínio (vale para todas as contas sem política refinada) comparada com a referência da Microsoft.</p>'
    foreach ($a in @($r.Avisos)) { if ($a) { $h += "<p class=""lead"">Observação: $(HE $a).</p>" } }
    $h += TabelaSenha $r.Padrao
    $refs = @($r.Refinadas)
    if ($refs.Count) {
        $h += '<h3 class="mod-sub">Políticas refinadas (FGPP)</h3><p class="lead">Substituem a política padrão para os grupos e usuários indicados; com mais de uma, vale a de menor precedência.</p>'
        $h += '<table class="mini"><thead><tr><th>Política</th><th class="num">Precedência</th><th>Aplica-se a</th></tr></thead><tbody>'
        foreach ($f in @($refs)) {
            $ap = if (@($f.AplicaA).Count) { @($f.AplicaA) -join ', ' } else { 'ninguém (sem efeito)' }
            $h += "<tr><td>$(HE $f.Nome)</td><td class=""num"">$(HE $f.Precedencia)</td><td>$(HE $ap)</td></tr>"
        }
        $h += '</tbody></table>'
        foreach ($f in @($refs)) { $h += "<h3 class=""mod-sub"">$(HE $f.Nome)</h3>" + (TabelaSenha $f.Itens) }
    } elseif (-not @($r.Avisos).Count) {
        $h += '<p class="lead">Nenhuma política refinada (FGPP) no domínio.</p>'
    }
    $h + "<p class=""lead"">$(HE $SenhaBaseline)</p></section>"
}

# ---------------------------------------------------------------------
# GPOs via LDAP e SYSVOL (módulo 4) — sem o módulo GroupPolicy
# ---------------------------------------------------------------------
# Achados na ordem dos cards. Grupo = card em que entra (seguranca = alta, integridade = media, limpeza = baixa).
$GpoAchadosDef = @(
    @{ Id = 'SenhaPreferencia';  Nome = 'Senha gravada em Preferências (cpassword)'; Grupo = 'seguranca'; Tag = 'senha exposta' }
    @{ Id = 'Orfao';            Nome = 'Vínculo para GPO inexistente'; Grupo = 'integridade'; Tag = 'órfão' }
    @{ Id = 'SemSysvol';         Nome = 'Pasta ausente no SYSVOL';      Grupo = 'integridade'; Tag = 'sem SYSVOL' }
    @{ Id = 'SemVinculo';        Nome = 'Sem nenhum vínculo';           Grupo = 'limpeza';     Tag = 'sem vínculo' }
    @{ Id = 'Desativada';        Nome = 'Toda desativada';              Grupo = 'limpeza';     Tag = 'desativada' }
    @{ Id = 'VinculoDesativado'; Nome = 'Vínculo desativado';           Grupo = 'limpeza';     Tag = 'vínculo desativado' }
    @{ Id = 'Vazia';             Nome = 'Vazia (versão 0)';             Grupo = 'limpeza';     Tag = 'vazia' }
)

# Lê o AD e o SYSVOL. Gpos = {Guid; Nome; Flags; Versao; CriadaEm; AlteradaEm; Pasta; PastaExiste ($null = não testada)};
# Alvos = {Dn; GPLink; GPOptions} do domínio, das OUs e dos sites.
function Read-ModGpo($dom) {
    $avisos = @()
    $props = @('displayName', 'flags', 'versionNumber', 'gPCFileSysPath', 'whenCreated', 'whenChanged', 'gPCMachineExtensionNames', 'gPCUserExtensionNames')
    $lidas = @(Get-ADObject -SearchBase "CN=Policies,CN=System,$($dom.DistinguishedName)" -LDAPFilter '(objectClass=groupPolicyContainer)' -Properties $props)
    # gPOptions sem gPLink: OU com herança bloqueada e nenhuma GPO vinculada
    $filtro = '(|(gPLink=*)(gPOptions=*))'
    $alvos = @()
    foreach ($o in @(Get-ADObject -LDAPFilter $filtro -Properties gPLink, gPOptions)) {
        if ($o) { $alvos += New-Object PSObject -Property @{ Dn = [string]$o.DistinguishedName; GPLink = [string]$o.gPLink; GPOptions = $o.gPOptions } }
    }
    try {
        $sites = "CN=Sites,$((Get-ADRootDSE).configurationNamingContext)"
        foreach ($o in @(Get-ADObject -SearchBase $sites -LDAPFilter '(gPLink=*)' -Properties gPLink, gPOptions)) {
            if ($o) { $alvos += New-Object PSObject -Property @{ Dn = [string]$o.DistinguishedName; GPLink = [string]$o.gPLink; GPOptions = $o.gPOptions } }
        }
    } catch { $avisos += "os vínculos nos sites não puderam ser lidos (GPOs vinculadas só a sites aparecem como sem vínculo): $($_.Exception.Message)" }

    $raizSysvol = "\\$($dom.DNSRoot)\SYSVOL\$($dom.DNSRoot)\Policies"
    $sysvolOk = $false
    try { $sysvolOk = [bool](Test-Path -LiteralPath $raizSysvol) } catch { }
    if (-not $sysvolOk) { $avisos += "a pasta $raizSysvol não pôde ser acessada; as pastas das GPOs no SYSVOL não foram verificadas" }

    # Pacotes de Instalação de software ficam no AD, abaixo do contêiner da GPO: "<guid>|Machine" → nomes
    $pacotes = @{}
    try {
        foreach ($o in @(Get-ADObject -SearchBase "CN=Policies,CN=System,$($dom.DistinguishedName)" -LDAPFilter '(objectClass=packageRegistration)' -Properties displayName)) {
            if ($o -and ([string]$o.DistinguishedName) -match ',CN=(Machine|User),CN=(\{[^}]+\}),CN=Policies,') {
                $k = "$($matches[2].ToUpperInvariant())|$($matches[1])"
                if (-not $pacotes.ContainsKey($k)) { $pacotes[$k] = @() }
                $pacotes[$k] += [string]$o.displayName
            }
        }
    } catch { $avisos += "os pacotes de instalação de software das GPOs não puderam ser lidos: $($_.Exception.Message)" }

    # SID → nome, com cache (usado nos direitos de usuário e grupos restritos)
    $cacheSid = @{}
    $resolver = {
        param($sid)
        if (-not $cacheSid.ContainsKey($sid)) {
            $n = $null
            try { $n = (New-Object System.Security.Principal.SecurityIdentifier $sid).Translate([System.Security.Principal.NTAccount]).Value } catch { }
            $cacheSid[$sid] = $n
        }
        $cacheSid[$sid]
    }

    $gpos = @()
    $semConteudo = 0
    foreach ($g in @($lidas)) {
        if (-not $g) { continue }
        $guid = ([string]$g.Name).ToUpperInvariant()
        $pasta = [string]$g.gPCFileSysPath
        $existe = $null
        if ($sysvolOk) {
            $existe = $false
            if ($pasta) { try { $existe = [bool](Test-Path -LiteralPath $pasta) } catch { } }
        }
        $conteudo = $null
        if ($existe) {
            try { $conteudo = Read-ConteudoGpo $pasta $resolver } catch { $conteudo = New-Object PSObject -Property @{ Itens = @(); SenhaEm = @(); Erros = @($_.Exception.Message) } }
            if (@(@($conteudo.Erros) | Where-Object { $_ }).Count) { $semConteudo++ }
        }
        $soft = @()
        foreach ($par in @(@('Machine', 'Computador'), @('User', 'Usuário'))) {
            if ($pacotes.ContainsKey("$guid|$($par[0])")) { $soft += NovoItemGpo $par[1] 'software' 'Instalação de software' @($pacotes["$guid|$($par[0])"] | Sort-Object) }
        }
        $gpos += New-Object PSObject -Property @{ Guid = $guid; Nome = [string]$g.displayName; Flags = [int]$g.flags
            Versao = [int]$g.versionNumber; CriadaEm = $g.whenCreated; AlteradaEm = $g.whenChanged; Pasta = $pasta; PastaExiste = $existe
            ExtMaquina = [string]$g.gPCMachineExtensionNames; ExtUsuario = [string]$g.gPCUserExtensionNames; Conteudo = $conteudo; Software = $soft }
    }
    if ($semConteudo) { $avisos += "parte do conteúdo de $semConteudo $(Plural $semConteudo 'GPO não pôde' 'GPOs não puderam') ser lida no SYSVOL (detalhes no arquivo 06_GPOs, coluna ErrosLeitura)" }
    New-Object PSObject -Property @{ Gpos = $gpos; Alvos = $alvos; SysvolOk = $sysvolOk; Avisos = $avisos }
}

# gPLink → vínculos {Guid; Dn; Opcao; Desativado; Imposto; Externo}. Formato: [LDAP://cn={GUID},<dn>;<op>]...
# Externo = a GPO fica em outro domínio (os DC= do <dn> diferem de $domDN; só avaliado com $domDN).
function ConvertFrom-GPLink([string]$s, [string]$domDN) {
    if (-not $s) { return }
    $pat = '\[LDAP://(?:[^/;\]]+/)?cn=(\{[^}]+\}),([^;\]]+);(\d+)\]'
    foreach ($m in @([regex]::Matches($s, $pat, 'IgnoreCase'))) {
        if (-not $m) { continue }
        $dn = $m.Groups[2].Value
        $op = [int]$m.Groups[3].Value
        # compara só os DC= (um domínio filho também termina no DN do pai)
        $dcs = (@($dn -split '(?<!\\),') | Where-Object { $_ -match '^\s*DC=' } | ForEach-Object { $_.Trim() }) -join ','
        $externo = [bool]($domDN -and $dcs -ne $domDN)
        New-Object PSObject -Property @{ Guid = $m.Groups[1].Value.ToUpperInvariant(); Dn = $dn; Opcao = $op
            Desativado = [bool]($op -band 1); Imposto = [bool]($op -band 2); Externo = $externo }
    }
}

# DN → caminho legível: 'empresa.local/TI/Sub', 'empresa.local' ou 'site Matriz'
function Get-CaminhoDN([string]$dn, [string]$domDN) {
    if ($dn -match '^CN=(.+?),CN=Sites,CN=Configuration,') { return "site $($matches[1].Replace('\,', ','))" }
    $partes = @($dn -split '(?<!\\),')
    $dc = @(); $resto = @()
    foreach ($p in @($partes)) {
        if ($p -match '^DC=(.*)$') { $dc += $matches[1] }
        elseif ($p -match '^[A-Za-z]+=(.*)$') { $resto += $matches[1].Replace('\,', ',') }
    }
    $nomes = @($dc -join '.')
    for ($i = $resto.Count - 1; $i -ge 0; $i--) { $nomes += $resto[$i] }
    $nomes -join '/'
}

# ---- Conteúdo das GPOs: arquivos da pasta no SYSVOL + extensões (CSE) gravadas no AD ----
$GpoLimiteDetalhes = 15

# Extensões do lado do cliente (gPCMachineExtensionNames / gPCUserExtensionNames) → Chave (a mesma dos itens lidos do SYSVOL) e nome da área
$GpoCse = @{
    '{35378EAC-683F-11D2-A89A-00C04FBBCFA2}' = @{ Chave = 'registro';         Nome = 'Modelos administrativos' }
    '{827D319E-6EAC-11D2-A4EA-00C04F79F83A}' = @{ Chave = 'seguranca';        Nome = 'Configurações de segurança' }
    '{F3CCC681-B74C-4060-9F26-CD84525DCA2A}' = @{ Chave = 'auditoria';        Nome = 'Auditoria avançada' }
    '{42B5FAAE-6536-11D2-AE5A-0000F87571E3}' = @{ Chave = 'scripts';          Nome = 'Scripts' }
    '{C6DC5466-785A-11D2-84D0-00C04FB169F7}' = @{ Chave = 'software';         Nome = 'Instalação de software' }
    '{25537BA6-77A8-11D2-9B6C-0000F8080861}' = @{ Chave = 'redirecionamento'; Nome = 'Redirecionamento de pastas' }
    '{B1BE8D72-6EAC-11D2-A4EA-00C04F79F83A}' = @{ Chave = 'efs';              Nome = 'Recuperação de EFS (criptografia de arquivos)' }
    '{E437BC1C-AA7D-11D2-A382-00C04F991E27}' = @{ Chave = 'ipsec';            Nome = 'Segurança IP (IPsec)' }
    '{0ACDD40C-75AC-47AB-BAA0-BF6DE7E7FE63}' = @{ Chave = 'wifi';             Nome = 'Rede sem fio (Wi-Fi)' }
    '{A2E30F80-D7DE-11D2-BBDE-00C04F86AE3B}' = @{ Chave = 'iem';              Nome = 'Manutenção do Internet Explorer' }
    '{B087BE9D-ED37-454F-AF9C-04291E351182}' = @{ Chave = 'pref:Registry';    Nome = 'Preferências: registro' }
    '{5794DAFD-BE60-433F-88A2-1A31939AC01F}' = @{ Chave = 'pref:Drives';      Nome = 'Mapeamento de unidades' }
    '{BC75B1ED-5833-4858-9BB8-CBF0B166DF9D}' = @{ Chave = 'pref:Printers';    Nome = 'Impressoras' }
    '{17D89FEC-5C44-4972-B12D-241CAEF74509}' = @{ Chave = 'pref:Groups';      Nome = 'Usuários e grupos locais' }
    '{AADCED64-746C-4633-A97C-D61349046527}' = @{ Chave = 'pref:ScheduledTasks'; Nome = 'Tarefas agendadas' }
    '{91FBB303-0CD5-4055-BF42-E512A681B325}' = @{ Chave = 'pref:Services';    Nome = 'Serviços (preferências)' }
    '{7150F9BF-48AD-4DA4-A49C-29EF4A8369BA}' = @{ Chave = 'pref:Files';       Nome = 'Arquivos copiados' }
    '{6232C319-91AC-4931-9385-E70C2B099F0E}' = @{ Chave = 'pref:Folders';     Nome = 'Pastas' }
    '{C418DD9D-0D14-4EFB-8FBF-CFE535C8FAC7}' = @{ Chave = 'pref:Shortcuts';   Nome = 'Atalhos' }
    '{0E28E245-9368-4853-AD84-6DA3BA35BB75}' = @{ Chave = 'pref:EnvironmentVariables'; Nome = 'Variáveis de ambiente' }
    '{6A4C88C6-C502-4F74-8F60-2CB23EDC24E2}' = @{ Chave = 'pref:NetworkShares'; Nome = 'Compartilhamentos de rede' }
    '{74EE6C03-5363-4554-B161-627540339CAB}' = @{ Chave = 'pref:IniFiles';    Nome = 'Arquivos .ini' }
    '{728EE579-943C-4519-9EF7-AB56765798ED}' = @{ Chave = 'pref:DataSources'; Nome = 'Fontes de dados ODBC' }
    '{1A6364EB-776B-4120-ADE1-B63A406A76B5}' = @{ Chave = 'pref:Devices';     Nome = 'Dispositivos' }
    '{E62688F0-25FD-4C90-BFF5-F508B9D2E31F}' = @{ Chave = 'pref:PowerOptions'; Nome = 'Opções de energia' }
    '{3A0DBA37-F8B2-4356-83DE-3E90BD5C261F}' = @{ Chave = 'pref:NetworkOptions'; Nome = 'Conexões de rede (VPN/discada)' }
    '{E47248BA-94CC-49C4-BBB5-9EB7F05183D0}' = @{ Chave = 'pref:InternetSettings'; Nome = 'Configurações de Internet' }
    '{A3F3E39B-5D83-4940-B954-28315B82F0A8}' = @{ Chave = 'pref:FolderOptions'; Nome = 'Opções de pasta' }
    '{E5094040-C46C-4115-B030-04FB2E545B00}' = @{ Chave = 'pref:RegionalOptions'; Nome = 'Opções regionais' }
    '{E4F48E54-F38D-4884-BFB9-D4D2E5729C18}' = @{ Chave = 'pref:StartMenu';   Nome = 'Menu Iniciar' }
    '{F9C77450-3A41-477E-9310-9ACD617BD9E3}' = @{ Chave = 'pref:Applications'; Nome = 'Aplicativos' }
}
# Pasta em Preferences\ → nome da área (a mesma da tabela de CSE)
$GpoPreferencias = @{
    Registry = 'Preferências: registro'; Drives = 'Mapeamento de unidades'; Printers = 'Impressoras'; Groups = 'Usuários e grupos locais'
    ScheduledTasks = 'Tarefas agendadas'; Services = 'Serviços (preferências)'; Files = 'Arquivos copiados'; Folders = 'Pastas'; Shortcuts = 'Atalhos'
    EnvironmentVariables = 'Variáveis de ambiente'; NetworkShares = 'Compartilhamentos de rede'; IniFiles = 'Arquivos .ini'; DataSources = 'Fontes de dados ODBC'
    Devices = 'Dispositivos'; PowerOptions = 'Opções de energia'; NetworkOptions = 'Conexões de rede (VPN/discada)'; InternetSettings = 'Configurações de Internet'
    FolderOptions = 'Opções de pasta'; RegionalOptions = 'Opções regionais'; StartMenu = 'Menu Iniciar'; Applications = 'Aplicativos'
}
# GptTmpl.inf: seção → área, na ordem em que aparecem
$GpoAreasSeg = @(
    @{ Secao = 'System Access';            Area = 'Política de contas (senha e bloqueio)' }
    @{ Secao = 'Kerberos Policy';          Area = 'Política de Kerberos' }
    @{ Secao = 'Event Audit';              Area = 'Auditoria' }
    @{ Secao = 'Privilege Rights';         Area = 'Direitos de usuário' }
    @{ Secao = 'Registry Values';          Area = 'Opções de segurança' }
    @{ Secao = 'Group Membership';         Area = 'Grupos restritos' }
    @{ Secao = 'Service General Setting';  Area = 'Serviços do sistema' }
    @{ Secao = 'File Security';            Area = 'Permissões de arquivos e pastas' }
    @{ Secao = 'Registry Keys';            Area = 'Permissões do registro' }
)
$GpoContas = @{
    MinimumPasswordLength = 'Tamanho mínimo da senha'; PasswordComplexity = 'Complexidade de senha (1 = exigida)'; PasswordHistorySize = 'Histórico de senhas'
    MaximumPasswordAge = 'Validade máxima da senha (dias)'; MinimumPasswordAge = 'Idade mínima da senha (dias)'; ClearTextPassword = 'Criptografia reversível (1 = ligada)'
    LockoutBadCount = 'Bloquear após tentativas inválidas'; LockoutDuration = 'Duração do bloqueio (min)'; ResetLockoutCount = 'Zerar contador de tentativas após (min)'
    RequireLogonToChangePassword = 'Exigir logon para trocar a senha'; ForceLogoffWhenHourExpire = 'Desconectar quando o horário de logon terminar'
    NewAdministratorName = 'Renomear a conta Administrador para'; NewGuestName = 'Renomear a conta Convidado para'
    EnableAdminAccount = 'Conta Administrador ativa (1 = sim)'; EnableGuestAccount = 'Conta Convidado ativa (1 = sim)'; LSAAnonymousNameLookup = 'Tradução anônima de SID/nome'
}
$GpoAuditoria = @{
    AuditSystemEvents = 'Eventos de sistema'; AuditLogonEvents = 'Eventos de logon'; AuditObjectAccess = 'Acesso a objetos'; AuditPrivilegeUse = 'Uso de privilégios'
    AuditPolicyChange = 'Alteração de políticas'; AuditAccountManage = 'Gerenciamento de contas'; AuditProcessTracking = 'Acompanhamento de processos'
    AuditDSAccess = 'Acesso ao serviço de diretório'; AuditAccountLogon = 'Eventos de logon de conta'
}
$GpoAuditoriaValor = @{ '0' = 'sem auditoria'; '1' = 'sucesso'; '2' = 'falha'; '3' = 'sucesso e falha' }
$GpoDireitos = @{
    SeInteractiveLogonRight = 'Permitir logon local'; SeDenyInteractiveLogonRight = 'Negar logon local'
    SeRemoteInteractiveLogonRight = 'Permitir logon pela Área de Trabalho Remota'; SeDenyRemoteInteractiveLogonRight = 'Negar logon pela Área de Trabalho Remota'
    SeNetworkLogonRight = 'Acessar este computador pela rede'; SeDenyNetworkLogonRight = 'Negar acesso a este computador pela rede'
    SeServiceLogonRight = 'Fazer logon como serviço'; SeDenyServiceLogonRight = 'Negar logon como serviço'
    SeBatchLogonRight = 'Fazer logon como tarefa em lotes'; SeDenyBatchLogonRight = 'Negar logon como tarefa em lotes'
    SeBackupPrivilege = 'Fazer backup de arquivos e pastas'; SeRestorePrivilege = 'Restaurar arquivos e pastas'
    SeShutdownPrivilege = 'Desligar o sistema'; SeRemoteShutdownPrivilege = 'Forçar o desligamento a partir de um sistema remoto'
    SeDebugPrivilege = 'Depurar programas'; SeTakeOwnershipPrivilege = 'Apropriar-se de arquivos ou objetos'; SeSecurityPrivilege = 'Gerenciar auditoria e log de segurança'
    SeSystemtimePrivilege = 'Alterar a hora do sistema'; SeTimeZonePrivilege = 'Alterar o fuso horário'; SeMachineAccountPrivilege = 'Adicionar estações de trabalho ao domínio'
    SeImpersonatePrivilege = 'Representar um cliente após a autenticação'; SeLoadDriverPrivilege = 'Carregar e descarregar drivers'
    SeIncreaseQuotaPrivilege = 'Ajustar cotas de memória de um processo'; SeAssignPrimaryTokenPrivilege = 'Substituir um token de processo'
    SeEnableDelegationPrivilege = 'Confiar em contas para delegação'; SeChangeNotifyPrivilege = 'Ignorar a verificação completa'
    SeCreateGlobalPrivilege = 'Criar objetos globais'; SeManageVolumePrivilege = 'Executar tarefas de manutenção de volume'
    SeSystemEnvironmentPrivilege = 'Modificar valores de ambiente de firmware'; SeProfileSingleProcessPrivilege = 'Analisar um único processo'
    SeUndockPrivilege = 'Remover o computador da estação de acoplamento'; SeIncreaseBasePriorityPrivilege = 'Aumentar a prioridade de agendamento'
}
# Opções de segurança ([Registry Values]): último trecho da chave → nome
$GpoOpcoesSeg = @{
    LmCompatibilityLevel = 'Nível de autenticação LAN Manager'; NoLMHash = 'Não armazenar o hash LM'; RestrictAnonymous = 'Restringir acesso anônimo'
    RestrictAnonymousSAM = 'Restringir enumeração anônima de contas'; EveryoneIncludesAnonymous = 'Todos inclui anônimos'
    EnableLUA = 'Controle de Conta de Usuário (UAC) ligado'; ConsentPromptBehaviorAdmin = 'UAC: aviso para administradores'; ConsentPromptBehaviorUser = 'UAC: aviso para usuários comuns'
    FilterAdministratorToken = 'UAC: modo de aprovação para o Administrador interno'; DontDisplayLastUserName = 'Não exibir o último usuário no logon'
    DisableCAD = 'Dispensar Ctrl+Alt+Del'; InactivityTimeoutSecs = 'Bloquear após inatividade (segundos)'; LegalNoticeCaption = 'Título do aviso de logon'
    LegalNoticeText = 'Texto do aviso de logon'; CachedLogonsCount = 'Logons armazenados em cache'; ScRemoveOption = 'Ação ao remover o cartão inteligente'
    RequireSecuritySignature = 'Exigir assinatura SMB'; EnableSecuritySignature = 'Assinar SMB quando possível'; EnablePlainTextPassword = 'Enviar senha sem criptografia a servidores SMB'
    LDAPClientIntegrity = 'Assinatura LDAP (cliente)'; LDAPServerIntegrity = 'Assinatura LDAP (servidor)'; RequireSignOrSeal = 'Netlogon: exigir assinatura ou criptografia'
    RequireStrongKey = 'Netlogon: exigir chave de sessão forte'; DisablePasswordChange = 'Netlogon: desativar troca de senha da conta de computador'
    MaximumPasswordAge = 'Netlogon: idade máxima da senha da conta de computador (dias)'; NoConnectedUser = 'Bloquear contas da Microsoft'
    SupportedEncryptionTypes = 'Tipos de criptografia permitidos no Kerberos'; NTLMMinClientSec = 'Segurança mínima de NTLM (cliente)'; NTLMMinServerSec = 'Segurança mínima de NTLM (servidor)'
    ShutdownWithoutLogon = 'Permitir desligar sem fazer logon'; AutoDisconnect = 'Tempo ocioso antes de desconectar sessão SMB (min)'; NullSessionShares = 'Compartilhamentos acessíveis anonimamente'
}
$GpoInicioServico = @{ '2' = 'automático'; '3' = 'manual'; '4' = 'desativado' }
# Modelos administrativos: chave do registro → área (testados em ordem)
$GpoAreasRegistro = @(
    @{ P = '\\WindowsUpdate';                         N = 'Windows Update' }
    @{ P = 'WindowsFirewall';                         N = 'Firewall do Windows' }
    @{ P = 'Windows Defender|Microsoft Antimalware';  N = 'Microsoft Defender' }
    @{ P = 'Terminal Services';                       N = 'Área de Trabalho Remota' }
    @{ P = '\\FVE\b|BitLocker';                       N = 'BitLocker' }
    @{ P = '\\Office\\|\\Outlook\\';                  N = 'Microsoft Office' }
    @{ P = 'Internet Explorer';                       N = 'Internet Explorer' }
    @{ P = 'Microsoft\\Edge';                         N = 'Microsoft Edge' }
    @{ P = 'Google\\Chrome';                          N = 'Google Chrome' }
    @{ P = 'Mozilla\\Firefox';                        N = 'Mozilla Firefox' }
    @{ P = 'OneDrive';                                N = 'OneDrive' }
    @{ P = 'PowerShell';                              N = 'PowerShell' }
    @{ P = 'WinRM';                                   N = 'Gerenciamento remoto (WinRM)' }
    @{ P = 'DeviceInstall|RemovableStorageDevices';   N = 'Dispositivos e armazenamento removível' }
    @{ P = 'Control Panel\\Desktop|Personalization';  N = 'Área de trabalho e proteção de tela' }
    @{ P = 'Windows Search';                          N = 'Pesquisa do Windows' }
    @{ P = 'DataCollection|CloudContent|AdvertisingInfo'; N = 'Privacidade e telemetria' }
    @{ P = 'Power\\|PowerSettings';                   N = 'Energia' }
    @{ P = '\\Printers';                              N = 'Impressoras' }
    @{ P = 'EventLog';                                N = 'Logs de eventos' }
    @{ P = 'WindowsStore';                            N = 'Microsoft Store' }
    @{ P = '\\Safer\\';                               N = 'Restrição de software' }
    @{ P = 'NetworkConnections|Lanman|DNSClient|TCPIP|NetworkProvider|NetBT|Netlogon'; N = 'Rede' }
    @{ P = 'Policies\\Explorer|\\Windows\\Explorer';  N = 'Explorador de Arquivos e menu Iniciar' }
    @{ P = 'Policies\\System\b|\\Windows\\System\b|Winlogon'; N = 'Sistema e logon' }
)
# SIDs conhecidos (usados quando a tradução pelo Windows falha)
$GpoSidsConhecidos = @{
    'S-1-1-0' = 'Everyone'; 'S-1-5-2' = 'NETWORK'; 'S-1-5-4' = 'INTERACTIVE'; 'S-1-5-6' = 'SERVICE'; 'S-1-5-9' = 'Enterprise Domain Controllers'
    'S-1-5-11' = 'Authenticated Users'; 'S-1-5-19' = 'LOCAL SERVICE'; 'S-1-5-20' = 'NETWORK SERVICE'; 'S-1-5-113' = 'Local account'
    'S-1-5-114' = 'Local account and member of Administrators group'; 'S-1-5-32-544' = 'Administrators'; 'S-1-5-32-545' = 'Users'; 'S-1-5-32-546' = 'Guests'
    'S-1-5-32-548' = 'Account Operators'; 'S-1-5-32-549' = 'Server Operators'; 'S-1-5-32-550' = 'Print Operators'; 'S-1-5-32-551' = 'Backup Operators'
    'S-1-5-32-555' = 'Remote Desktop Users'
}

# Primeiro GUID de cada [ ] de gPC*ExtensionNames = a extensão (CSE); os outros são ferramentas do editor
function Get-CseGpo([string]$s) {
    if (-not $s) { return }
    foreach ($m in @([regex]::Matches($s, '\[(\{[0-9A-Fa-f-]{36}\})'))) {
        if (-not $m) { continue }
        $g = $m.Groups[1].Value.ToUpperInvariant()
        if ($g -ne '{00000000-0000-0000-0000-000000000000}') { $g }
    }
}

# Item de conteúdo: Lado = Computador/Usuário; Cse = chave da área ($GpoCse); Detalhes = uma linha por configuração
function NovoItemGpo([string]$lado, [string]$cse, [string]$area, $detalhes) {
    New-Object PSObject -Property @{ Lado = $lado; Cse = $cse; Area = $area; Detalhes = @(@($detalhes) | Where-Object { $_ }) }
}

function Curto80([string]$s) { if ($s.Length -gt 80) { $s.Substring(0, 77) + '...' } else { $s } }

# .ini → {Secao; Chave; Valor} (linha sem '=' vira Chave inteira, como nas seções de serviços e permissões do GptTmpl.inf)
function ConvertFrom-Ini($linhas) {
    $sec = ''
    foreach ($l in @($linhas)) {
        $t = ([string]$l).Trim()
        if (-not $t -or $t.StartsWith(';')) { continue }
        if ($t -match '^\[(.+)\]$') { $sec = $matches[1]; continue }
        $i = $t.IndexOf('=')
        if ($i -ge 0) { $k = $t.Substring(0, $i).Trim(); $v = $t.Substring($i + 1).Trim() } else { $k = $t; $v = '' }
        New-Object PSObject -Property @{ Secao = $sec; Chave = $k; Valor = $v }
    }
}

# '*S-1-5-32-544' → nome (pelo $resolver, senão pela tabela, senão o próprio SID); nomes sem '*' ficam como estão
function NomeSidGpo([string]$s, [scriptblock]$resolver) {
    $s = $s.Trim().Trim('"')
    if (-not $s.StartsWith('*')) { return $s }
    $sid = $s.Substring(1)
    if ($resolver) { $n = & $resolver $sid; if ($n) { return [string]$n } }
    if ($GpoSidsConhecidos.ContainsKey($sid)) { return $GpoSidsConhecidos[$sid] }
    $sid
}

function NomesSidGpo([string]$lista, [scriptblock]$resolver) {
    $n = @()
    foreach ($x in @($lista -split ',')) { if ($x.Trim()) { $n += NomeSidGpo $x $resolver } }
    if ($n.Count) { $n -join ', ' } else { '(ninguém)' }
}

function Add-DetalheGpo($det, [string]$area, [string]$texto) {
    if (-not $texto) { return }
    if (-not $det.ContainsKey($area)) { $det[$area] = New-Object System.Collections.ArrayList }
    [void]$det[$area].Add($texto)
}

# GptTmpl.inf (Configurações de segurança, só do lado Computador) → itens por área (pura)
function ConvertFrom-GptTmpl($linhas, [scriptblock]$resolver) {
    $areaDe = @{}
    foreach ($a in @($GpoAreasSeg)) { $areaDe[$a.Secao] = $a.Area }
    $det = @{}
    foreach ($e in @(ConvertFrom-Ini $linhas)) {
        if (-not $e -or -not $areaDe.ContainsKey($e.Secao)) { continue }
        $sec = $e.Secao; $k = $e.Chave; $v = $e.Valor; $t = $null
        if ($sec -eq 'System Access') {
            $nome = if ($GpoContas.ContainsKey($k)) { $GpoContas[$k] } else { $k }
            $t = "${nome}: $($v.Trim('"'))"
        } elseif ($sec -eq 'Kerberos Policy') {
            $t = "${k}: $v"
        } elseif ($sec -eq 'Event Audit') {
            $nome = if ($GpoAuditoria.ContainsKey($k)) { $GpoAuditoria[$k] } else { $k }
            $val = if ($GpoAuditoriaValor.ContainsKey($v)) { $GpoAuditoriaValor[$v] } else { $v }
            $t = "${nome}: $val"
        } elseif ($sec -eq 'Privilege Rights') {
            $nome = if ($GpoDireitos.ContainsKey($k)) { $GpoDireitos[$k] } else { $k }
            $t = "${nome}: $(NomesSidGpo $v $resolver)"
        } elseif ($sec -eq 'Registry Values') {
            $partes = @($k -split '\\')
            $ult = $partes[$partes.Count - 1]
            $nome = if ($GpoOpcoesSeg.ContainsKey($ult)) { $GpoOpcoesSeg[$ult] } else { $ult }
            if ($ult -match 'SecuritySignature$') { if ($k -match 'LanmanServer') { $nome += ' (servidor)' } elseif ($k -match 'LanManWorkstation') { $nome += ' (cliente)' } }
            $val = if ($v -match '^\d+,(.*)$') { $matches[1].Trim('"') } else { $v }
            $t = "${nome}: $(Curto80 $val)"
        } elseif ($sec -eq 'Group Membership') {
            if ($k -match '^(.+?)__(Members|Memberof)$') {
                $sidGrupo = $matches[1]; $tipo = $matches[2]
                $grupo = NomeSidGpo $sidGrupo $resolver
                if ($tipo -eq 'Members') { $t = "${grupo}: membros definidos = $(NomesSidGpo $v $resolver)" }
                elseif ($v) { $t = "${grupo}: passa a ser membro de $(NomesSidGpo $v $resolver)" }
            }
        } elseif ($sec -eq 'Service General Setting') {
            if ($k -match '^"?([^",]+)"?\s*,\s*(\d)') {
                $ini = if ($GpoInicioServico.ContainsKey($matches[2])) { $GpoInicioServico[$matches[2]] } else { "modo $($matches[2])" }
                $t = "$($matches[1]): $ini"
            }
        } else {
            # File Security / Registry Keys: só o caminho
            if ($k -match '^"([^"]+)"') { $t = $matches[1] }
        }
        Add-DetalheGpo $det $areaDe[$sec] $t
    }
    foreach ($a in @($GpoAreasSeg)) {
        if ($det.ContainsKey($a.Area)) { NovoItemGpo 'Computador' 'seguranca' $a.Area @($det[$a.Area]) }
    }
}

# Registry.pol (formato PReg) → {Chave; Valor; Tipo; Dado} (pura)
function ConvertFrom-RegistryPol([byte[]]$b) {
    if (-not $b -or $b.Length -lt 8 -or [Text.Encoding]::ASCII.GetString($b, 0, 4) -ne 'PReg') { throw 'Registry.pol em formato desconhecido' }
    $u = [Text.Encoding]::Unicode
    $i = 8; $n = $b.Length
    while ($i + 2 -le $n) {
        if ($b[$i] -ne 0x5B -or $b[$i + 1] -ne 0) { break }
        $i += 2
        $campos = @()
        for ($c = 0; $c -lt 2; $c++) {
            $ini = $i
            while ($i + 1 -lt $n -and -not ($b[$i] -eq 0 -and $b[$i + 1] -eq 0)) { $i += 2 }
            $campos += $u.GetString($b, $ini, $i - $ini)
            $i += 4   # nulo + ';'
        }
        if ($i + 10 -gt $n) { break }
        $tipo = [int][BitConverter]::ToUInt32($b, $i); $i += 6
        $tam = [int][BitConverter]::ToUInt32($b, $i); $i += 6
        $dado = ''
        if ($tam -gt 0 -and $i + $tam -le $n) {
            if ($tipo -eq 4 -and $tam -ge 4) { $dado = [string][BitConverter]::ToUInt32($b, $i) }
            elseif ($tipo -eq 11 -and $tam -ge 8) { $dado = [string][BitConverter]::ToUInt64($b, $i) }
            elseif ($tipo -eq 1 -or $tipo -eq 2 -or $tipo -eq 7) { $dado = $u.GetString($b, $i, $tam).TrimEnd([char]0).Replace([string][char]0, '; ') }
            else { $dado = "($tam bytes)" }
        }
        $i += $tam + 2   # dado + ']'
        New-Object PSObject -Property @{ Chave = $campos[0]; Valor = $campos[1]; Tipo = $tipo; Dado = $dado }
    }
}

function Get-AreaRegistro([string]$chave) {
    foreach ($a in @($GpoAreasRegistro)) { if ($chave -match $a.P) { return $a.N } }
    if ($chave -match 'Policies\\(?:Microsoft\\)?(?:Windows\\)?([^\\]+)') { return $matches[1] }
    'Outras'
}

# Entradas do Registry.pol → um item por área ("Modelos administrativos: Windows Update") (pura)
function ConvertTo-ItensRegistro($entradas, [string]$lado) {
    $det = @{}; $ordem = New-Object System.Collections.ArrayList
    foreach ($e in @($entradas)) {
        if (-not $e) { continue }
        $area = "Modelos administrativos: $(Get-AreaRegistro $e.Chave)"
        if (-not $det.ContainsKey($area)) { [void]$ordem.Add($area) }
        $chave = $e.Chave -replace '^(?i)Software\\Policies\\(?:Microsoft\\)?(?:Windows\\)?', ''
        $v = [string]$e.Valor
        if ($v -match '^\*\*del\.(.+)$') { $t = "$chave\$($matches[1]): removido" }
        elseif ($v -match '^\*\*delvals') { $t = "${chave}: todos os valores removidos" }
        elseif ($v.StartsWith('**')) { $t = "${chave}: $v" }
        else { $t = "$chave\$(if ($v) { $v } else { '(padrão)' }) = $(Curto80 $e.Dado)" }
        Add-DetalheGpo $det $area $t
    }
    foreach ($a in @($ordem)) { NovoItemGpo $lado 'registro' $a @($det[$a]) }
}

# scripts.ini / psscripts.ini → linhas "Logon: script.bat parâmetros" (pura)
function ConvertFrom-ScriptsIni($linhas, [string]$sufixo) {
    $nomes = @{ Startup = 'Inicialização'; Shutdown = 'Desligamento'; Logon = 'Logon'; Logoff = 'Logoff' }
    $cmd = @{}; $par = @{}; $ordem = New-Object System.Collections.ArrayList
    foreach ($e in @(ConvertFrom-Ini $linhas)) {
        if (-not $e -or $e.Chave -notmatch '^(\d+)(CmdLine|Parameters)$') { continue }
        $id = "$($e.Secao)|$($matches[1])"
        if ($matches[2] -eq 'CmdLine') { $cmd[$id] = $e.Valor; [void]$ordem.Add($id) } else { $par[$id] = $e.Valor }
    }
    foreach ($id in @($ordem)) {
        $sec = ($id -split '\|')[0]
        $quando = if ($nomes.ContainsKey($sec)) { $nomes[$sec] } else { $sec }
        $t = "${quando}${sufixo}: $($cmd[$id])"
        if ($par[$id]) { $t += " $($par[$id])" }
        $t
    }
}

function AtributoGpo($no, [string]$nome) { if ($no) { [string]$no.GetAttribute($nome) } else { '' } }

# XML de Preferências → {Detalhes; Senha ([bool] algum cpassword preenchido)} (pura). O valor do cpassword nunca é lido para o relatório.
function ConvertFrom-PreferenciaXml($x, [string]$tipo) {
    $acoes = @{ C = 'criar'; R = 'substituir'; U = 'atualizar'; D = 'excluir' }
    $det = @()
    foreach ($n in @($x.SelectNodes('//*[Properties]'))) {
        if (-not $n) { continue }
        $p = $n.SelectSingleNode('Properties')
        $nome = AtributoGpo $n 'name'
        $t = $nome
        if ($tipo -eq 'Drives') {
            $l = AtributoGpo $p 'letter'
            $t = "$(if ($l) { "${l}:" } else { $nome }) → $(AtributoGpo $p 'path')"
        } elseif ($tipo -eq 'Printers') {
            $t = AtributoGpo $p 'path'
            if (-not $t) { $t = AtributoGpo $p 'ipAddress' }
            if (-not $t) { $t = $nome }
            if ((AtributoGpo $p 'default') -eq '1') { $t += ' (padrão)' }
        } elseif ($tipo -eq 'Groups' -and $n.LocalName -eq 'Group') {
            $g = AtributoGpo $p 'groupName'; if (-not $g) { $g = $nome }
            $add = @(); $rem = @()
            foreach ($m in @($p.SelectNodes('Members/Member'))) {
                if (-not $m) { continue }
                if ((AtributoGpo $m 'action') -eq 'REMOVE') { $rem += AtributoGpo $m 'name' } else { $add += AtributoGpo $m 'name' }
            }
            $partes = @()
            if ($add.Count) { $partes += "adiciona $($add -join ', ')" }
            if ($rem.Count) { $partes += "remove $($rem -join ', ')" }
            if ((AtributoGpo $p 'deleteAllUsers') -eq '1') { $partes += 'remove os demais usuários' }
            if ((AtributoGpo $p 'deleteAllGroups') -eq '1') { $partes += 'remove os demais grupos' }
            $t = "grupo local ${g}$(if ($partes.Count) { ': ' + ($partes -join '; ') })"
        } elseif ($tipo -eq 'Groups') {
            $u = AtributoGpo $p 'userName'; if (-not $u) { $u = $nome }
            $t = "usuário local $u"
            if (AtributoGpo $p 'newName') { $t += " → renomear para $(AtributoGpo $p 'newName')" }
            if ((AtributoGpo $p 'acctDisabled') -eq '1') { $t += ' (desativado)' }
            if (AtributoGpo $p 'cpassword') { $t += ' (define a senha)' }
        } elseif ($tipo -eq 'ScheduledTasks') {
            $cmdTarefa = AtributoGpo $p 'appName'
            if (-not $cmdTarefa) { $c = $p.SelectSingleNode('Task/Actions/Exec/Command'); if ($c) { $cmdTarefa = $c.InnerText } }
            if ($cmdTarefa) { $t += ": $cmdTarefa" }
            $ra = AtributoGpo $p 'runAs'
            if (-not $ra) { $c = $p.SelectSingleNode('Task/Principals/Principal/UserId'); if ($c) { $ra = $c.InnerText } }
            if ($ra) { $t += " (executa como $ra)" }
        } elseif ($tipo -eq 'Files') {
            $t = "$(AtributoGpo $p 'fromPath') → $(AtributoGpo $p 'targetPath')"
        } elseif ($tipo -eq 'Folders') {
            $t = AtributoGpo $p 'path'
        } elseif ($tipo -eq 'Shortcuts') {
            $t = "$nome → $(AtributoGpo $p 'targetPath')"
        } elseif ($tipo -eq 'Registry') {
            $vn = AtributoGpo $p 'name'; if (-not $vn) { $vn = '(padrão)' }
            $t = "$(AtributoGpo $p 'hive')\$(AtributoGpo $p 'key')\$vn = $(Curto80 (AtributoGpo $p 'value'))"
        } elseif ($tipo -eq 'EnvironmentVariables') {
            $t = "$(AtributoGpo $p 'name') = $(AtributoGpo $p 'value')"
        } elseif ($tipo -eq 'Services') {
            $t = "$(AtributoGpo $p 'serviceName'): $((AtributoGpo $p 'startupType').ToLowerInvariant())"
        } elseif ($tipo -eq 'NetworkShares') {
            $t = "$(AtributoGpo $p 'name') → $(AtributoGpo $p 'path')"
        } elseif ($tipo -eq 'IniFiles') {
            $t = "$(AtributoGpo $p 'path') [$(AtributoGpo $p 'section')] $(AtributoGpo $p 'property') = $(AtributoGpo $p 'value')"
        }
        $ac = AtributoGpo $p 'action'
        if ($ac -and $ac -ne 'U' -and $acoes.ContainsKey($ac)) { $t += " ($($acoes[$ac]))" }
        if ($t.Length -gt 160) { $t = $t.Substring(0, 157) + '...' }
        $det += $t
    }
    $senha = @(@($x.SelectNodes('//*[@cpassword]')) | Where-Object { $_ -and $_.GetAttribute('cpassword') }).Count -gt 0
    New-Object PSObject -Property @{ Detalhes = $det; Senha = [bool]$senha }
}

# fdeploy1.ini (Redirecionamento de pastas) → linhas "Pasta → destino" (pura)
function ConvertFrom-Fdeploy($linhas) {
    $ini = @(ConvertFrom-Ini $linhas)
    foreach ($e in @($ini | Where-Object { $_ -and $_.Secao -eq 'FolderStatus' })) {
        $dest = @($ini | Where-Object { $_ -and $_.Secao -eq $e.Chave } | Select-Object -First 1)
        if ($dest.Count -and $dest[0]) { "$($e.Chave) → $($dest[0].Valor)" } else { $e.Chave }
    }
}

# Lê a pasta da GPO no SYSVOL → {Itens; SenhaEm (arquivos com senha em Preferências); Erros}. Cada arquivo em try próprio.
function Read-ConteudoGpo([string]$pasta, [scriptblock]$resolver) {
    $itens = @(); $senha = @(); $erros = @()
    foreach ($par in @(@('Machine', 'Computador'), @('User', 'Usuário'))) {
        $base = Join-Path $pasta $par[0]; $lado = $par[1]
        if (-not (Test-Path -LiteralPath $base)) { continue }
        $arq = { param($rel) $f = Join-Path $base $rel; if (Test-Path -LiteralPath $f) { $f } }

        $f = & $arq 'Microsoft\Windows NT\SecEdit\GptTmpl.inf'
        if ($f) { try { $itens += @(ConvertFrom-GptTmpl @(Get-Content -LiteralPath $f -ErrorAction Stop) $resolver) } catch { $erros += "GptTmpl.inf: $($_.Exception.Message)" } }

        $f = & $arq 'Microsoft\Windows NT\Audit\audit.csv'
        if ($f) {
            try {
                $det = @(Import-Csv -LiteralPath $f | Where-Object { $_.Subcategory } | ForEach-Object { "$($_.Subcategory): $($_.'Inclusion Setting')" })
                if ($det.Count) { $itens += NovoItemGpo $lado 'auditoria' 'Auditoria avançada' $det }
            } catch { $erros += "audit.csv: $($_.Exception.Message)" }
        }

        $f = & $arq 'Registry.pol'
        if ($f) { try { $itens += @(ConvertTo-ItensRegistro @(ConvertFrom-RegistryPol ([IO.File]::ReadAllBytes($f))) $lado) } catch { $erros += "$($par[0])\Registry.pol: $($_.Exception.Message)" } }

        $det = @()
        foreach ($s in @(@('Scripts\scripts.ini', ''), @('Scripts\psscripts.ini', ' (PowerShell)'))) {
            $f = & $arq $s[0]
            if ($f) { try { $det += @(ConvertFrom-ScriptsIni @(Get-Content -LiteralPath $f -ErrorAction Stop) $s[1]) } catch { $erros += "$($s[0]): $($_.Exception.Message)" } }
        }
        if ($det.Count) { $itens += NovoItemGpo $lado 'scripts' 'Scripts' $det }

        $pref = & $arq 'Preferences'
        if ($pref) {
            foreach ($d in @(Get-ChildItem -LiteralPath $pref | Where-Object { $_.PSIsContainer })) {
                $f = Join-Path $d.FullName "$($d.Name).xml"
                if (-not (Test-Path -LiteralPath $f)) { continue }
                try {
                    $x = New-Object System.Xml.XmlDocument
                    $x.Load($f)
                    $r = ConvertFrom-PreferenciaXml $x $d.Name
                    $area = if ($GpoPreferencias.ContainsKey($d.Name)) { $GpoPreferencias[$d.Name] } else { "Preferências: $($d.Name)" }
                    if (@($r.Detalhes).Count) { $itens += NovoItemGpo $lado "pref:$($d.Name)" $area $r.Detalhes }
                    if ($r.Senha) { $senha += "$($par[0])\Preferences\$($d.Name)\$($d.Name).xml" }
                } catch { $erros += "Preferences\$($d.Name): $($_.Exception.Message)" }
            }
        }

        $f = & $arq 'Documents & Settings\fdeploy1.ini'
        if (-not $f) { $f = & $arq 'Documents & Settings\fdeploy.ini' }
        if ($f) {
            try {
                $det = @(ConvertFrom-Fdeploy @(Get-Content -LiteralPath $f -ErrorAction Stop))
                if ($det.Count) { $itens += NovoItemGpo $lado 'redirecionamento' 'Redirecionamento de pastas' $det }
            } catch { $erros += "fdeploy.ini: $($_.Exception.Message)" }
        }
    }
    New-Object PSObject -Property @{ Itens = $itens; SenhaEm = $senha; Erros = $erros }
}

# Itens de uma GPO: lidos do SYSVOL + software (AD) + áreas que só aparecem nas extensões (sem detalhe). Computador antes de Usuário. (pura)
function Get-ItensGpo($g) {
    $todos = @()
    if ($g.Conteudo) { $todos += @(@($g.Conteudo.Itens) | Where-Object { $_ }) }
    $todos += @(@($g.Software) | Where-Object { $_ })
    $cobertas = @{}
    foreach ($i in @($todos)) { $cobertas["$($i.Lado)|$($i.Cse)"] = 1 }
    foreach ($par in @(@([string]$g.ExtMaquina, 'Computador'), @([string]$g.ExtUsuario, 'Usuário'))) {
        foreach ($cse in @(Get-CseGpo $par[0])) {
            if (-not $cse -or -not $GpoCse.ContainsKey($cse)) { continue }
            $d = $GpoCse[$cse]
            $k = "$($par[1])|$($d.Chave)"
            if ($cobertas.ContainsKey($k)) { continue }
            $cobertas[$k] = 1
            $todos += NovoItemGpo $par[1] $d.Chave $d.Nome @()
        }
    }
    @(@($todos) | Where-Object { $_.Lado -eq 'Computador' }) + @(@($todos) | Where-Object { $_.Lado -ne 'Computador' })
}

# Vinculada na raiz do domínio (onde a política de senha passa a valer para as contas do domínio)
function Test-VinculoRaiz($g) {
    [bool]@(@($g.Vinculos) | Where-Object { $_ -and -not $_.Desativado -and $_.Caminho -notmatch '/' -and $_.Caminho -notmatch '^site ' }).Count
}

function Add-AchadoGpo($achados, [string]$id, $x) {
    if (-not $achados.ContainsKey($id)) { $achados[$id] = @() }
    $achados[$id] += $x
}

# Análise (pura). Devolve o resultado do módulo: Gpos (com Vinculos e Achados), Orfaos, Bloqueios, Externos e Achados @{ Id = lista }.
function Get-AchadosGpo($gpos, $alvosVinculo, $domDN, $sysvolOk) {
    $porGuid = @{}
    foreach ($g in @($gpos)) {
        if (-not $g) { continue }
        $porGuid[$g.Guid] = New-Object PSObject -Property @{ Guid = $g.Guid; Nome = $(if ($g.Nome) { $g.Nome } else { $g.Guid }); Flags = [int]$g.Flags
            Status = $SnapGpoStatus[[string][int]$g.Flags]; Versao = [int]$g.Versao; CriadaEm = $g.CriadaEm; AlteradaEm = $g.AlteradaEm
            Pasta = $g.Pasta; PastaExiste = $g.PastaExiste; Vinculos = @(); Achados = @()
            Itens = @(Get-ItensGpo $g); ConteudoLido = [bool]$g.Conteudo
            SenhaEm = $(if ($g.Conteudo) { @(@($g.Conteudo.SenhaEm) | Where-Object { $_ }) } else { @() })
            ErrosLeitura = $(if ($g.Conteudo) { @(@($g.Conteudo.Erros) | Where-Object { $_ }) } else { @() }) }
    }

    # Vínculos de cada alvo (domínio, OUs, sites) e bloqueio de herança
    $orfaos = @{}; $bloqueios = @(); $externos = @()
    foreach ($a in @(@($alvosVinculo) | Where-Object { $_ } | Sort-Object Dn)) {
        $caminho = Get-CaminhoDN $a.Dn $domDN
        if ([int]$a.GPOptions -band 1) { $bloqueios += $caminho }
        foreach ($v in @(ConvertFrom-GPLink $a.GPLink $domDN)) {
            if (-not $v) { continue }
            $lig = New-Object PSObject -Property @{ Caminho = $caminho; Desativado = $v.Desativado; Imposto = $v.Imposto }
            if ($porGuid.ContainsKey($v.Guid)) { $porGuid[$v.Guid].Vinculos += $lig }
            elseif ($v.Externo) { $externos += New-Object PSObject -Property @{ Guid = $v.Guid; Caminho = $caminho; Dn = $v.Dn } }
            else {
                if (-not $orfaos.ContainsKey($v.Guid)) {
                    $orfaos[$v.Guid] = New-Object PSObject -Property @{ Guid = $v.Guid; Nome = '(inexistente)'; Vinculos = @(); Achados = @('Orfao') }
                }
                $orfaos[$v.Guid].Vinculos += $lig
            }
        }
    }

    $achados = @{}
    $lista = @(@($porGuid.Values) | Sort-Object Nome, Guid)
    foreach ($g in @($lista)) {
        $ids = @()
        if (@($g.SenhaEm).Count) { $ids += 'SenhaPreferencia' }
        if ($sysvolOk -and $g.PastaExiste -eq $false) { $ids += 'SemSysvol' }
        if (-not @($g.Vinculos).Count) { $ids += 'SemVinculo' }
        if ($g.Flags -eq 3) { $ids += 'Desativada' }
        if (@(@($g.Vinculos) | Where-Object { $_.Desativado }).Count) { $ids += 'VinculoDesativado' }
        if ($g.Versao -eq 0) { $ids += 'Vazia' }
        $g.Achados = $ids
        foreach ($id in @($ids)) { Add-AchadoGpo $achados $id $g }
    }
    $listaOrfaos = @(@($orfaos.Values) | Sort-Object Guid)
    foreach ($o in @($listaOrfaos)) { Add-AchadoGpo $achados 'Orfao' $o }

    New-Object PSObject -Property @{ Ok = $true; Erro = $null; Avisos = @(); Gpos = $lista; Orfaos = $listaOrfaos
        Bloqueios = @($bloqueios | Sort-Object -Unique); Externos = $externos; Achados = $achados; SysvolOk = [bool]$sysvolOk; Csv = $null; CsvConfig = $null }
}

function TextoVinculos($g, [switch]$Html) {
    $itens = @()
    foreach ($v in @($g.Vinculos)) {
        if (-not $v) { continue }
        if ($Html) {
            $t = HE $v.Caminho
            if ($v.Desativado) { $t += '<span class="tag t-warn">desativado</span>' }
            if ($v.Imposto) { $t += '<span class="tag">imposto</span>' }
        } else {
            $t = $v.Caminho
            if ($v.Desativado) { $t += ' (desativado)' }
            if ($v.Imposto) { $t += ' (imposto)' }
        }
        $itens += $t
    }
    if (-not $itens.Count) { if ($Html) { return '—' } else { return '' } }
    if ($Html) { $itens -join '<br>' } else { $itens -join ', ' }
}

function NomesAchadosGpo($g) {
    (@($GpoAchadosDef) | Where-Object { @($g.Achados) -contains $_.Id } | ForEach-Object { $_.Nome }) -join ', '
}

# Resumo das áreas: "Computador: Scripts (2), Opções de segurança (5) | Usuário: Mapeamento de unidades (3)"
function TextoAreasGpo($g) {
    $lados = @()
    foreach ($lado in @('Computador', 'Usuário')) {
        $a = @(@($g.Itens) | Where-Object { $_ -and $_.Lado -eq $lado } | ForEach-Object { $n = @($_.Detalhes).Count; if ($n) { "$($_.Area) ($n)" } else { $_.Area } })
        if ($a.Count) { $lados += "${lado}: $($a -join ', ')" }
    }
    $lados -join ' | '
}

# CSV 06_GPOs (GPOs e vínculos órfãos); se falhar, só avisa na seção
function Write-CsvGpo($r, [string]$arquivo) {
    try {
        @(@($r.Gpos) + @($r.Orfaos)) | Where-Object { $_ } | ForEach-Object {
            $g = $_
            $orfao = (@($g.Achados) -contains 'Orfao')
            $pasta = if ($orfao) { '' } elseif ($null -eq $g.PastaExiste) { 'não verificada' } elseif ($g.PastaExiste) { 'ok' } else { 'ausente' }
            $desat = @(@($g.Vinculos) | Where-Object { $_.Desativado } | ForEach-Object { $_.Caminho }) -join ', '
            New-Object PSObject -Property @{ GPO = $g.Nome; GUID = $g.Guid; Status = $g.Status; Versao = $g.Versao; CriadaEm = $g.CriadaEm
                AlteradaEm = $g.AlteradaEm; VinculadaEm = (TextoVinculos $g); VinculosDesativados = $desat; PastaSYSVOL = $pasta
                Configuracoes = (TextoAreasGpo $g); Achados = (NomesAchadosGpo $g); ErrosLeitura = (@($g.ErrosLeitura) -join ' / ') }
        } | Select-Object GPO, GUID, Status, Versao, CriadaEm, AlteradaEm, VinculadaEm, VinculosDesativados, PastaSYSVOL, Configuracoes, Achados, ErrosLeitura |
            Export-Csv $arquivo @csvOpts
        $r.Csv = $arquivo
    } catch { $r.Avisos += "não foi possível gravar o arquivo 06_GPOs: $($_.Exception.Message)" }
}

# CSV 06_GPOs_Configuracoes: uma linha por configuração lida (sem o limite de linhas do HTML); se falhar, só avisa na seção
function Write-CsvGpoConfig($r, [string]$arquivo) {
    try {
        $linhas = New-Object System.Collections.ArrayList
        foreach ($g in @($r.Gpos)) {
            if (-not $g) { continue }
            foreach ($i in @($g.Itens)) {
                if (-not $i) { continue }
                $det = @(@($i.Detalhes) | Where-Object { $_ })
                if (-not $det.Count) { $det = @($(if ($g.ConteudoLido) { '(área configurada, sem detalhe neste relatório)' } else { '(área configurada; SYSVOL não lido)' })) }
                foreach ($d in @($det)) {
                    [void]$linhas.Add((New-Object PSObject -Property @{ GPO = $g.Nome; GUID = $g.Guid; Parte = $i.Lado; Area = $i.Area; Configuracao = $d }))
                }
            }
        }
        @($linhas) | Select-Object GPO, GUID, Parte, Area, Configuracao | Export-Csv $arquivo @csvOpts
        $r.CsvConfig = $arquivo
    } catch { $r.Avisos += "não foi possível gravar o arquivo 06_GPOs_Configuracoes: $($_.Exception.Message)" }
}

function CardsGpo($r) {
    $cards = @(
        @{ Grupo = 'seguranca'; Sev = 'alta'; Um = 'GPO com senha exposta no SYSVOL'; Varios = 'GPOs com senha exposta no SYSVOL'
           Texto = 'As Preferências de GPO guardavam senhas (campo cpassword) cifradas com uma chave que a própria Microsoft publicou (boletim MS14-025): qualquer usuário do domínio lê o SYSVOL e recupera a senha em segundos. Trocar já a senha da conta afetada, remover o item (ou a senha) da GPO e usar o LAPS para senhas de administrador local.' }
        @{ Grupo = 'integridade'; Sev = 'media'; Um = 'GPO com problema de integridade'; Varios = 'GPOs com problema de integridade'
           Texto = 'Vínculos que apontam para GPOs apagadas ou GPOs sem a pasta no SYSVOL geram erros de aplicação de política nas máquinas (eventos 1058/1096). Remover o vínculo órfão e recriar ou excluir a GPO sem pasta, depois de confirmar a replicação do SYSVOL.' }
        @{ Grupo = 'limpeza'; Sev = 'baixa'; Um = 'GPO para limpeza'; Varios = 'GPOs para limpeza'
           Texto = 'GPOs sem vínculo, desativadas, com vínculo desativado ou vazias não fazem nada, mas confundem quem administra o domínio e atrasam a análise de problemas. Confirmar com o cliente e excluir as que não forem mais usadas (fazer backup antes).' }
    )
    foreach ($c in @($cards)) {
        $tr = ''; $n = 0
        foreach ($d in @($GpoAchadosDef)) {
            if ($d.Grupo -ne $c.Grupo -or -not $r.Achados.ContainsKey($d.Id)) { continue }
            foreach ($g in @($r.Achados[$d.Id])) {
                $n++
                $det = if ($d.Id -eq 'SenhaPreferencia') { @($g.SenhaEm) -join ', ' } elseif ($d.Id -eq 'SemSysvol') { $g.Pasta } elseif ($d.Id -eq 'Desativada' -or $d.Id -eq 'Vazia') { "alterada em $(DT $g.AlteradaEm)" } else { TextoVinculos $g }
                $nome = if ($d.Id -eq 'Orfao') { $g.Guid } else { $g.Nome }
                $tr += "<tr><td>$($d.Nome)</td><td>$(HE $nome)</td><td>$(HE $det)</td></tr>"
            }
        }
        if (-not $n) { continue }
        Card $c.Sev "$n $(Plural $n $c.Um $c.Varios)" ("<p>$($c.Texto) Detalhes na seção <a href=""#mod-gpos"">GPOs</a>.</p><div class=""scroll""><table class=""mini""><thead><tr><th>Problema</th><th>GPO</th><th>Detalhe</th></tr></thead><tbody>$tr</tbody></table></div>")
    }
}

function TagsAchadosGpo($g) {
    $tags = ''
    foreach ($d in @($GpoAchadosDef)) {
        if (@($g.Achados) -contains $d.Id) { $cls = if ($d.Grupo -eq 'limpeza') { 't-warn' } else { 't-bad' }; $tags += "<span class=""tag $cls"">$($d.Tag)</span>" }
    }
    $tags
}

# Lista de configurações de um item (até $GpoLimiteDetalhes linhas)
function DetalhesItemGpo($i, [bool]$lido) {
    $det = @(@($i.Detalhes) | Where-Object { $_ })
    if (-not $det.Count) {
        if ($lido) { return '<i>área configurada; o conteúdo dela não é detalhado neste relatório</i>' }
        return '<i>área configurada; detalhes não lidos (SYSVOL inacessível)</i>'
    }
    $h = '<ul class="gpo-det">'
    $n = 0
    foreach ($d in @($det)) {
        if ($n -ge $GpoLimiteDetalhes) { break }
        $h += "<li>$(HE $d)</li>"
        $n++
    }
    if ($det.Count -gt $n) { $h += "<li>+$($det.Count - $n) no arquivo 06_GPOs_Configuracoes</li>" }
    $h + '</ul>'
}

# "O que cada GPO faz": um bloco por GPO com as configurações por parte (Computador/Usuário)
function ConteudoGpoHtml($g) {
    $h = "<h4 class=""gpo"">$(HE $g.Nome)$(TagsAchadosGpo $g)</h4>"
    $h += "<p class=""gpo-meta"">Status: $(HE $g.Status) · alterada em $(DT $g.AlteradaEm) · aplicada em: $(TextoVinculos $g -Html)</p>"
    $itens = @(@($g.Itens) | Where-Object { $_ })
    if (-not $itens.Count) {
        $msg = if ($g.ConteudoLido -or $g.Versao -eq 0) { 'Nenhuma configuração encontrada: a GPO não faz nada.' } else { 'Conteúdo não lido (pasta no SYSVOL inacessível ou ausente).' }
        return $h + "<p class=""gpo-meta""><i>$msg</i></p>"
    }
    # Flags: 1 = parte Usuário desativada, 2 = parte Computador desativada, 3 = toda desativada
    $desligado = @{ 'Usuário' = [bool]($g.Flags -band 1); 'Computador' = [bool]($g.Flags -band 2) }
    $h += '<table class="mini"><thead><tr><th>Parte</th><th>Área</th><th>O que configura</th></tr></thead><tbody>'
    foreach ($i in @($itens)) {
        $tag = if ($desligado[$i.Lado]) { '<span class="tag t-warn">parte desativada</span>' } else { '' }
        $h += "<tr><td>$($i.Lado)$tag</td><td>$(HE $i.Area)</td><td>$(DetalhesItemGpo $i $g.ConteudoLido)</td></tr>"
    }
    $h += '</tbody></table>'
    if (@($itens | Where-Object { $_.Area -like 'Política de contas*' }).Count -and -not (Test-VinculoRaiz $g)) {
        $h += '<p class="gpo-meta">Atenção: a política de senha e bloqueio só vale para as contas do domínio quando a GPO está vinculada na raiz do domínio. Vinculada só em OUs, como aqui, ela muda apenas as contas locais dos computadores.</p>'
    }
    if (@($g.ErrosLeitura).Count) { $h += "<p class=""gpo-meta"">Não foi possível ler: $(HE (@($g.ErrosLeitura) -join ' / ')).</p>" }
    $h
}

function SecaoGpo($r, $num) {
    $h = "<section id=""mod-gpos""><h2><span>$num</span>GPOs</h2>"
    $h += '<p class="lead">Uma GPO (objeto de Política de Grupo) é um pacote de configurações que o domínio aplica sozinho nos computadores e usuários: senhas, mapeamento de unidades, impressoras, scripts, restrições, atualizações etc. Ela só tem efeito onde está <b>vinculada</b> (domínio, OU ou site) e vale para os objetos daquele local e das OUs abaixo dele. A parte <b>Computador</b> é aplicada às máquinas na inicialização e a cada ~90 minutos; a parte <b>Usuário</b>, a quem faz logon nessas OUs, no logon e a cada ~90 minutos. Ordem de aplicação: site → domínio → OUs, da mais alta para a mais baixa; em conflito vence a mais próxima do objeto, exceto os vínculos <b>impostos</b>, que vencem sempre e passam por cima do bloqueio de herança.</p>'
    $h += '<p class="lead"><b>Como ler esta seção:</b> as GPOs e os vínculos foram lidos direto do AD e o conteúdo, dos arquivos de cada GPO no SYSVOL, sem o console de GPO e sem alterar nada. Aparecem as principais configurações de cada área; nos Modelos administrativos aparece o nome técnico da chave do registro, agrupado por assunto. Não são analisados os filtros de segurança e WMI (que podem restringir a quem cada GPO se aplica) nem o direcionamento por item das Preferências. Para o relatório completo de uma GPO, use o Gerenciamento de Política de Grupo (botão direito na GPO → Salvar relatório) ou <code>Get-GPOReport -Name &lt;GPO&gt; -ReportType Html</code>. Lista completa nos arquivos 06_GPOs e 06_GPOs_Configuracoes.</p>'
    foreach ($a in @($r.Avisos)) { if ($a) { $h += "<p class=""lead"">Observação: $(HE $a).</p>" } }
    $lista = @($r.Gpos)
    if (-not $lista.Count) { $h += '<p class="lead">Nenhuma GPO encontrada.</p>' }
    else {
        $h += '<div class="scroll"><table class="mini"><thead><tr><th>GPO</th><th>Status</th><th class="num">Versão</th><th class="num">Alterada em</th><th>Vinculada em</th></tr></thead><tbody>'
        foreach ($g in @($lista)) {
            $h += "<tr><td><a href=""#gpo-$($g.Guid.Trim('{}'))"">$(HE $g.Nome)</a>$(TagsAchadosGpo $g)</td><td>$(HE $g.Status)</td><td class=""num"">$($g.Versao)</td><td class=""num"">$(DT $g.AlteradaEm)</td><td>$(TextoVinculos $g -Html)</td></tr>"
        }
        $h += '</tbody></table></div>'
        $h += '<h3 class="mod-sub">O que cada GPO faz</h3>'
        foreach ($g in @($lista)) { $h += "<div class=""gpo-bloco"" id=""gpo-$($g.Guid.Trim('{}'))"">$(ConteudoGpoHtml $g)</div>" }
    }
    $orf = @($r.Orfaos)
    if ($orf.Count) {
        $h += '<h3 class="mod-sub">Vínculos para GPOs inexistentes</h3><table class="mini"><thead><tr><th>GUID</th><th>Vinculado em</th></tr></thead><tbody>'
        foreach ($o in @($orf)) { $h += "<tr><td>$(HE $o.Guid)</td><td>$(TextoVinculos $o -Html)</td></tr>" }
        $h += '</tbody></table>'
    }
    $bl = @($r.Bloqueios)
    $h += '<h3 class="mod-sub">Herança bloqueada</h3>'
    if ($bl.Count) {
        $h += '<p class="lead">Nestes locais as GPOs de cima (domínio e OUs-pai) não se aplicam, exceto as impostas:</p><table class="mini"><tbody>'
        foreach ($b in @($bl)) { $h += "<tr><td>$(HE $b)</td></tr>" }
        $h += '</tbody></table>'
    } else { $h += '<p class="lead">Nenhuma OU com herança bloqueada.</p>' }
    $ext = @($r.Externos)
    if ($ext.Count) { $h += "<p class=""lead"">Observação: $($ext.Count) $(Plural $ext.Count 'vínculo aponta' 'vínculos apontam') para GPOs de outro domínio da floresta; não foram verificados.</p>" }
    $h + '</section>'
}

# ---------------------------------------------------------------------
# Computadores do domínio (módulo 5)
# ---------------------------------------------------------------------
# Fim do suporte = último dia com suporte (o Microsoft Lifecycle mostra o dia seguinte às 06:59:59 PT).
# Datas conferidas no Microsoft Lifecycle e no Windows release health em 2026-10-28.
$TabelaSuporte = @{
    # Testados em ordem (8.1 antes de 8)
    Clientes = @(
        @{ Padrao = 'Windows 2000';  Familia = 'Windows 2000';  Fim = '2010-07-13' }
        @{ Padrao = 'Windows XP';    Familia = 'Windows XP';    Fim = '2014-04-08' }
        @{ Padrao = 'Windows Vista'; Familia = 'Windows Vista'; Fim = '2017-04-11' }
        @{ Padrao = 'Windows 7\b';   Familia = 'Windows 7';     Fim = '2020-01-14' }
        @{ Padrao = 'Windows 8\.1';  Familia = 'Windows 8.1';   Fim = '2023-01-10' }
        @{ Padrao = 'Windows 8\b';   Familia = 'Windows 8';     Fim = '2016-01-12' }
    )
    Win10     = @{ Familia = 'Windows 10'; Fim = '2025-10-14' }
    Win10Ltsc = @{
        '10240' = @{ Familia = 'Windows 10 LTSB 2015'; Fim = '2025-10-14' }
        '14393' = @{ Familia = 'Windows 10 LTSB 2016'; Fim = '2026-10-13' }
        '17763' = @{ Familia = 'Windows 10 LTSC 2019'; Fim = '2029-01-09' }
        '19044' = @{ Familia = 'Windows 10 LTSC 2021'; Fim = '2027-01-12' }
    }
    Win11Ltsc = @{
        '26100' = @{ Familia = 'Windows 11 LTSC 2024'; Fim = '2029-10-09' }
    }
    # Home/Pro (Enterprise/Education: 12 meses a mais)
    Win11 = @{
        '22000' = @{ Familia = 'Windows 11 21H2'; Fim = '2023-10-10' }
        '22621' = @{ Familia = 'Windows 11 22H2'; Fim = '2024-10-08' }
        '22631' = @{ Familia = 'Windows 11 23H2'; Fim = '2025-11-11' }
        '26100' = @{ Familia = 'Windows 11 24H2'; Fim = '2026-10-13' }
        '26200' = @{ Familia = 'Windows 11 25H2'; Fim = '2027-10-12' }
        '26300' = @{ Familia = 'Windows 11 26H2'; Fim = '2028-10-09' }
        '28000' = @{ Familia = 'Windows 11 26H1'; Fim = '2028-03-14' }
    }
    ServidorAno = @(
        @{ Ano = '2003'; Fim = '2015-07-14' }
        @{ Ano = '2008'; Fim = '2020-01-14' }
        @{ Ano = '2012'; Fim = '2023-10-10' }
        @{ Ano = '2016'; Fim = '2027-01-12' }
        @{ Ano = '2019'; Fim = '2029-01-09' }
        @{ Ano = '2022'; Fim = '2031-10-14' }
        @{ Ano = '2025'; Fim = '2034-11-14' }
    )
    # "Windows Server Standard/Datacenter" sem ano: build → ano do LTSC
    ServidorBuild = @{ '14393' = '2016'; '17763' = '2019'; '20348' = '2022'; '26100' = '2025' }
    # Canal anual (23H2) e semestral (1709 a 20H2, retirado em 2022-08-09)
    Servidor23H2 = @{ Build = '25398'; Familia = 'Windows Server 23H2 (canal anual)'; Fim = '2026-05-12' }
    ServidorSac  = @{ Familia = 'Windows Server (canal semestral)'; Fim = '2022-08-09' }
}
$Win11Obs = 'datas de Home/Pro; as edições Enterprise/Education têm 12 meses a mais'
$CompLimiteLista = 300
# Sem logon há mais deste número de meses = máquina que provavelmente não existe mais (candidata à remoção)
$CompMesesAntigo = 7

function NovoFim([string]$familia, [string]$fim, [string]$obs) {
    $d = if ($fim) { [datetime]::ParseExact($fim, 'yyyy-MM-dd', [Globalization.CultureInfo]::InvariantCulture) } else { $null }
    New-Object PSObject -Property @{ Familia = $familia; Fim = $d; Obs = $obs }
}

function Get-MaiorBuild($tabela) {
    $max = 0
    foreach ($k in @($tabela.Keys)) { if ([int]$k -gt $max) { $max = [int]$k } }
    $max
}

# Fim do suporte (pura) a partir de OperatingSystem e OperatingSystemVersion ('10.0 (19045)') → {Familia; Fim; Obs}
function Get-FimSuporte([string]$so, [string]$versao) {
    $t = $TabelaSuporte
    $so = $so.Trim()
    $build = if ($versao -match '\((\d+)\)') { $matches[1] } else { '' }
    if (-not $so -or $so -notmatch 'Windows') { return NovoFim 'Não Windows / desconhecido' '' '' }
    if ($so -match 'Windows 2000') { return NovoFim 'Windows 2000' '2010-07-13' '' }

    if ($so -match 'Server') {
        $ano = $null
        foreach ($s in @($t.ServidorAno)) { if ($so -match "\b$($s.Ano)\b") { $ano = $s; break } }
        if (-not $ano -and $build -and $t.ServidorBuild.ContainsKey($build)) {
            foreach ($s in @($t.ServidorAno)) { if ($s.Ano -eq $t.ServidorBuild[$build]) { $ano = $s; break } }
        }
        if ($ano) {
            $f = "Windows Server $($ano.Ano)"
            if ($so -match '\bR2\b') { $f += ' R2' }
            return NovoFim $f $ano.Fim ''
        }
        if ($build -eq $t.Servidor23H2.Build) { return NovoFim $t.Servidor23H2.Familia $t.Servidor23H2.Fim '' }
        if ($build -and [int]$build -gt (Get-MaiorBuild $t.ServidorBuild)) { return NovoFim 'Windows Server (mais nova que a tabela)' '' 'versão mais nova que a tabela' }
        return NovoFim $t.ServidorSac.Familia $t.ServidorSac.Fim ''
    }

    if ($so -match 'Windows 10\b') {
        if ($so -match 'LTS[BC]' -and $t.Win10Ltsc.ContainsKey($build)) { $x = $t.Win10Ltsc[$build]; return NovoFim $x.Familia $x.Fim '' }
        return NovoFim $t.Win10.Familia $t.Win10.Fim ''
    }
    if ($so -match 'Windows 11\b') {
        if ($so -match 'LTSC' -and $t.Win11Ltsc.ContainsKey($build)) { $x = $t.Win11Ltsc[$build]; return NovoFim $x.Familia $x.Fim '' }
        if ($build -and $t.Win11.ContainsKey($build)) { $x = $t.Win11[$build]; return NovoFim $x.Familia $x.Fim $Win11Obs }
        if ($build -and [int]$build -gt (Get-MaiorBuild $t.Win11)) { return NovoFim 'Windows 11 (mais nova que a tabela)' '' 'versão mais nova que a tabela' }
        return NovoFim "Windows 11 (build $(if ($build) { $build } else { '?' }))" '' 'versão fora da tabela'
    }
    foreach ($c in @($t.Clientes)) { if ($so -match $c.Padrao) { return NovoFim $c.Familia $c.Fim '' } }
    NovoFim $so '' 'sistema fora da tabela'
}

function Read-ModComputadores {
    @(Get-ADComputer -Filter * -Properties OperatingSystem, OperatingSystemVersion, LastLogonDate, whenCreated)
}

# Análise (pura). Devolve o resultado do módulo: Computadores, Familias (resumo), ForaSuporte, Antigos, Inativos e Desativados.
#   Antigo  = sem logon (ou, se nunca logou, criado) há mais de $CompMesesAntigo meses, ativo ou desativado
#   Inativo = ativo, sem uso há mais de $DiasInativo dias e não antigo
#   ForaSuporte = ativo, não antigo e com o fim do suporte já passado
function Get-AchadosComputadores($comps, $hoje, $DiasInativo) {
    $limite = $hoje.AddDays(-$DiasInativo)
    $limiteAntigo = $hoje.AddMonths(-$CompMesesAntigo)
    $cache = @{}; $fams = @{}
    $lista = New-Object System.Collections.ArrayList
    $desativados = 0
    foreach ($pc in @($comps)) {
        if (-not $pc) { continue }
        $so = [string]$pc.OperatingSystem; $ver = [string]$pc.OperatingSystemVersion
        $k = "$so|$ver"
        if (-not $cache.ContainsKey($k)) { $cache[$k] = Get-FimSuporte $so $ver }
        $f = $cache[$k]
        $ativo = ($pc.Enabled -ne $false)
        $logon = $pc.LastLogonDate; $criado = $pc.whenCreated
        $ref = if ($logon) { $logon } else { $criado }
        $antigo = [bool]($ref -and $ref -lt $limiteAntigo)
        $inativo = [bool]($ativo -and -not $antigo -and $ref -and $ref -lt $limite)
        $fora = [bool]($ativo -and -not $antigo -and $f.Fim -and $f.Fim -lt $hoje.Date)
        $dias = if ($logon) { [int][math]::Floor(($hoje - $logon).TotalDays) } else { $null }
        [void]$lista.Add((New-Object PSObject -Property @{ Nome = [string]$pc.Name; SO = $so; Versao = $ver; Ativo = $ativo; UltimoLogon = $logon
            CriadoEm = $criado; Familia = $f.Familia; Fim = $f.Fim; Obs = $f.Obs; ForaSuporte = $fora; Inativo = $inativo; Antigo = $antigo
            DiasSemLogon = $dias; OU = (Get-OUConta ([string]$pc.DistinguishedName)) }))
        if (-not $ativo) { $desativados++ }
        if (-not $fams.ContainsKey($f.Familia)) {
            $fams[$f.Familia] = New-Object PSObject -Property @{ Familia = $f.Familia; Fim = $f.Fim; Obs = $f.Obs; Total = 0; Ativos = 0; Inativos = 0; Antigos = 0; ForaSuporte = 0 }
        }
        $s = $fams[$f.Familia]
        $s.Total = $s.Total + 1
        if ($ativo) { $s.Ativos = $s.Ativos + 1 }
        if ($inativo) { $s.Inativos = $s.Inativos + 1 }
        if ($antigo) { $s.Antigos = $s.Antigos + 1 }
        if ($fora) { $s.ForaSuporte = $s.ForaSuporte + 1 }
    }
    $todos = @($lista | Sort-Object Nome)
    # Famílias pela data de fim (sem data no fim)
    $familias = @(@($fams.Values) | Sort-Object @{ Expression = { if ($_.Fim) { $_.Fim } else { [datetime]::MaxValue } } }, Familia)
    New-Object PSObject -Property @{ Ok = $true; Erro = $null; Avisos = @(); Computadores = $todos; Familias = $familias
        ForaSuporte = @(@($todos) | Where-Object { $_.ForaSuporte } | Sort-Object Fim, Nome)
        Antigos = @(@($todos) | Where-Object { $_.Antigo } | Sort-Object @{ Expression = { if ($_.UltimoLogon) { $_.UltimoLogon } else { $_.CriadoEm } } }, Nome)
        Inativos = @(@($todos) | Where-Object { $_.Inativo } | Sort-Object UltimoLogon, Nome)
        Desativados = $desativados; DiasInativo = $DiasInativo; MesesAntigo = $CompMesesAntigo; Csv = $null }
}

# CSV 07_Computadores (todos os computadores); se falhar, só avisa na seção
function Write-CsvComputadores($r, [string]$arquivo) {
    try {
        @($r.Computadores) | Where-Object { $_ } | ForEach-Object {
            $c = $_
            New-Object PSObject -Property @{ Nome = $c.Nome; SO = $c.SO; Versao = $c.Versao; Ativo = $c.Ativo; UltimoLogon = $c.UltimoLogon
                CriadoEm = $c.CriadoEm; Familia = $c.Familia; FimSuporte = $c.Fim; ForaDeSuporte = $c.ForaSuporte; Inativo = $c.Inativo; OU = $c.OU
                DiasSemLogon = $c.DiasSemLogon; CandidatoRemocao = $c.Antigo }
        } | Select-Object Nome, SO, Versao, Ativo, UltimoLogon, DiasSemLogon, CriadoEm, Familia, FimSuporte, ForaDeSuporte, Inativo, CandidatoRemocao, OU |
            Export-Csv $arquivo @csvOpts
        $r.Csv = $arquivo
    } catch { $r.Avisos += "não foi possível gravar o arquivo 07_Computadores: $($_.Exception.Message)" }
}

# Fim do suporte na tabela: data (com * nas datas de Home/Pro) ou a observação
function TextoFimSuporte($f) {
    if ($f.Fim) { return "$(DT $f.Fim)$(if ($f.Obs -eq $Win11Obs) { '*' })" }
    if ($f.Obs) { $f.Obs } else { '—' }
}

function CardsComputadores($r) {
    $n = @($r.ForaSuporte).Count
    if ($n) {
        $tr = ''
        foreach ($f in @($r.Familias)) {
            if ($f.ForaSuporte) { $tr += "<tr><td>$(HE $f.Familia)</td><td class=""num"">$($f.ForaSuporte)</td><td class=""num"">$(TextoFimSuporte $f)</td></tr>" }
        }
        Card 'alta' "$n $(Plural $n 'computador' 'computadores') com sistema fora de suporte" ("<p>Sistemas sem atualização de segurança da Microsoft ficam expostos a falhas já conhecidas e exploradas. Atualizar ou substituir; enquanto isso, isolar da rede o que não puder sair. Só entram computadores ativos que fizeram logon nos últimos $($r.MesesAntigo) meses. Lista na seção <a href=""#mod-computadores"">Computadores</a> e no arquivo 07_Computadores.</p><table class=""mini""><thead><tr><th>Sistema</th><th class=""num"">Computadores</th><th class=""num"">Fim do suporte</th></tr></thead><tbody>$tr</tbody></table>")
    }
    $ant = @($r.Antigos)
    $n = $ant.Count
    if ($n) {
        $nAtivos = @($ant | Where-Object { $_.Ativo }).Count
        Card 'media' "$n $(Plural $n 'computador' 'computadores') sem logon há mais de $($r.MesesAntigo) meses (candidatos à remoção)" "<p>Contas de computador que não fazem logon no domínio há mais de $($r.MesesAntigo) meses ($nAtivos $(Plural $nAtivos 'ainda ativa' 'ainda ativas'), $($n - $nAtivos) já $(Plural ($n - $nAtivos) 'desativada' 'desativadas')): quase sempre máquinas descartadas ou reinstaladas com outro nome, que não existem mais na rede. Confirmar com o cliente, desativar e mover para uma OU de desativados; excluir depois de algumas semanas sem reclamação. Lista na seção <a href=""#mod-computadores"">Computadores</a> e coluna CandidatoRemocao do arquivo 07_Computadores.</p>"
    }
    $n = @($r.Inativos).Count
    if ($n) {
        Card 'baixa' "$n $(Plural $n 'computador ativo' 'computadores ativos') sem uso há mais de $($r.DiasInativo) dias (e menos de $($r.MesesAntigo) meses)" "<p>Contas de computador habilitadas sem logon no domínio há mais de $($r.DiasInativo) dias, mas que ainda fizeram logon nos últimos $($r.MesesAntigo) meses: notebooks fora da rede, máquinas desligadas ou em manutenção. Verificar com o cliente antes de desativar. Lista na seção <a href=""#mod-computadores"">Computadores</a>.</p>"
    }
    $n = [int]$r.Desativados
    if ($n) {
        $nAnt = @(@($r.Antigos) | Where-Object { -not $_.Ativo }).Count
        $obs = if ($nAnt) { " $nAnt $(Plural $nAnt 'delas está' 'delas estão') na lista de candidatos à remoção (sem logon há mais de $($r.MesesAntigo) meses)." } else { '' }
        Card 'baixa' "$n $(Plural $n 'conta de computador desativada' 'contas de computador desativadas')" "<p>Contas de computador já desativadas continuam no AD. Se as máquinas não voltarem a ser usadas, podem ser excluídas (de preferência movidas antes para uma OU de desativados).$obs</p>"
    }
}

# Linhas da lista (até $max), com "+N no arquivo 07_Computadores" no fim
function LinhasLimitadas($lista, [scriptblock]$linha, [int]$cols, [int]$max) {
    $h = ''; $i = 0
    foreach ($c in @($lista)) {
        if ($i -ge $max) { break }
        $h += & $linha $c
        $i++
    }
    $resto = @($lista).Count - $i
    if ($resto -gt 0) { $h += "<tr><td colspan=""$cols"">+$resto no arquivo 07_Computadores</td></tr>" }
    $h
}

function SecaoComputadores($r, $num) {
    $h = "<section id=""mod-computadores""><h2><span>$num</span>Computadores</h2>"
    $h += "<p class=""lead"">Contas de computador do domínio por sistema operacional, com a data de fim do suporte da Microsoft. Os computadores sem logon há mais de $($r.MesesAntigo) meses (ativos ou desativados) ficam separados como candidatos à remoção; fora de suporte e sem uso só contam os ativos que fizeram logon nesse período. Lista completa no arquivo 07_Computadores.</p>"
    foreach ($a in @($r.Avisos)) { if ($a) { $h += "<p class=""lead"">Observação: $(HE $a).</p>" } }
    $todos = @($r.Computadores)
    if (-not $todos.Count) { return $h + '<p class="lead">Nenhum computador encontrado.</p></section>' }
    $h += "<table class=""mini""><thead><tr><th>Sistema</th><th class=""num"">Total</th><th class=""num"">Ativos</th><th class=""num"">Sem uso ($($r.DiasInativo) dias a $($r.MesesAntigo) meses)</th><th class=""num"">Sem logon há mais de $($r.MesesAntigo) meses</th><th class=""num"">Fora de suporte</th><th class=""num"">Fim do suporte</th></tr></thead><tbody>"
    $nota = $false
    foreach ($f in @($r.Familias)) {
        if ($f.Obs -eq $Win11Obs) { $nota = $true }
        $cls = if ($f.ForaSuporte) { ' class="num sev-alta"' } else { ' class="num"' }
        $clsAnt = if ($f.Antigos) { ' class="num sev-media"' } else { ' class="num"' }
        $h += "<tr><td>$(HE $f.Familia)</td><td class=""num"">$($f.Total)</td><td class=""num"">$($f.Ativos)</td><td class=""num"">$($f.Inativos)</td><td$clsAnt>$($f.Antigos)</td><td$cls>$($f.ForaSuporte)</td><td class=""num"">$(TextoFimSuporte $f)</td></tr>"
    }
    $h += "<tr><td><b>Total</b></td><td class=""num""><b>$($todos.Count)</b></td><td class=""num""><b>$($todos.Count - [int]$r.Desativados)</b></td><td class=""num""><b>$(@($r.Inativos).Count)</b></td><td class=""num""><b>$(@($r.Antigos).Count)</b></td><td class=""num""><b>$(@($r.ForaSuporte).Count)</b></td><td></td></tr>"
    $h += '</tbody></table>'
    if ($nota) { $h += "<p class=""lead"">* Windows 11: $Win11Obs.</p>" }

    $fora = @($r.ForaSuporte)
    $h += '<h3 class="mod-sub">Fora de suporte</h3>'
    if ($fora.Count) {
        $h += '<div class="scroll"><table class="mini"><thead><tr><th>Computador</th><th>Sistema</th><th class="num">Fim do suporte</th><th class="num">Último logon</th><th>OU</th></tr></thead><tbody>'
        $h += LinhasLimitadas $fora { param($c) "<tr><td>$(HE $c.Nome)</td><td>$(HE $c.SO)</td><td class=""num"">$(DT $c.Fim)</td><td class=""num"">$(DT $c.UltimoLogon)</td><td>$(HE (Get-CaminhoDN $c.OU ''))</td></tr>" } 5 $CompLimiteLista
        $h += '</tbody></table></div>'
    } else { $h += '<p class="lead">Nenhum computador ativo com sistema fora de suporte.</p>' }

    $ant = @($r.Antigos)
    $h += "<h3 class=""mod-sub"">Sem logon há mais de $($r.MesesAntigo) meses (candidatos à remoção)</h3>"
    if ($ant.Count) {
        $h += "<p class=""lead"">Provavelmente não existem mais na rede. Do mais antigo para o mais recente; confirmar com o cliente, desativar e excluir depois de algumas semanas.</p>"
        $h += '<div class="scroll"><table class="mini"><thead><tr><th>Computador</th><th>Situação</th><th>Sistema</th><th class="num">Último logon</th><th class="num">Dias sem logon</th><th class="num">Criado em</th><th>OU</th></tr></thead><tbody>'
        $h += LinhasLimitadas $ant { param($c) "<tr><td>$(HE $c.Nome)</td><td>$(if ($c.Ativo) { 'ativo' } else { 'desativado' })</td><td>$(HE $c.SO)</td><td class=""num"">$(if ($c.UltimoLogon) { DT $c.UltimoLogon } else { 'nunca' })</td><td class=""num"">$($c.DiasSemLogon)</td><td class=""num"">$(DT $c.CriadoEm)</td><td>$(HE (Get-CaminhoDN $c.OU ''))</td></tr>" } 7 $CompLimiteLista
        $h += '</tbody></table></div>'
    } else { $h += "<p class=""lead"">Nenhum computador sem logon há mais de $($r.MesesAntigo) meses.</p>" }

    $ina = @($r.Inativos)
    $h += "<h3 class=""mod-sub"">Sem uso há mais de $($r.DiasInativo) dias (e menos de $($r.MesesAntigo) meses)</h3>"
    if ($ina.Count) {
        $h += '<p class="lead">Ainda fizeram logon no período recente: verificar com o cliente antes de desativar.</p>'
        $h += '<div class="scroll"><table class="mini"><thead><tr><th>Computador</th><th>Sistema</th><th class="num">Último logon</th><th class="num">Dias sem logon</th><th class="num">Criado em</th><th>OU</th></tr></thead><tbody>'
        $h += LinhasLimitadas $ina { param($c) "<tr><td>$(HE $c.Nome)</td><td>$(HE $c.SO)</td><td class=""num"">$(if ($c.UltimoLogon) { DT $c.UltimoLogon } else { 'nunca' })</td><td class=""num"">$($c.DiasSemLogon)</td><td class=""num"">$(DT $c.CriadoEm)</td><td>$(HE (Get-CaminhoDN $c.OU ''))</td></tr>" } 6 $CompLimiteLista
        $h += '</tbody></table></div>'
    } else { $h += '<p class="lead">Nenhum computador ativo nessa faixa.</p>' }
    $h + '</section>'
}

# Número da seção do módulo: 6 + posição em Marcados
function NumeroSecao($MOD, [string]$id) {
    $i = 0
    foreach ($m in @($MOD.Marcados)) { if ($m -eq $id) { return 6 + $i }; $i++ }
    $null
}

# ---------------------------------------------------------------------
# Ganchos de HTML: com nenhum módulo marcado, devolvem '' (ou nenhum card)
# ---------------------------------------------------------------------
function NavModulos($MOD) {
    $h = ''
    foreach ($id in @($MOD.Marcados)) {
        $d = Get-ModuloDef $id
        $h += "<a href=""#$($d.Ancora)"">$(NumeroSecao $MOD $id). $($d.Titulo)</a>"
    }
    $h
}

function CardsModulos($MOD) {
    $cards = @()
    foreach ($id in @($MOD.Marcados)) {
        $d = Get-ModuloDef $id
        $r = $MOD[$d.Chave]
        if ($r -and $r.Ok) { $cards += @(& $d.Cards $r) }
    }
    $cards
}

function SecoesModulos($MOD) {
    $h = ''
    foreach ($id in @($MOD.Marcados)) {
        $d = Get-ModuloDef $id
        $num = NumeroSecao $MOD $id
        $r = $MOD[$d.Chave]
        if ($r -and $r.Ok) {
            $h += [string](& $d.Secao $r $num)
        } else {
            $erro = if ($r) { $r.Erro } else { 'o módulo não foi executado' }
            $h += "<section id=""$($d.Ancora)""><h2><span>$num</span>$($d.Titulo)</h2><p class=""lead"">Não foi possível ler: $(HE $erro)</p></section>"
        }
    }
    $h
}

function CssModulos($MOD) {
    if (-not @($MOD.Marcados).Count) { return '' }
    '.t-ok{background:var(--l3bg);color:var(--l3)}.t-bad{background:var(--redbg);color:var(--red)}' +
    'table.mini td.sev-alta{color:var(--red);font-weight:600}table.mini td.sev-media{color:var(--amb);font-weight:600}' +
    'table.mini td.sev-baixa{color:var(--blu)}table.mini td.sev-ok{color:var(--l3)}' +
    'h3.mod-sub{margin:22px 0 8px}' +
    'h4.gpo{margin:18px 0 2px;font-size:15px;text-transform:none;letter-spacing:0;color:var(--ink)}p.gpo-meta{margin:2px 0 6px;color:var(--mute);font-size:13px}ul.gpo-det{margin:0;padding-left:18px}'
}

function MetodoModulos($MOD) {
    $itens = @()
    foreach ($id in @($MOD.Marcados)) { $itens += (Get-ModuloDef $id).Metodo }
    if (-not $itens.Count) { return '' }
    " Também foram verificados: $($itens -join '; ')."
}
