#Requires -Version 5.1
<#
.SYNOPSIS
  Gera os 7 scripts por versão do Windows Server, valida, compila o RelatorioAD.exe,
  assina (se houver certificado), empacota e escaneia com o Defender.

.EXEMPLO
  .\build.ps1 -SemScan
  .\build.ps1 -CertThumbprint 0123ABCD...
#>
param(
    [string]$CertThumbprint,
    [switch]$SemScan
)

$ErrorActionPreference = 'Stop'
$raiz     = $PSScriptRoot
$dirBuild = Join-Path $raiz 'build\scripts'
$dirDist  = Join-Path $raiz 'dist'
$dirSrc   = Join-Path $raiz 'src'
$lint     = Join-Path $raiz 'tools\Test-ScriptCompat.ps1'
$csc      = 'C:\Windows\Microsoft.NET\Framework64\v3.5\csc.exe'
$utf8Bom  = New-Object System.Text.UTF8Encoding($true)

# Id → base, rótulo gravado em $VersaoSO e alvo do lint
$versoes = @(
    @{ Id = '2008R2'; Rotulo = '2008 R2'; Base = 'Relatorio-PS2.ps1';  Alvo = '2.0' }
    @{ Id = '2012';   Rotulo = '2012';    Base = 'Relatorio-PS3.ps1';  Alvo = '3.0' }
    @{ Id = '2012R2'; Rotulo = '2012 R2'; Base = 'Relatorio-PS3.ps1';  Alvo = '3.0' }
    @{ Id = '2016';   Rotulo = '2016';    Base = 'Relatorio-PS51.ps1'; Alvo = '5.1' }
    @{ Id = '2019';   Rotulo = '2019';    Base = 'Relatorio-PS51.ps1'; Alvo = '5.1' }
    @{ Id = '2022';   Rotulo = '2022';    Base = 'Relatorio-PS51.ps1'; Alvo = '5.1' }
    @{ Id = '2025';   Rotulo = '2025';    Base = 'Relatorio-PS51.ps1'; Alvo = '5.1' }
)
function ScriptDe($v) { Join-Path $dirBuild "Relatorio-$($v.Id).ps1" }

function Etapa([string]$t) { Write-Host ''; Write-Host "== $t" -ForegroundColor Cyan }

function Assinar([string]$arquivo) {
    $sig = Set-AuthenticodeSignature -FilePath $arquivo -Certificate $script:cert `
        -TimestampServer 'http://timestamp.digicert.com' -HashAlgorithm SHA256
    if ($sig.Status -ne 'Valid') { throw "Falha ao assinar ${arquivo}: $($sig.Status) $($sig.StatusMessage)" }
    Write-Host "  assinado: $(Split-Path $arquivo -Leaf)"
}

# Limpa as saídas anteriores
foreach ($d in @((Join-Path $raiz 'build'), $dirDist)) {
    if (Test-Path -LiteralPath $d) { Remove-Item -LiteralPath $d -Recurse -Force }
}
[void](New-Item -ItemType Directory -Path $dirBuild -Force)
[void](New-Item -ItemType Directory -Path $dirDist -Force)

# 1. Scripts por versão
Etapa '1. Gerando scripts por versão'
$padrao = "(?m)^(\s*\[string\]\`$VersaoSO\s*=\s*)''"
foreach ($v in $versoes) {
    $base = Join-Path $raiz "scripts\$($v.Base)"
    $txt  = [IO.File]::ReadAllText($base, [Text.Encoding]::UTF8)
    $n    = ([regex]::Matches($txt, $padrao)).Count
    if ($n -ne 1) { throw "$($v.Base): esperado 1 parâmetro `$VersaoSO = '' e foram encontrados $n" }
    $txt = [regex]::Replace($txt, $padrao, "`${1}'$($v.Rotulo)'")
    [IO.File]::WriteAllText((ScriptDe $v), $txt, $utf8Bom)
    Write-Host "  Relatorio-$($v.Id).ps1  ← $($v.Base)  (VersaoSO = '$($v.Rotulo)')"
}

# 2. Lint (no mesmo processo: abrir outro powershell.exe com Bypass é bloqueado em alguns ambientes)
Etapa '2. Lint de compatibilidade'
foreach ($v in $versoes) {
    $global:LASTEXITCODE = 0
    & $lint -Path (ScriptDe $v) -Alvo $v.Alvo
    if ($LASTEXITCODE -ne 0) { throw "Lint reprovou Relatorio-$($v.Id).ps1 no alvo $($v.Alvo)" }
}

# Certificado (opcional)
$cert = $null
if ($CertThumbprint) {
    $tp = $CertThumbprint -replace '\s', ''
    $cert = Get-ChildItem -Path "Cert:\CurrentUser\My\$tp", "Cert:\LocalMachine\My\$tp" -CodeSigningCert -ErrorAction SilentlyContinue |
        Select-Object -First 1
    if (-not $cert) { throw "Certificado de assinatura de código $tp não encontrado em CurrentUser\My nem LocalMachine\My" }
    Etapa '   Assinando scripts'
    foreach ($v in $versoes) { Assinar (ScriptDe $v) }
} else {
    Write-Host ''
    Write-Host '   (sem -CertThumbprint: assinatura pulada)' -ForegroundColor DarkGray
}

# 3. Compilação
Etapa '3. Compilando'
if (-not (Test-Path -LiteralPath $csc)) { throw "csc.exe do .NET 3.5 não encontrado: $csc" }
$fontes = @(Get-ChildItem -LiteralPath $dirSrc -Filter '*.cs' -ErrorAction SilentlyContinue | ForEach-Object { $_.FullName })
if ($fontes.Count -eq 0) { throw "Nenhum .cs em $dirSrc (os fontes do aplicativo são da Fase 3)" }
foreach ($f in @('app.manifest', 'RelatorioAD.exe.config')) {
    if (-not (Test-Path -LiteralPath (Join-Path $dirSrc $f))) { throw "Falta src\$f" }
}
$exe = Join-Path $dirDist 'RelatorioAD.exe'
$argsCsc = @(
    '/nologo', '/utf8output', '/target:winexe', '/platform:anycpu', '/optimize+', '/warnaserror+',
    "/win32manifest:$(Join-Path $dirSrc 'app.manifest')", "/out:$exe",
    '/r:System.dll', '/r:System.Drawing.dll', '/r:System.Windows.Forms.dll'
)
foreach ($v in $versoes) { $argsCsc += "/resource:$(ScriptDe $v),RelatorioAD.Scripts.$($v.Id).ps1" }
$argsCsc += $fontes
$saida = & $csc $argsCsc 2>&1 | ForEach-Object { "$_" }
$saida | ForEach-Object { Write-Host "  $_" }
if ($LASTEXITCODE -ne 0) { throw "csc falhou (código $LASTEXITCODE)" }
if ($saida -match 'warning') { throw 'csc emitiu avisos' }
Copy-Item -LiteralPath (Join-Path $dirSrc 'RelatorioAD.exe.config') -Destination $dirDist
if ($cert) { Assinar $exe }

# 5. Pacote
Etapa '5. Empacotando'
Add-Type -AssemblyName System.IO.Compression, System.IO.Compression.FileSystem
$zip = Join-Path $dirDist 'RelatorioAD.zip'
$arq = [IO.Compression.ZipFile]::Open($zip, [IO.Compression.ZipArchiveMode]::Create)
try {
    foreach ($f in @('RelatorioAD.exe', 'RelatorioAD.exe.config')) {
        [void][IO.Compression.ZipFileExtensions]::CreateEntryFromFile($arq, (Join-Path $dirDist $f), $f)
    }
} finally { $arq.Dispose() }

# 6. Defender (por último, para cobrir também o .zip e os scripts embutidos)
if ($SemScan) {
    Write-Host ''
    Write-Host '   (-SemScan: scan do Defender pulado)' -ForegroundColor DarkGray
} else {
    Etapa '6. Scan do Microsoft Defender'
    $mp = Join-Path $env:ProgramFiles 'Windows Defender\MpCmdRun.exe'
    if (-not (Test-Path -LiteralPath $mp)) { throw "MpCmdRun.exe não encontrado: $mp" }
    $st = Get-MpComputerStatus -ErrorAction SilentlyContinue
    if ($st) { Write-Host "  Assinaturas: $($st.AntivirusSignatureVersion) ($($st.AntivirusSignatureLastUpdated))" }
    foreach ($alvo in @($dirDist, $dirBuild)) {
        $saidaMp = & $mp -Scan -ScanType 3 -File $alvo -DisableRemediation 2>&1 | ForEach-Object { "$_" }
        $codMp = $LASTEXITCODE
        $saidaMp | Where-Object { $_.Trim() } | ForEach-Object { Write-Host "  $_" }
        if ($codMp -ne 0) { throw "Defender: MpCmdRun terminou com código $codMp em $alvo`n$($saidaMp -join "`n")" }
    }
}

Etapa 'Pronto'
Write-Host ("  RelatorioAD.exe  {0:N0} bytes" -f (Get-Item -LiteralPath $exe).Length)
Write-Host "  SHA256 exe: $((Get-FileHash -LiteralPath $exe -Algorithm SHA256).Hash)"
Write-Host "  SHA256 zip: $((Get-FileHash -LiteralPath $zip -Algorithm SHA256).Hash)"
