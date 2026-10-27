#Requires -Version 5.1
<#
.SYNOPSIS
  Verifica se um script usa só o que existe na versão de PowerShell do alvo.
  Faz o parse com o parser do PowerShell e reprova construções do PS 3+ (alvo 2.0)
  e do PS 5 (alvo 3.0). Saída: "Linha: problema"; exit code 1 se houver falhas.

.EXEMPLO
  .\tools\Test-ScriptCompat.ps1 -Path .\scripts\Relatorio-PS2.ps1 -Alvo 2.0
#>
param(
    [Parameter(Mandatory = $true)][string]$Path,
    [Parameter(Mandatory = $true)][ValidateSet('2.0', '3.0', '5.1')][string]$Alvo
)

$ErrorActionPreference = 'Stop'
$Path = (Resolve-Path -LiteralPath $Path).Path
$vAlvo = [version]$Alvo
$falhas = New-Object System.Collections.ArrayList
function Falha([int]$linha, [string]$msg) { [void]$falhas.Add([pscustomobject]@{ Linha = $linha; Msg = $msg }) }

$tokens = $null; $erros = $null
$ast = [System.Management.Automation.Language.Parser]::ParseFile($Path, [ref]$tokens, [ref]$erros)
foreach ($e in $erros) { Falha $e.Extent.StartLineNumber "erro de sintaxe: $($e.Message)" }

# #Requires acima do alvo
if ($ast.ScriptRequirements -and $ast.ScriptRequirements.RequiredPSVersion -and $ast.ScriptRequirements.RequiredPSVersion -gt $vAlvo) {
    Falha 1 "#Requires -Version $($ast.ScriptRequirements.RequiredPSVersion) é maior que o alvo $Alvo"
}

$todos = $ast.FindAll({ $true }, $true)
$cmds  = @($todos | Where-Object { $_ -is [System.Management.Automation.Language.CommandAst] })

function Params($c) {
    @($c.CommandElements | Where-Object { $_ -is [System.Management.Automation.Language.CommandParameterAst] } | ForEach-Object { $_.ParameterName })
}
$gci   = @('Get-ChildItem', 'gci', 'dir', 'ls')
$where = @('Where-Object', 'where', '?')
$each  = @('ForEach-Object', 'foreach', '%')

if ($vAlvo -lt [version]'5.0') {
    foreach ($c in $cmds) {
        if ($gci -contains $c.GetCommandName() -and (Params $c) -contains 'Depth') {
            Falha $c.Extent.StartLineNumber 'Get-ChildItem -Depth não existe antes do PS 5.0'
        }
    }
    foreach ($n in $todos) {
        if ($n -is [System.Management.Automation.Language.TypeDefinitionAst]) { Falha $n.Extent.StartLineNumber "class/enum só existem no PS 5.0+" }
        if ($n -is [System.Management.Automation.Language.UsingStatementAst]) { Falha $n.Extent.StartLineNumber "'using' só existe no PS 5.0+" }
        if ($n -is [System.Management.Automation.Language.InvokeMemberExpressionAst] -and $n.Static -and
            $n.Member -is [System.Management.Automation.Language.StringConstantExpressionAst] -and $n.Member.Value -eq 'new') {
            Falha $n.Extent.StartLineNumber '::new() só existe no PS 5.0+'
        }
    }
}

if ($vAlvo -lt [version]'3.0') {
    $cmdsProibidos = @('Get-SmbShare', 'Get-SmbShareAccess', 'Get-CimInstance', 'Invoke-RestMethod', 'Invoke-WebRequest', 'ConvertTo-Json', 'ConvertFrom-Json')
    foreach ($c in $cmds) {
        $nome = $c.GetCommandName()
        $ps = Params $c
        if ($cmdsProibidos -contains $nome) { Falha $c.Extent.StartLineNumber "$nome não existe no PS 2.0" }
        if ($gci -contains $nome) {
            foreach ($p in @('Directory', 'File', 'Attributes', 'Hidden', 'ReadOnly', 'System')) {
                if ($ps -contains $p) { Falha $c.Extent.StartLineNumber "Get-ChildItem -$p não existe no PS 2.0" }
            }
        }
        if ($nome -eq 'Get-Acl' -and $ps -contains 'LiteralPath') { Falha $c.Extent.StartLineNumber 'Get-Acl -LiteralPath não existe no PS 2.0' }
        if (($where + $each) -contains $nome) {
            $temBloco = @($c.CommandElements | Select-Object -Skip 1 | Where-Object { $_ -is [System.Management.Automation.Language.ScriptBlockExpressionAst] }).Count
            if (-not $temBloco) { Falha $c.Extent.StartLineNumber "$nome sem bloco de script só existe no PS 3.0+" }
        }
    }
    $tiposRx = '^(pscustomobject|ordered)$|HashSet|SortedSet|WebUtility'
    foreach ($n in $todos) {
        if ($n -is [System.Management.Automation.Language.TypeExpressionAst] -or $n -is [System.Management.Automation.Language.TypeConstraintAst]) {
            if ($n.TypeName.FullName -match $tiposRx) { Falha $n.Extent.StartLineNumber "tipo [$($n.TypeName.FullName)] não existe no PS 2.0 / CLR 2" }
        }
        if ($n -is [System.Management.Automation.Language.BinaryExpressionAst] -and
            [string]$n.Operator -match '^([IC]in|[IC]notin|Shl|Shr)$') {
            Falha $n.Extent.StartLineNumber "operador $($n.ErrorPosition.Text) não existe no PS 2.0"
        }
        # No PS 2.0, foreach sobre $null (lista vazia devolvida por função/propriedade) roda uma vez com $null
        if ($n -is [System.Management.Automation.Language.ForEachStatementAst]) {
            $c = $n.Condition
            $ehArray = $c -is [System.Management.Automation.Language.PipelineAst] -and $c.PipelineElements.Count -eq 1 -and
                $c.PipelineElements[0].Expression -is [System.Management.Automation.Language.ArrayExpressionAst]
            if (-not $ehArray) { Falha $n.Extent.StartLineNumber "foreach sem @(...): no PS 2.0 uma lista vazia roda uma vez com `$null" }
        }
    }
    # Tipos citados em texto (ex.: New-Object 'System.Collections.Generic.HashSet[string]'), fora de comentários
    foreach ($t in $tokens) {
        if ($t.Kind -eq 'Comment') { continue }
        if ($t -is [System.Management.Automation.Language.StringToken] -and $t.Value -match 'HashSet|SortedSet|WebUtility') {
            Falha $t.Extent.StartLineNumber "tipo citado em texto não existe no PS 2.0 / CLR 2: $($Matches[0])"
        }
    }
}

$nome = Split-Path $Path -Leaf
if ($falhas.Count) {
    $falhas | Sort-Object Linha, Msg -Unique | ForEach-Object { Write-Output ("{0}: {1}" -f $_.Linha, $_.Msg) }
    Write-Output "REPROVADO: $nome (alvo $Alvo) - $($falhas.Count) problema(s)"
    exit 1
}
Write-Output "OK: $nome (alvo $Alvo)"
exit 0
