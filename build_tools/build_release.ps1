param(
    [string]$CertificatePath = $env:FAS_SIGN_CERTIFICATE_PATH,
    [string]$CertificatePassword = $env:FAS_SIGN_CERTIFICATE_PASSWORD
)

$ErrorActionPreference = 'Stop'
$projectRoot = (Resolve-Path -LiteralPath (Join-Path $PSScriptRoot '..')).Path
Set-Location -LiteralPath $projectRoot

function Invoke-Checked([scriptblock]$Command, [string]$Description) {
    & $Command
    if ($LASTEXITCODE -ne 0) {
        throw "$Description falhou com código $LASTEXITCODE."
    }
}

function Remove-BuildDirectory([string]$RelativePath) {
    $target = [IO.Path]::GetFullPath((Join-Path $projectRoot $RelativePath))
    if (-not $target.StartsWith($projectRoot + [IO.Path]::DirectorySeparatorChar, [StringComparison]::OrdinalIgnoreCase)) {
        throw "Diretório de build fora do projeto: $target"
    }
    if (Test-Path -LiteralPath $target) {
        Remove-Item -LiteralPath $target -Recurse -Force
    }
}

function Find-SignTool {
    $tool = Get-ChildItem "${env:ProgramFiles(x86)}\Windows Kits\10\bin" -Filter signtool.exe -Recurse -ErrorAction SilentlyContinue |
        Sort-Object FullName -Descending |
        Select-Object -First 1
    if (-not $tool) { throw 'SignTool não encontrado.' }
    return $tool.FullName
}

function Sign-Artifact([string]$Path) {
    if (-not $CertificatePath) {
        Write-Warning "Sem certificado configurado: $Path permanecerá sem assinatura Authenticode."
        return
    }
    $resolvedCertificate = (Resolve-Path -LiteralPath $CertificatePath).Path
    $signTool = Find-SignTool
    & $signTool sign /fd SHA256 /td SHA256 /tr http://timestamp.digicert.com /f $resolvedCertificate /p $CertificatePassword $Path
    if ($LASTEXITCODE -ne 0) { throw "Falha ao assinar $Path." }
}

$version = (python -c "from app.core.version import APP_VERSION; print(APP_VERSION)").Trim()
Invoke-Checked { python build_tools\verify_release_version.py } 'Validação da versão'
Invoke-Checked { python -m pytest -q } 'Testes'
Invoke-Checked { python -m pyright app } 'Análise estática'
Invoke-Checked { python -m compileall -q app main.py } 'Compilação Python'
Invoke-Checked { python -m pip check } 'Validação das dependências'

Remove-BuildDirectory 'build'
Remove-BuildDirectory 'dist'
Invoke-Checked { pyinstaller --clean --noconfirm main.spec } 'Build do executável'
Sign-Artifact 'dist\FASJornada.exe'

$inno = 'C:\Program Files (x86)\Inno Setup 6\ISCC.exe'
if (-not (Test-Path -LiteralPath $inno)) {
    throw 'Inno Setup 6 não foi encontrado.'
}
Invoke-Checked { & $inno 'build_tools\FASJornada.iss' } 'Build do instalador'
$installer = "dist\FASJornada_Setup_$version.exe"
Sign-Artifact $installer

$portable = "dist\FASJornada_v$version`_portable.zip"
Compress-Archive -LiteralPath 'dist\FASJornada.exe' -DestinationPath $portable -Force
$lines = foreach ($file in @($installer, $portable)) {
    $hash = Get-FileHash -LiteralPath $file -Algorithm SHA256
    "$($hash.Hash.ToLowerInvariant())  $([IO.Path]::GetFileName($file))"
}
Set-Content -LiteralPath 'dist\SHA256SUMS.txt' -Value $lines -Encoding ascii

Write-Host "Build v$version concluído em $projectRoot\dist"
