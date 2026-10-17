param()

$ErrorActionPreference = 'Stop'
$root = $PSScriptRoot
$out = Join-Path $root 'build'
New-Item -ItemType Directory -Force -Path $out | Out-Null

$vswhere = Join-Path ${env:ProgramFiles(x86)} 'Microsoft Visual Studio\Installer\vswhere.exe'
if (-not (Test-Path $vswhere)) { throw '需要安装 Visual Studio Build Tools（含 C# 编译器）。' }
$vs = & $vswhere -latest -products '*' -property installationPath
if (-not $vs) { throw '未找到 Visual Studio Build Tools。' }
$csc = Join-Path $vs 'MSBuild\Current\Bin\Roslyn\csc.exe'
if (-not (Test-Path $csc)) { throw "未找到 C# 编译器：$csc" }

function Find-GacAssembly([string]$folder, [string]$name, [string]$file) {
    $gac = Join-Path $env:WINDIR "assembly\$folder\$name"
    if (-not (Test-Path $gac)) { throw "未找到 Office/.NET 组件：$name" }
    $result = Get-ChildItem -LiteralPath $gac -Recurse -File -Filter $file |
        Sort-Object { try { [version]($_.Directory.Name.Split('_')[0]) } catch { [version]'0.0' } } -Descending |
        Select-Object -First 1
    if (-not $result) { throw "未找到 Office/.NET 组件：$file" }
    return $result.FullName
}

$office = Find-GacAssembly 'GAC_MSIL' 'office' 'OFFICE.DLL'
$word = Find-GacAssembly 'GAC_MSIL' 'Microsoft.Office.Interop.Word' 'Microsoft.Office.Interop.Word.dll'
$ext = Find-GacAssembly 'GAC' 'Extensibility' 'extensibility.dll'
$dll = Join-Path $out 'CodeStudio.Addin.dll'
$exe = Join-Path $out 'CodeStudio-Word-Setup.exe'
$snk = Join-Path $out 'CodeStudio.snk'

if (-not (Test-Path $snk)) {
    $rsa = [System.Security.Cryptography.RSACryptoServiceProvider]::new(2048)
    try { [IO.File]::WriteAllBytes($snk, $rsa.ExportCspBlob($true)) }
    finally { $rsa.Dispose() }
}

$pluginArgs = @('/nologo', '/target:library', '/platform:x64', '/optimize+', "/keyfile:$snk", "/out:$dll",
    '/reference:System.Windows.Forms.dll', '/reference:System.Drawing.dll', "/reference:$office", "/reference:$word", "/reference:$ext",
    (Join-Path $root 'src\AssemblyInfo.cs'), (Join-Path $root 'src\Connect.cs'),
    (Join-Path $root 'src\WordCodePane.cs'), (Join-Path $root 'src\CodeRenderer.cs'))
& $csc @pluginArgs
if ($LASTEXITCODE -ne 0) { throw '插件编译失败。' }

$installerArgs = @('/nologo', '/target:winexe', '/platform:x64', '/optimize+', "/out:$exe",
    '/reference:System.Windows.Forms.dll', "/resource:$dll,CodeStudio.Addin.dll", (Join-Path $root 'src\Setup.cs'))
& $csc @installerArgs
if ($LASTEXITCODE -ne 0) { throw '安装包编译失败。' }

Get-Item $dll, $exe | Select-Object Name, Length, LastWriteTime
