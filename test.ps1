$ErrorActionPreference = 'Stop'
if (Get-Process WINWORD -ErrorAction SilentlyContinue) {
    throw '测试需要独立的 Word 实例；请先保存并关闭已有 Word 文档。'
}

& (Join-Path $PSScriptRoot 'build.ps1')
$exe = Join-Path $PSScriptRoot 'build\TestInsert.exe'
$csc = Join-Path (& (Join-Path ${env:ProgramFiles(x86)} 'Microsoft Visual Studio\Installer\vswhere.exe') -latest -products '*' -property installationPath) 'MSBuild\Current\Bin\Roslyn\csc.exe'
$office = Get-ChildItem (Join-Path $env:WINDIR 'assembly\GAC_MSIL\office') -Recurse -File -Filter 'OFFICE.DLL' | Sort-Object FullName -Descending | Select-Object -First 1 -ExpandProperty FullName
$word = Get-ChildItem (Join-Path $env:WINDIR 'assembly\GAC_MSIL\Microsoft.Office.Interop.Word') -Recurse -File -Filter 'Microsoft.Office.Interop.Word.dll' | Sort-Object FullName -Descending | Select-Object -First 1 -ExpandProperty FullName
& $csc /nologo /target:exe /platform:x64 "/out:$exe" /reference:System.Windows.Forms.dll "/reference:$(Join-Path $PSScriptRoot 'build\CodeStudio.Addin.dll')" "/reference:$office" "/reference:$word" (Join-Path $PSScriptRoot 'tests\TestInsert.cs')
if ($LASTEXITCODE -ne 0) { throw '测试程序编译失败。' }
& $exe
if ($LASTEXITCODE -ne 0) { throw 'Word 插入测试失败。' }
