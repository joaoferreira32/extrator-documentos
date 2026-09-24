param(
    [Parameter(Mandatory=$true)][string]$Dll,
    [Parameter(Mandatory=$true)][string]$Arquivo
)
# Validador oficial da Microsoft (Open XML SDK, alvo net46) no Windows
# PowerShell 5.1, sem instalar nada: carrega a DLL do pacote NuGet. Uma linha
# por erro, "parte<TAB>xpath<TAB>descricao". Usado por tests/test_ooxml.py.
Add-Type -AssemblyName WindowsBase
Add-Type -Path $Dll
$doc = [DocumentFormat.OpenXml.Packaging.SpreadsheetDocument]::Open($Arquivo, $false)
try {
    $validador = New-Object DocumentFormat.OpenXml.Validation.OpenXmlValidator(
        [DocumentFormat.OpenXml.FileFormatVersions]::Office2016)
    foreach ($e in $validador.Validate([DocumentFormat.OpenXml.Packaging.OpenXmlPackage]$doc)) {
        Write-Output ("{0}`t{1}`t{2}" -f $e.Part.Uri, $e.Path.XPath, $e.Description)
    }
} finally {
    $doc.Dispose()
}
