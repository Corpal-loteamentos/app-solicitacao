param(
    [Parameter(Mandatory = $true)]
    [string]$AzureClientId,

    [Parameter(Mandatory = $true)]
    [string]$AzureAppIdUri
)

$parsedGuid = [Guid]::Empty
if (-not [Guid]::TryParse($AzureClientId, [ref]$parsedGuid)) {
    throw "AzureClientId deve ser um GUID valido."
}

if (-not $AzureAppIdUri.StartsWith("api://")) {
    throw "AzureAppIdUri deve comecar com api://"
}

$scriptDir = Split-Path -Parent $MyInvocation.MyCommand.Path
$workDir = Join-Path $env:TEMP ("corpal-teams-" + [Guid]::NewGuid())
New-Item -ItemType Directory -Path $workDir | Out-Null

try {
    $manifestTemplate = Join-Path $scriptDir "manifest.template.json"
    $manifest = Get-Content -Raw -Path $manifestTemplate
    $manifest = $manifest.Replace("__AZURE_CLIENT_ID__", $AzureClientId)
    $manifest = $manifest.Replace("__AZURE_APP_ID_URI__", $AzureAppIdUri)

    $manifestPath = Join-Path $workDir "manifest.json"
    Set-Content -Path $manifestPath -Value $manifest -Encoding utf8

    $colorIcon = Join-Path $scriptDir "color.png"
    $outlineIcon = Join-Path $scriptDir "outline.png"
    Copy-Item $colorIcon, $outlineIcon -Destination $workDir

    $output = Join-Path $scriptDir "Corpal-Solicitacoes-Teams-1.0.29.zip"
    if (Test-Path $output) {
        Remove-Item $output -Force
    }

    Compress-Archive -Path @(
        $manifestPath,
        (Join-Path $workDir "color.png"),
        (Join-Path $workDir "outline.png")
    ) -DestinationPath $output

    Write-Host "Pacote gerado: $output"
}
finally {
    Remove-Item $workDir -Recurse -Force -ErrorAction SilentlyContinue
}
