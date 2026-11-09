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
$workDir = Join-Path $env:TEMP ("corpal-teams-preview-" + [Guid]::NewGuid())

New-Item -ItemType Directory -Path $workDir | Out-Null

try {

    $manifestTemplate = Join-Path $scriptDir "manifest.preview.template.json"

    $manifest = Get-Content -Raw -Path $manifestTemplate

    $manifest = $manifest.Replace(
        "**AZURE_CLIENT_ID**",
        $AzureClientId
    )

    $manifest = $manifest.Replace(
        "**AZURE_APP_ID_URI**",
        $AzureAppIdUri
    )

    $manifestPath = Join-Path $workDir "manifest.json"

    Set-Content `
        -Path $manifestPath `
        -Value $manifest `
        -Encoding utf8

    $colorIcon = Join-Path $scriptDir "color.png"
    $outlineIcon = Join-Path $scriptDir "outline.png"

    Copy-Item $colorIcon, $outlineIcon -Destination $workDir

    $output = Join-Path $scriptDir "Corpal-Solicitacoes-DEV-Teams.zip"

    if (Test-Path $output) {
        Remove-Item $output -Force
    }

    Compress-Archive `
        -Path @(
            $manifestPath,
            (Join-Path $workDir "color.png"),
            (Join-Path $workDir "outline.png")
        ) `
        -DestinationPath $output

    Write-Host ""
    Write-Host "Pacote DEV gerado:"
    Write-Host $output
    Write-Host ""
}
finally {

    Remove-Item `
        $workDir `
        -Recurse `
        -Force `
        -ErrorAction SilentlyContinue
}