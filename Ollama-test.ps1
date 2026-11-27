param(
    [Parameter(Mandatory = $true)]
    [string]$ImagePath,
    [string]$Model = 'qwen3.5:4b'
)

$ErrorActionPreference = 'Stop'
$file = Get-Item -LiteralPath $ImagePath
if ($file.PSIsContainer -or $file.Extension -notin '.jpg', '.jpeg', '.png') {
    throw 'Select a JPG or PNG image of the product table.'
}

$properties = @{
    nomre    = @{ type = @('integer', 'null') }
    malin_adi = @{ type = @('string', 'null') }
    say      = @{ type = @('integer', 'null') }
    qiymet   = @{ type = @('string', 'null') }
    mebleg   = @{ type = @('string', 'null') }
}
$body = @{
    model = $Model
    stream = $false
    think = $false
    images = @([Convert]::ToBase64String([IO.File]::ReadAllBytes($file.FullName)))
    prompt = 'Read the product rows only. Copy printed values exactly, preserve decimal commas, and use null for unreadable fields. Exclude totals, contact information and bank details.'
    format = @{
        type = 'array'
        items = @{
            type = 'object'
            properties = $properties
            required = @('nomre', 'malin_adi', 'say', 'qiymet', 'mebleg')
            additionalProperties = $false
        }
    }
    options = @{ temperature = 0 }
} | ConvertTo-Json -Depth 10 -Compress

Write-Host "Reading the image with $Model. JSON appears when the request finishes."
$answer = Invoke-RestMethod -Uri 'http://127.0.0.1:11434/api/generate' `
    -Method Post -ContentType 'application/json' -Body $body -TimeoutSec 300
if (-not $answer.response) {
    throw 'Ollama returned an empty response.'
}
$answer.response
