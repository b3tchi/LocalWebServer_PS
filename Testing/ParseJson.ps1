$jsonS = "{""name"":""getItems"",""keys"":[]}"

$json = $jsonS | ConvertFrom-Json

Write-Host $json.name