
Import-Module ./AccessRunDb.ps1

# $scriptPath = Split-Path $psise.CurrentFile.FullPath #$Pwd.Path.ToString()
$scriptPath = $PSScriptRoot
$scriptPath = Split-Path -Parent $PSCommandPath
$shelperName = "shelper.accdb"
$shelperPath = $scriptPath + "\" + $shelperName

$app = GetApp $dbFullPath $shelperPath
# $dbName = "ApiDb.accdb"
# $dbFullPath= $scriptPath + "\" + $dbName

# $app = InitDb $dbFullPath

$jsonR = AccessJSON $app "Test"

Write-Information $jsonR 

CloseDb $app 