
Import-Module ./AccessRunDb.ps1

# $scriptPath = Split-Path $psise.CurrentFile.FullPath #$Pwd.Path.ToString()
$scriptPath = $PSScriptRoot
$scriptPath = Split-Path -Parent $PSCommandPath
$dbName = "Test.accdb"
$dbFullPath= $scriptPath + "\" + $dbName

$app = GetApp $dbFullPath

$jsonR = AccessJSON $app "Test"

Write-Information $jsonR 

CloseDb $app 