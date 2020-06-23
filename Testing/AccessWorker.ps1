
Import-Module $PSScriptRoot/../AccessRunDb.ps1

# $scriptPath = Split-Path $psise.CurrentFile.FullPath #$Pwd.Path.ToString()
# $scriptPath = $PSScriptRoot
# $scriptPath = Split-Path -Parent $PSCommandPath
# $dbName = "Test.accdb"
# $dbFullPath= $scriptPath + "\" + $dbName
$dbFullPath= "C:\Users\czJaBeck\Documents\Vbox\LocalWeb_Ps\TestDb.accdb"

$app = GetApp $dbFullPath

function jsonTest(){

  $jsonR = AccessJSON $app "Test"

  Write-Information $jsonR 

}

function cmdTest(){
$jsonS = @'
{"pStageID":1,"pItemID":1}
'@

$pson = $jsonS | ConvertFrom-Json

  $res = AccessCmd $app "UpdateStage" $pson
}

#what to test
cmdXTest
# jsonTest