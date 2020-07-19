
Import-Module $PSScriptRoot/../AccessRunDb.ps1
Import-Module $PSScriptRoot/../Config.ps1


# $scriptPath = Split-Path $psise.CurrentFile.FullPath #$Pwd.Path.ToString()
# $scriptPath = $PSScriptRoot
# $scriptPath = Split-Path -Parent $PSCommandPath
# $dbName = "Test.accdb"
# $dbFullPath= $scriptPath + "\" + $dbName
# $dbFullPath= "C:\Users\czJaBeck\Documents\Vbox\LocalWeb_Ps\TestDb.accdb"

$app = GetApp $dbFullPath

function jsonTest(){

  $jsonR = AccessJSON $app "Test"

  Write-Information $jsonR 

}

function cmdTestx(){
  $jsonS = @'
{"name":"SaveTitle","arguments":{"pText":"Itemx","pItemID":23}}
'@

  $pson = $jsonS | ConvertFrom-Json

  $res = AccessCmd $app $pson.name $pson.arguments
  Write-Host ($res | Format-List | Out-String)
}

function cmdTest(){
  $jsonS = @'
{"name":"UpdateStage","arguments":{"pStageID":6,"pItemID":8}}
'@

  $pson = $jsonS | ConvertFrom-Json

  $res = AccessCmd $app $pson.name $pson.arguments
  Write-Host ($res | Format-List | Out-String)
}

#what to test
cmdTestx
# jsonTest
