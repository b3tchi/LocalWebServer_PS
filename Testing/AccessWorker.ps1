
Import-Module $PSScriptRoot/../AccessRunDb.ps1
Import-Module $PSScriptRoot/../Config.ps1


# $scriptPath = Split-Path $psise.CurrentFile.FullPath #$Pwd.Path.ToString()
# $scriptPath = $PSScriptRoot
# $scriptPath = Split-Path -Parent $PSCommandPath
# $dbName = "Test.accdb"
# $dbFullPath= $scriptPath + "\" + $dbName
# $dbFullPath= "C:\Users\czJaBeck\Documents\Vbox\LocalWeb_Ps\TestDb.accdb"

# $app = GetApp $dbFullPath

function jsonTest(){
  $app = GetApp $dbFullPath

  $jsonR = AccessJSON $app "Test"

  Write-Information $jsonR 

}

function cmdTestx(){
  $app = GetApp $dbFullPath

  $jsonS = @'
{"name":"SaveTitle","arguments":{"pText":"Itemx","pItemID":23}}
'@

  $pson = $jsonS | ConvertFrom-Json

  $res = AccessCmd $app $pson.name $pson.arguments
  Write-Host ($res | Format-List | Out-String)
}

function cmdTest(){
  $app = GetApp $dbFullPath

  $jsonS = @'
{"name":"UpdateStage","arguments":{"pStageID":6,"pItemID":8}}
'@

  $pson = $jsonS | ConvertFrom-Json

  $res = AccessCmd $app $pson.name $pson.arguments
  Write-Host ($res | Format-List | Out-String)
}

function ExportCodeTest(){
  $app = GetApp $dbFullPath

  $dbCodePath= "C:\Users\czJaBeck\Documents\Vbox\LocalWeb_Ps\"
  $vbProj = AppVbProj $app
  CodeExport $vbProj $dbCodePath
  Write-Host "test"
}

function ReadModuleName(){

  $dbCodePath= "C:\Users\czJaBeck\Documents\Vbox\LocalWeb_Ps\src\Modules\Module2.bas"
  $sln = Get-Content $dbCodePath

  Write-Host NameModuleFromFile $sln


}

#what to test
# ExportCodeTest
# cmdTestx
ReadModuleName
# jsonTest
