
function AccessRecordSet {

    #$db = $Access.OpenCurrentDatabase("C:\Users\czJaBeck\Onedrive - LEGO\Documents\Database74.accdb") # -ComObject Access.Application.Database

    #$Access.Visible = $true

    #$rs = $Access.Run("Test")

    #$rs.MoveFirst()
    #$rs.MoveLast()

    #$rows = $rs.RecordCount()
    #$output = $rs.GetRows($rows)


    #$rs.Close()
    #$db.Close()

}
function GetApp($scriptPath, $shelperPath) {

    # Write-Information $scriptPath
    
    $Access = New-Object -ComObject Access.Application

    #just open to be sure that application had shelper openend
    $db = $Access.OpenCurrentDatabase($shelperPath) # -ComObject Access.Application.Database

    $Access.Visible = -1

    #connect to database using GetDb shelper function which is wrapper for GetObject
    $TargetApp = $Access.Run("GetApp", [ref]$scriptPath) #use [ref] for optinal COM parameters
    
    return $TargetApp

}
function ConnectDb($scriptPath, $shelperPath) {

    # Write-Information $scriptPath
    
    $Access = New-Object -ComObject Access.Application

    #just open to be sure that application had shelper openend
    $db = $Access.OpenCurrentDatabase($shelperPath) # -ComObject Access.Application.Database

    $Access.Visible = -1

    #connect to database using GetDb shelper function which is wrapper for GetObject
    $output = $Access.Run("GetDb", [ref]$scriptPath) #use [ref] for optinal COM parameters
    
    return $Access

}

function InitDb($scriptPath) {

    Write-Information $scriptPath
    
    $Access = New-Object -ComObject Access.Application

    $db = $Access.OpenCurrentDatabase($scriptPath) # -ComObject Access.Application.Database

    $Access.Visible = -1

    return $Access

}


function CloseDb($Access) {

    $Access.Quit(2) 

}

function AccessJSON($Access, $command) {

    $output = $Access.Run("DbJson", [ref]$command) #use [ref] for optional COM parameters

    # $myTestObject = $output | ConvertFrom-Json

    # Write-Information $myTestObject

    return $output

}
function AccessCmd($Access, $command, $arguments) {

    $output = $Access.Run("DbCommand", [ref]$command, [ref]$arguments) #use [ref] for optinal COM parameters

    # $myTestObject = $output | ConvertFrom-Json

    # Write-Information $myTestObject

    return $output

}