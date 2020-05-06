
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

    $output = $Access.Run("DbJson", [ref]$command) #use [ref] for optinal COM parameters

    # $myTestObject = $output | ConvertFrom-Json

    # Write-Information $myTestObject

    return $output

}