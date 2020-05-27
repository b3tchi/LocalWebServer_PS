
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
function GetApp($scriptPath) {

    [void][System.Reflection.Assembly]::LoadWithPartialName("Microsoft.VisualBasic")
        # Write-Information $scriptPath
    
    # $Access = New-Object -ComObject Access.Application

    #just open to be sure that application had shelper openend
    # $db = $Access.OpenCurrentDatabase($shelperPath) # -ComObject Access.Application.Database

    # $Access.Visible = -1

    #connect to database using GetDb shelper function which is wrapper for GetObject
    # $TargetApp = $Access.Run("GetApp", [ref]$scriptPath) #use [ref] for optinal COM parameters
    $TargetApp = [Microsoft.VisualBasic.Interaction]::GetObject($scriptPath) 

    return $TargetApp

    # $db.close
    # $Access.Quit

}

function GetApp_Old($scriptPath, $shelperPath) {

    # Write-Information $scriptPath
    
    $Access = New-Object -ComObject Access.Application

    #just open to be sure that application had shelper openend
    $db = $Access.OpenCurrentDatabase($shelperPath) # -ComObject Access.Application.Database

    $Access.Visible = -1

    #connect to database using GetDb shelper function which is wrapper for GetObject
    $TargetApp = $Access.Run("GetApp", [ref]$scriptPath) #use [ref] for optinal COM parameters
    
    return $TargetApp

    $db.close
    $Access.Quit

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

function ConvertToRs($db, $psO){

    $itemprops = $psO.PsObject.Properties
    $table = $itemprops | Select-Object -First 1 

    $tableName = $table.Name
    $records = $table.Value
    
    #Open recordset
    $db.Execute("DELETE FROM $tableName")
    $rs = $db.OpenRecordset($tableName)
    
    foreach($record in $records){

        # $tableName
        $fields = $record.PsObject.Properties
        # $fields = $fields | Get-Member -MemberType NoteProperty # | Select-Object -Property Name
        # write-host ------
        $rs.AddNew()

        foreach($field in $fields){
            # Access the name of the property
            # write-host $object_properties.Name
            # Access the value of the property
            
            $value = $field.Value
            if($value.GetType().Name -eq 'String') {
                $rs.Fields($field.Name).Value = "$value" #$strA
            } else {
                $rs.Fields($field.Name).Value = $value
            }
            
            # $fld = $rs.Fields($field.Name)
            # write-host $field.Name $field.Value $fld.Name
        }

        $rs.Update()
    }
    
    $rs.close()

}

function ConvertFromRs($db, $queryName){

    $rs = $db.OpenRecordset($queryName)
    $rs.MoveLast()
    $rs.MoveFirst()
    # $rs.RecordCount

    $fldCount = $rs.Fields.Count
    $data = @()
    while($rs.EOF -ne $true){
        $rec = @{}
        
        for ($i = 0; $i -lt $fldCount; $i++) {
            $rec | Add-Member -NotePropertyName $rs.Fields($i).Name -NotePropertyValue $rs.Fields($i).Value
        }
        
        # $rec | ConvertTo-Json
        $data += $rec

        $rs.MoveNext()
    }

    $rs.Close()

    return $data
}
