function GetApp($scriptPath) {
  [void][System.Reflection.Assembly]::LoadWithPartialName("Microsoft.VisualBasic")

  # $TargetApp = $Access.Run("GetApp", [ref]$scriptPath) #use [ref] for optinal COM parameters
  $TargetApp = [Microsoft.VisualBasic.Interaction]::GetObject($scriptPath)

  return $TargetApp
}

function CloseDb($Access) {
  $Access.Quit(2)
}

function AccessJSON($Access, $command) {
  $rs = $Access.Run("QueryGet", [ref]$command) #use [ref] for optional COM parameters
  $json = ConvertFromRs($rs) | ConvertTo-Json

  # Write-Information $myTestObject
  return $json
}

function AccessCmd($app, $command, $arguments) {
  $data = $arguments."data" #get first object in array

  # Fill Json Data
  $db = $app.CurrentDb()

  foreach ($item in $data) {
    ConvertToRs $db $item
  }

  $output = $app.Run("ExecCommand", [ref]$command) #use [ref] for optinal COM parameters
  # $myTestObject = $output | ConvertFrom-Json
  # Write-Information $myTestObject

  #return output tbd
  return $output
}

function ConvertToRs($db, $psO) {
  $itemprops = $psO.PsObject.Properties
  $table = $itemprops | Select-Object -First 1

  $tableName = $table.Name
  $records = $table.Value

  #Open recordset
  $db.Execute("DELETE FROM $tableName")
  $rs = $db.OpenRecordset($tableName)

  foreach ($record in $records) {

    # $tableName
    $fields = $record.PsObject.Properties
    # $fields = $fields | Get-Member -MemberType NoteProperty # | Select-Object -Property Name
    # write-host ------
    $rs.AddNew()

    foreach ($field in $fields) {
      # Access the name of the property
      # write-host $object_properties.Name
      # Access the value of the property
      try {
        $rsfld = $rs.Fields($field.name)
      }
      catch {
        $rsfld = $null
        write-host $field.name + " not found in $tablename"
      }

      if ($null -ne $rsfld) {
        $value = $field.Value
        if ($value.GetType().Name -eq 'String') {
          $rs.Fields($field.Name).Value = "$value" #$strA
        }
        else {
          $rs.Fields($field.Name).Value = $value
        }
      }
      # $fld = $rs.Fields($field.Name)
      # write-host $field.Name $field.Value $fld.Name
    }

    $rs.Update()
  }

  $rs.close()
}

function ConvertFromRs($rs) {
  $rs.MoveLast()
  $rs.MoveFirst()
  # $rs.RecordCount

  $fldCount = $rs.Fields.Count
  $data = @()
  while ($rs.EOF -ne $true) {
    $rec = @{ }

    for ($i = 0; $i -lt $fldCount; $i++) {
      $rec | Add-Member -NotePropertyName $rs.Fields($i).Name -NotePropertyValue $rs.Fields($i).Value
    }

    # $rec | ConvertTo-Json
    $data += $rec

    $rs.MoveNext()
  }

  # $rs.Close()

  return $data
}
