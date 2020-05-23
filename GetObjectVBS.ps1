start-job {
    $sc = New-Object -ComObject MSScriptControl.ScriptControl.1
    $sc.Language = 'VBScript'

    $sc.AddCode('
      Function MyFunction(byval x)
            Set MyFunction = GetObject(x)
      End Function
    ')
  
    $Ac = $sc.codeobject.MyFunction("C:\Users\czJaBeck\Documents\Vbox\LocalWeb_Ps\TestDb.accdb")

    # $Ac.Run("Test")

    $rs = New-Object -ComObject 

    $rs = $Ac.CurrentDb.OpenRecordset("SELECT * FROM Table1")

    $rsText = $rs.GetString()

    # $sc.codeobject.MyFunction(2)
  } -runas32 | wait-job | receive-job