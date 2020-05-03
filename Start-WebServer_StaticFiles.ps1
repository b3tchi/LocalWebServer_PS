$Hso = New-Object Net.HttpListener
$Hso.Prefixes.Add("http://localhost:8001/")
$Hso.Start()
# Register-EngineEvent PowerShell.Exiting –Action { 
#     Write-Host "Close Event"
#  }
try{
	"$(Get-Date -Format s) Custom  Powershell webserver started."
    While ($Hso.IsListening) {
        $HC = $Hso.GetContext()
        $HRes = $HC.Response
        # Write-Output $HC.Request
        # $HRes.Headers.Add("Content-Type","text/plain")
        $Request = (Join-Path $Pwd ($HC.Request).RawUrl)
        if($Request -like "*.css"){
            Write-Host "Css"
            $HRes.Headers.Add("Content-Type","text/css")
        }
        Write-Host $Request  
        $Buf = [Text.Encoding]::UTF8.GetBytes((Get-Content (Join-Path $Pwd ($HC.Request).RawUrl)-Raw))
        # $Buf =Get-Content(Join-Path $Pwd ($HC.Request).RawUrl) -Raw
        $HRes.ContentLength64 = $Buf.Length
        $HRes.OutputStream.Write($Buf,0,$Buf.Length)
        $HRes.Close()
    }
}finally{
    Write-Host "Closed"
    $Hso.Stop()
    $Hso.Close()
}