#!/snap/bin/powershell

Import-Module ./AccessRunDb.ps1


# $RootPath = "/home/jan/DevProjects/SvelteDnD/public"
$RootPath = "C:\Users\czJaBeck\Documents\svelte-ie-template\public"
# $RootPath = $Pwd


#MS ACCESS FILE OPENING
$scriptPath = $PSScriptRoot
$scriptPath = Split-Path -Parent $PSCommandPath
$dbName = "TestDb.accdb"
$dbFullPath= $scriptPath + "\" + $dbName

# $app = InitDb $dbFullPath
$app = ConnectDb $dbFullPath

#HTTP LISTENER PREPARATION
$Hso = New-Object Net.HttpListener
$Hso.Prefixes.Add("http://localhost:8001/")
$Hso.Start()
# Register-EngineEvent PowerShell.Exiting –Action { 
#     Write-Host "Close Event"
# }
try{
	"$(Get-Date -Format s) Custom  Powershell webserver started."
  While ($Hso.IsListening) {
    $HC = $Hso.GetContext()
    $HRes = $HC.Response
    # Write-Output $HC.Request
    # $HRes.Headers.Add("Content-Type","text/plain")
    $RequestItem = $HC.Request
    
    $RECEIVED = '{0} {1}' -f $RequestItem.httpMethod, $RequestItem.Url.LocalPath
    Write-Host $RECEIVED  

    # stop powershell webserver, nothing to do here
    if($RECEIVED -eq "GET /quit"){
			"$(Get-Date -Format s) Stopping powershell webserver..."
      $HRes.Close()
      break
    }
    
    switch($RequestItem.httpMethod){

      "GET"{
        $Path = (Join-Path $RootPath ($RequestItem).RawUrl)
        if($Path -like "*.css"){
          Write-Host "Css"
          $HRes.Headers.Add("Content-Type","text/css")
        }
        Write-Host $Path
        if(Test-Path $Path -PathType Leaf){
          # Buf =Get-Content(Join-Path $Pwd ($HC.Request).RawUrl) -Raw
          $Buf = [Text.Encoding]::UTF8.GetBytes((Get-Content $Path -Raw))
          $HRes.ContentLength64 = $Buf.Length
          $HRes.OutputStream.Write($Buf,0,$Buf.Length)
        }else{
          Write-Host "file not found"
        }
      break
      }

      "POST"{
      # "OPTIONS"{
        Write-Host "Post"
        # only if there is body data in the request
        if ($RequestItem.HasEntityBody){

          # set default message to error message (since we just stop processing on error)
          # $RESULT = "Received corrupt or incomplete form data"

          # check content type
          if ($RequestItem.ContentType){

            if($RECEIVED -eq "POST /query"){
              # if($RECEIVED -eq "OPTIONS /query"){
            
              # retrieve boundary marker for header separation
              # $BOUNDARY = $NULL
              # if ($RequestItem.ContentType -match "boundary=(.*);")
              # {	$BOUNDARY = "--" + $MATCHES[1] }
              # else
              # { # marker might be at the end of the line
              # 	if ($RequestItem.ContentType -match "boundary=(.*)$")
              # 	{ $BOUNDARY = "--" + $MATCHES[1] }
              # }
              # if ($BOUNDARY)
              # { # only if header separator was found

              # read complete header (inkl. file data) into string
              
              $inputStream = $RequestItem.InputStream
              $Encoding = $RequestItem.ContentEncoding
              
              
              $READER = New-Object System.IO.StreamReader($inputStream, $Encoding)
              $DATA = $READER.ReadToEnd()
              $READER.Close()
              $RequestItem.InputStream.Close()

              # }
              Write-Host "Request Data:"
              Write-Host $DATA
              # TODO Prepare response Script
              
              # $JSONRESPONSE = AccessJSON $app "Test"
              $JSONRESPONSE = AccessCmd $app "DbMsg" "Test Messagebox"

              $HRes.AddHeader("Content-Type","text/json")
              $HRes.AddHeader("Last-Modified", [DATETIME]::Now.ToString('r'))
              $HRes.AddHeader("Server", "Powershell Webserver/1.2 on ")

              # return HTML answer to caller
              $BUFFER = [Text.Encoding]::UTF8.GetBytes($JSONRESPONSE )
              $HRes.ContentLength64 = $BUFFER.Length
              $HRes.OutputStream.Write($BUFFER, 0, $BUFFER.Length)
            }
          }
        }
      break 
      } 
    }
    $HRes.Close()
  }
}finally{
  #Close Listener  
  $Hso.Stop()
  $Hso.Close()

  #Close MS ACCESS
  CloseDb $app 

  Write-Host "Closed"
}
