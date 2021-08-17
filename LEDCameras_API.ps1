$host.UI.RawUI.WindowTitle = "LedCameras"

$homeDir = "C:\SN_Scripts\LedCameras"
$jsonPath = "$homeDir\sn_data.json"
$mtd = "C:\Metadane_do_skryptow"
$ssDir = "$homeDir\ss"

$WindowsScriptShell = New-Object -ComObject wscript.shell

# API
$requestURL = '***'
$requestHeaders = @{'sntoken' = '***'; 'Content-Type' = 'application/json' }

# Dont check this SN
$dontCheck = @()

# Metadata CSV
$csv = @(
  "$mtd\meta_premium.csv",
  "$mtd\meta_city.csv",
  "$mtd\meta_super.csv",
  "$mtd\meta_pakiet.csv"
)

# HTML -File
$htmlSiteNew = "$homeDir\index_local.html"
$phpPage = "$homeDir\index_logged.php"

# HTML - Page 
$checkSession = @"
<?php
  // We need to use sessions, so you should always start sessions using the below code.
  session_start();
  // If the user is not logged in redirect to the login page...
  if (!isset(`$_SESSION['loggedin'])) {
  	header('Location: index.html');
  	exit;
  }
?>

"@

$begginning = @"
<!DOCTYPE html>
<html lang="pl-PL">

<head>
  <link rel="stylesheet" type="text/css" href="normalize.css" />
  <link rel="stylesheet" type="text/css" href="style.css" />
  <META HTTP-EQUIV="CACHE-CONTROL" CONTENT="no-cache, no-store, must-revalidate">
  <meta http-equiv="refresh" content="900">
  <meta charset="UTF-8">
  <link rel="icon" href="icon.png" />
  <title>
    LedCameras - Home
  </title>
</head>

<body>
  <div class="progress"></div>
  <div class="progress"></div>
  <nav>
  <div>
  <svg xmlns="http://www.w3.org/2000/svg" viewBox="0 0 24 24" fill="#FFFFFF"><path d="M0 0h24v24H0V0z" fill="none"/><path d="M12 2C6.48 2 2 6.48 2 12s4.48 10 10 10 10-4.48 10-10S17.52 2 12 2zM7.07 18.28c.43-.9 3.05-1.78 4.93-1.78s4.51.88 4.93 1.78C15.57 19.36 13.86 20 12 20s-3.57-.64-4.93-1.72zm11.29-1.45c-1.43-1.74-4.9-2.33-6.36-2.33s-4.93.59-6.36 2.33C4.62 15.49 4 13.82 4 12c0-4.41 3.59-8 8-8s8 3.59 8 8c0 1.82-.62 3.49-1.64 4.83zM12 6c-1.94 0-3.5 1.56-3.5 3.5S10.06 13 12 13s3.5-1.56 3.5-3.5S13.94 6 12 6zm0 5c-.83 0-1.5-.67-1.5-1.5S11.17 8 12 8s1.5.67 1.5 1.5S12.83 11 12 11z"/></svg>
    <p><?=`$_SESSION['name']?></p>
  </div>
  <div>
  <img id="sn_logo" src="./sn_logo.png" alt="sn_logo">
  </div>
  <div>
  <a id="arch" href="./archive_page/.index.php">
    <p>Archiwum</p>
    <svg xmlns="http://www.w3.org/2000/svg" enable-background="new 0 0 24 24" viewBox="0 0 24 24" fill="#FFFFFF"><g><rect fill="none" height="24" width="24"/></g><g><g><path d="M20,2H4C3,2,2,2.9,2,4v3.01C2,7.73,2.43,8.35,3,8.7V20c0,1.1,1.1,2,2,2h14c0.9,0,2-0.9,2-2V8.7c0.57-0.35,1-0.97,1-1.69V4 C22,2.9,21,2,20,2z M19,20H5V9h14V20z M20,7H4V4h16V7z"/><rect height="2" width="6" x="9" y="12"/></g></g></svg>      </a>
  <a href="./logout.php">
    <p>Wyloguj</p>
    <svg xmlns="http://www.w3.org/2000/svg" enable-background="new 0 0 24 24" viewBox="0 0 24 24" fill="#FFFFFF"><g><path d="M0,0h24v24H0V0z" fill="none"/></g><g><path d="M17,8l-1.41,1.41L17.17,11H9v2h8.17l-1.58,1.58L17,16l4-4L17,8z M5,5h7V3H5C3.9,3,3,3.9,3,5v14c0,1.1,0.9,2,2,2h7v-2H5V5z"/></g></svg>
  </a>
  </div>
</nav>
    <p id="p_nav"></p>
  <div class="grid-container">
"@

$ending = @"
  </div>
  <center>
    <p>
      Ostatnia aktualizacja</br>
      $(get-date -Format "dd MMMM yyyy HH:mm:ss")
    </p>
    <p>
      <img id="kk-logo"
        src="./konkowit_logo.png" 
        alt="konkowit_logo">
      <br>
      Made by KonkowIT &#169;
    </p>
  </center>
  <script>
    function getMobileOperatingSystem() {
      var userAgent = navigator.userAgent || navigator.vendor || window.opera;
      if (userAgent.match(/iPad/i) || userAgent.match(/iPhone/i) || userAgent.match(/iPod/i)) {
        document.getElementsByTagName('body')[0].className += ' ios';
        return 'iOS';
      }
      else if (userAgent.match(/Android/i)) {
        document.getElementsByTagName('body')[0].className += ' android';
        return 'Android';
      }
      else {
        document.getElementsByTagName('body')[0].className += ' unknown';
        return 'unknown';
      }
    }

    var deviceType = getMobileOperatingSystem();
    document.getElementById("device").innerHTML = deviceType;
  </script>
  <script>
    var h = document.documentElement,
      b = document.body,
      st = 'scrollTop',
      sh = 'scrollHeight',
      progress = document.querySelector('.progress'),
      scroll;

    document.addEventListener('scroll', function () {
      scroll = (h[st] || b[st]) / ((h[sh] || b[sh]) - h.clientHeight) * 100;
      progress.style.setProperty('--scroll', scroll + '%');
    });
  </script>
</body>

</html>
"@

class FileObject {
  [string]$File
  [string]$DestinationPath
}

function GetComputersFromAPI {
  [CmdletBinding()]
  param(
    [parameter(ValueFromPipeline)]
    [ValidateNotNull()]
    [String]$networkName,
    [Array]$dontCheck
  )
    
  # Body
  $requestBody = @"
{

"network": [$($networkName)]

}
"@

  # Request
  try {
    $request = Invoke-WebRequest -Uri $requestURL -Method POST -Body $requestBody -Headers $requestHeaders -ea Stop
  }
  catch [exception] {
    $Error[0]
    Exit 1
  }

  # Creating PS array of sn
  if ($request.StatusCode -eq 200) {
    $requestContent = $request.content | ConvertFrom-Json
  }
  else {
    Write-host ( -join ("Received bad StatusCode for request: ", $request.StatusCode, " - ", $request.StatusDescription)) -ForegroundColor Red
    Exit 1
  }

  $snList = @()
  $requestContent | ForEach-Object {
    if ((!($dontCheck -match $_.name)) -and ($_.lok -ne "LOK0014")) {
      $hash = [ordered]@{
        sn       = $_.name;
        ip       = $_.ip;
        lok_id   = $_.lok
        placowka = $_.lok_name.toString();
        sim      = "NULL";
      }

      $snList = [array]$snList + (New-Object psobject -Property $hash)
    }
  }

  return $snList
}

function Start-SleepTimer($seconds) {
  $doneDT = (Get-Date).AddSeconds($seconds)
  while ($doneDT -gt (Get-Date)) {
    $secondsLeft = $doneDT.Subtract((Get-Date)).TotalSeconds
    $percent = ($seconds - $secondsLeft) / $seconds * 100
    Write-Progress -activity "LED Cameras screenshots" -Status "Ponowne uruchomienie skryptu za" -SecondsRemaining $secondsLeft -PercentComplete $percent
    [System.Threading.Thread]::Sleep(500)
  } 
  Write-Progress -activity "LED Cameras screenshots" -Status "Ponowne uruchomienie skryptu za" -SecondsRemaining 0 -Completed
}

if (!(Test-Path "C:\Program Files (x86)\WinSCP\WinSCPnet.dll")) {
  "`n"
  Write-Host "Brak zainstalowanego programu WinSCP! Zainstaluj, a nastepnie ponow uruchomienie skryptu" -ForegroundColor red -BackgroundColor Black
  "`n"
  Start-Sleep -s 3

  $Browser = new-object -com internetexplorer.application
  $Browser.navigate2("https://winscp.net/eng/downloads.php#additional")
  $Browser.visible = $true
  break
}

Get-Content "C:\Metadane_do_skryptow\sn_disabled_LedCameras.txt" -ErrorAction SilentlyContinue | ForEach-Object { 
  if ($dontCheck -notcontains $_) {
    $dontCheck = [array]$dontCheck + $_
  }
}

while (!(Test-Connection -ComputerName "10.99.99.10" -Count 3 -Quiet)) {
  Write-Host "VPN not connected!" -ForegroundColor Red
  Start-Sleep -Seconds 30
}

# Delete previosuly used files
if (test-path $ssDir) { remove-item $ssDir -Force -Recurse }
if (test-path $phpPage) { remove-item $phpPage -Force }
if (test-path $htmlSiteNew) { remove-item $htmlSiteNew -Force }

$freshData = GetComputersFromAPI -networkName '"LED City", "LED Premium", "Super Screen", "Pakiet Tranzyt"' -dontCheck $dontCheck
$csvContent = @()

foreach ($c in $csv) {
  $csvContent = $csvContent + (Import-Csv -Path $c -Delimiter ';' | Select-Object 'ID', 'Restart_SMS')
}

if (Test-Path $jsonPath) { 
  try { 
    [System.Collections.ArrayList]$localData = ConvertFrom-Json (Get-Content $jsonPath -Raw -ea Continue) -ea Continue 
    $runUpdate = $true
  }
  catch {
    Write-Host "ERROR: $($_.Exception.message)" 
  }
}
else {
  ConvertTo-Json -InputObject $freshData | Out-File $jsonPath
  $runUpdate = $false
}

if ($runUpdate) {
  foreach ($f in $freshData) {
    $counter = 0
    $ldCount = $localData.Count
    For ($i = 0; $i -lt $ldCount; $i++) {
      if ($f.sn -eq $localData[$i].sn) {
        # IP update
        if (($f.ip -ne $localData[$i].ip) -and ($f.ip -ne "NULL") -and ($f.ip -ne "")) {
          $localData[$i].ip = $f.ip
        }

        # LOK_ID update
        if (($f.lok_id -ne $localData[$i].lok_id) -and ($f.lok_id -ne "null") -and ($f.lok_id -ne "")) {
          $localData[$i].lok_id = $f.lok_id
        }

        # PLACOWKA update
        if (($f.placowka -ne $localData[$i].placowka) -and ($f.placowka -ne "null") -and ($f.placowka -ne "")) {
          $localData[$i].placowka = $f.placowka
        }
      }
      else {
        $counter++
        if ($counter -eq $ldCount) {
          # ADD NEW SN
          $hash = [ordered]@{
            sn              = $f.sn;
            ip              = $f.ip;
            lok_id          = $f.lok_id;
            placowka        = $f.placowka;
            sim             = "NULL";
          }
          
          Write-host "Adding $($f.sn) to json"
          $localData = [array]$localData + (New-Object psobject -Property $hash)
        }
      }
    }
  }

  for ($l = 0; $l -lt $localData.count; $l++) {
    $n = $localData[$l].sn
  
    # REMOVE MISSING SN
    if (!($freshData.sn -contains $n)) {
      Write-host "Removing $n from json"
      $localData.Remove($localData[$l])
    }

    # SIM NUMBER UPDATE
    if ($csvContent.ID -contains $localData[$l].lok_id) {
      $simNumber = ($csvContent | Where-Object { $_.ID -eq $localData[$l].lok_id }).Restart_SMS
      
      if (($simNumber -eq "N/A") -or ($simNumber -eq "") -or ($null -eq $simNumber)) {
        $localData[$l].sim = "null"
      }
      else {
        $localData[$l].sim = $simNumber
      }
    }
  }

  ConvertTo-Json -InputObject $localData | Out-File $jsonPath -Force
}

[System.Collections.ArrayList]$servers = ConvertFrom-Json (Get-Content $jsonPath -Raw -ea Continue) -ea Continue

foreach ($led in $servers) {
  $snIP = $led.ip
  $sn = $led.sn
  $snLoc = $led.placowka
  $ssDirDownloaded = (New-Item -ItemType Directory -Path ( -join ($ssDir, "`\", $sn)) -Force).FullName

  $scriptBlock = {
    $sn = $args[0]
    $ip = $args[1]
    $snLoc = $args[2]
    $ssDirDownloaded = $args[3]

    if ($ip -ne "null") {
      try { 
        Add-Type -Path "$(${env:ProgramFiles(x86)})\WinSCP\WinSCPnet.dll" -ErrorAction Stop

        $sessionOptions = New-Object WinSCP.SessionOptions -Property @{
          Protocol   = [WinSCP.Protocol]::ftp
          HostName   = $ip
          PortNumber = 0
          UserName   = "***"
          Password   = "***"
        }

        $session = New-Object WinSCP.Session

        try {
          $session.Open($sessionOptions)
          $transferOptions = New-Object WinSCP.TransferOptions
          $transferOptions.TransferMode = [WinSCP.TransferMode]::Binary
          $transferOptions.OverwriteMode.Overwrite
        
          $fileList = ($session.ListDirectory("/screennetwork/player")).Files
          $fileList | Where-Object { $_.Name -eq "screenshot_lan_camera.jpg" } | ForEach-Object {
            if (!(Test-path $_.FullName)) {
              Write-host "`nBrak zdjecia z kamery: $sn - $snLoc" -ForegroundColor Red
              copy-item "C:\SN_Scripts\LedCameras\errorNew.jpg" -Destination "$ssDirDownloaded\screenshot_lan_camera.jpg" -Force 
            }
            else {
              if ($_.LastWriteTime -gt ((get-date).AddMinutes(-30))) {
                Write-host "`nDownloading screenshot from:"$sn" - "$snLoc
                $result = $session.GetFiles("/screennetwork/player/screenshot_lan_camera.jpg", "$ssDirDownloaded\screenshot_lan_camera.jpg", $False, $transferOptions)
                Write-host "Successfully completed: $($result.IsSuccess)"
              }
              else {
                Write-host "`nBrak nowego zdjecia z kamery: $sn - $snLoc" -ForegroundColor Red
                copy-item "C:\SN_Scripts\LedCameras\errorNew.jpg" -Destination "$ssDirDownloaded\screenshot_lan_camera.jpg" -Force 
              }
            }
          }        
        }
        catch {
          $eMsg = ( -join ("`n", $_.Exception.Message, "`n`nLine ", $error[0].InvocationInfo.ScriptLineNumber, " : " + ($error[0].InvocationInfo.Line | Out-String).Trim() ))
          Write-Host "Wystapil blad pobierania zrzutu ekranu" -ForegroundColor Red -BackgroundColor Black
          Write-Host $eMsg
          copy-item "C:\SN_Scripts\LedCameras\errorDownload.jpg" -Destination "$ssDirDownloaded\screenshot_lan_camera.jpg" -Force 
        }
        finally {
          $session.Dispose()
        }
      }
      catch {
        Write-host "`nBlad polaczenia z komputerem: $sn - $snLoc"  -ForegroundColor Red
        copy-item "C:\SN_Scripts\LedCameras\errorConnection.jpg" -Destination "$ssDirDownloaded\screenshot_lan_camera.jpg" -Force
      }
    }
    else {
      Write-host "`nBlad polaczenia z komputerem: $sn - $snLoc"  -ForegroundColor Red
      copy-item "C:\SN_Scripts\LedCameras\errorConnection.jpg" -Destination "$ssDirDownloaded\screenshot_lan_camera.jpg" -Force
    }
  }

  Start-Job -ScriptBlock $scriptBlock -Name $sn -ArgumentList $sn, $snIP, $snLoc, $ssDirDownloaded
}

# Wait all jobs
Get-Job | Wait-Job

# Receive all jobs
Get-Job | Receive-Job

# Remove all jobs
Get-Job | Remove-Job

# HTML
# Add begginning
Add-Content $htmlSiteNew -Value $begginning
# Add ss
Get-ChildItem $ssDir -Recurse -Include "*.jpg" | ForEach-Object { 
  # Add ss - img and text
  $imgSN = $_.Directory.name
  $imgLokalizacja = ($servers | Where-Object { $_.sn -match $imgSN }).placowka.toString()
  $contentSN = @"
        <a href="#$($imgSN)">
          <img src="./ss/$($imgSN)/screenshot_lan_camera.jpg"
              alt="$($imgSN)">
        </a>
        <a href="#_" class="lightbox" id="$($imgSN)">
          <img src="./ss/$($imgSN)/screenshot_lan_camera.jpg"
              alt="$($imgSN)">
        </a>
        <div class="grid-item-text">
          $($imgSN) - $($imgLokalizacja)
        </div>
"@
  # Add ss - sim
  $simNumber = ($servers | Where-Object { $_.sn -eq $imgSN }).sim

  if ($simNumber -ne "NULL") {
    $contentSMS = @"
        <div class="btn-set">
          <button class="button txt-ios btn-style"
            onclick="window.location.href = 'sms://$($simNumber)/&body=1234:off=0,1,2,3,4';">
            OFF
          </button>
          <button class="button txt-ios btn-style"
            onclick="window.location.href = 'sms://$($simNumber)/&body=1234:on=0,1,2,3,4';">
            ON
          </button>
          <button class="button txt-ios btn-style"
            onclick="window.location.href = 'sms://$($simNumber)/&body=1234:reboot=0,1,2,3,4';">
            Restart
          </button>
          <button class="button txt-android btn-style"
            onclick="window.location.href = 'sms:$($simNumber)?body=1234:off=0,1,2,3,4';">
            OFF
          </button>
          <button class="button txt-android btn-style"
            onclick="window.location.href = 'sms:$($simNumber)?body=1234:on=0,1,2,3,4';">
            ON
          </button>
          <button class="button txt-android btn-style"
            onclick="window.location.href = 'sms:$($simNumber)?body=1234:reboot=0,1,2,3,4';">
            Restart
          </button>
        </div>
"@
  }

  # Adding
  Add-Content $htmlSiteNew -Value '      <div class="grid-container-cell">'
  Add-Content $htmlSiteNew -Value $contentSN -Encoding UTF8
  if ($contentSMS -ne $null) { Add-Content $htmlSiteNew -Value $contentSMS }
  Add-Content $htmlSiteNew -Value '      </div>'

  # Clean variables
  $contentSN = $null
  $contentSMS = $null
  $imgLokalizacja = $null
  $imgSN = $null
}

# Add ending
Add-Content $htmlSiteNew -Value $ending

# Create PHP
New-item -Path $homeDir -Name "index_logged.php" -ItemType file -Value $checkSession -force | out-null 
Add-Content "$homeDir\index_logged.php" -Value (Get-Content $htmlSiteNew) -Force

# Refresh html
$chromeProc = Get-Process chrome -ea SilentlyContinue
if ($null -ne $chromeProc) {
  $chromeProc | % { Stop-Process $_.Id -Force -ErrorAction SilentlyContinue }
}

Start-Process cmd.exe -WindowStyle Minimized -ArgumentList "/c",  "C:\SN_Scripts\LedCameras\RunChrome.bat"

$error.Clear()

# Upload to server
Add-Type -Path "$(${env:ProgramFiles(x86)})\WinSCP\WinSCPnet.dll" -ErrorAction Stop
$LastWriteTimeDate = Get-Date
$lcRemotePath = "/www/public/LedCameras"

$sessionOptions = New-Object WinSCP.SessionOptions -Property @{
  Protocol   = [WinSCP.Protocol]::ftp
  HostName   = "***"
  PortNumber = 0
  UserName   = "***"
  Password   = "***"
}

$session = New-Object WinSCP.Session
$session.Open($sessionOptions)
$transferOptions = New-Object WinSCP.TransferOptions
$transferOptions.TransferMode = [WinSCP.TransferMode]::Binary
$transferOptions.OverwriteMode.Overwrite
$transferOptions.FilePermissions = New-Object WinSCP.FilePermissions
$transferOptions.FilePermissions.Octal = "777"

$files = ( @(
  [FileObject]@{ File = "$homeDir\index_logged.php"; DestinationPath = "$lcRemotePath/" }
  Get-ChildItem "$homeDir\ss" -File -Recurse | ForEach-Object { 
      $dir = Split-Path $_.Directory -Leaf
      $toArchiveName = "$($_.Directory)\CameraScreenshot_$dir.jpg"
      $_.LastWriteTime = $LastWriteTimeDate
      Copy-Item $_.FullName -Destination $toArchiveName
      [FileObject]@{ File = "$homeDir\ss\$dir\screenshot_lan_camera.jpg"; DestinationPath = "$lcRemotePath/ss/$dir/" } 
      [FileObject]@{ File = $toArchiveName; DestinationPath = "$lcRemotePath/ss/$dir/" } 
  }
))

$fileList = ($session.ListDirectory("$lcRemotePath/ss")).Files
""
Foreach ($f in $files) {
  $par = Split-Path -leaf (Split-Path -Parent $f.file)
  if ((!($fileList.Name -contains $par)) -and $f.DestinationPath -ne "$lcRemotePath/") {
      Write-host "Creating directory for computer $par on server"
      $session.CreateDirectory("$lcRemotePath/ss/$par")
  }

  $r = $session.PutFiles($f.File, $f.DestinationPath, $False, $transferOptions)
  if ($r.IsSuccess -eq $true) {
      Write-host "Uploading '$($f.File)' to server completed successfully"
  }
  else {
      Write-host "Uploading '$($f.File)' to '$($f.DestinationPath)' failed: $($error[0].exception.message)"
  }
}

$session.Dispose()

# SSH
$qbiviewIP = "***"
$username = "***"
$secpasswd = ConvertTo-SecureString "***" -AsPlainText -Force
$credential = new-object -typename System.Management.Automation.PSCredential -argumentlist $username, $secpasswd

try {
    # Create ssh connection
    New-SSHSession -ComputerName $qbiviewIP -Credential $credential -ConnectionTimeout 300 -force -ErrorAction Stop -WarningAction silentlyContinue | out-null
    $getSSHSessionId = (Get-SSHSession | Where-Object { $_.Host -eq $qbiviewIP }).SessionId
}
catch {
    Write-Host "Error while connecting to SSH server: $($_.exception.message)"
}

if ($null -ne $getSSHSessionId) {
    # Invoke command
    (Invoke-SSHCommand -SessionId $getSSHSessionId -Command "chmod --changes -R 777 www/public/LedCameras/").output[0]

    # Terminate connection
    Write-Host ( -join ("Closing SSH connection: ", (Remove-SSHSession -SessionId $getSSHSessionId)))
}

# Restart script
Start-SleepTimer 180
$arguments = "& '$homeDir\LEDCameras_API.ps1'"
Start-Process powershell -ArgumentList $arguments -WindowStyle Minimized
Exit
