$ErrorActionPreference = "silentlycontinue"

set-service wuauserv -StartupType Automatic
start-service wuauserv
set-service trustedinstaller -StartupType Automatic
start-service trustedinstaller

$manualguids = @( @guidlist@ )

$manualguids

$os = Get-CimInstance Win32_OperatingSystem | Select-Object -ExpandProperty Caption
$OperatingSystem = $os.Replace("Standard", "").Replace("Microsoft ", "").Replace(" Pro", "").Replace("Professional", "").Replace("Home", "").Replace("Enterprise", "").Replace("Datacenter", "").Trim()
$url = "https://www.catalog.update.microsoft.com/Search.aspx?q=$OperatingSystem"
$results = Invoke-WebRequest -Uri $url -UseBasicParsing
$kbids = $results.InputFields |Where-Object { $_.type -eq 'Button' -and $_.Value -eq 'Download' } | Select-Object -ExpandProperty ID
$resultlinks = $results.Links | Where-Object ID -match '_link' |  Where-Object { $_.OuterHTML -match ( "(?=.*" + ( $OperatingSystem -join ")(?=.*" ) + ")" ) }
$guids = @()
            foreach ($resultlink in $resultlinks) {
                $itemguid = $resultlink.id.replace('_link', '')
                $itemtitle = ($resultlink.outerHTML -replace '<[^>]+>', '').Trim()
                if ($itemguid -in $kbids) {
                    $guids += [pscustomobject]@{
                        Guid  = $itemguid
                        Title = $itemtitle
                    }
                }
            }

#$guids

foreach($id in $manualguids) {$guids+= @{Guid = $id; Title = $null}}

foreach ($item in $guids) {
                    $guid = $item.guid
                    $itemtitle = $item.Title
                    write-output $guid 
                    $kbname= $item.title.split("(")[1].split(")")[0]
                    if(!$kbname) {$kbname = "manualdownload"}
                    $checkhotfix = Get-HotFix -id $kbname
                    if(!$checkhotfix) {
                    Write-Verbose -Message "Downloading information for $itemtitle"
                    $post = @{ size = 0; updateID = $guid; uidInfo = $guid } | ConvertTo-Json -Compress
                    $body = @{ updateIDs = "[$post]" }
                    $downloaddialog=  Invoke-WebRequest -Uri 'https://www.catalog.update.microsoft.com/DownloadDialog.aspx' -Method Post -Body $body -UseBasicParsing | Select-Object -ExpandProperty Content 
                    $content = $downloaddialog -split "`n"|select-string "download.windowsupdate"
                    $content = $content -replace 'www.download.windowsupdate', 'download.windowsupdate'
                    $link = $content.split("'")[1].ToString()
                    write-output $link 
                    if($link -like "*windows10.0-kb*") {$separator = "windows10.0-"} 
                        elseif ($link -like "*windows-kb*") {$separator = "windows-"}
                        else {$separator = $null}
                    if($separator) {
                    $name = ($link -split $separator)[1]
                    write-output $name
                    $separator2 = "-"
                    $kbname =($name -split $separator2)[0]}
                        else {$name = "updatefile.exe"; $kbname = 'kb0000'; write-output $name}
                    Invoke-WebRequest -Uri $link -OutFile  "C:\windows\temp\$name" -UseBasicParsing
                    $updateprogress = Invoke-Command -ScriptBlock{&cmd /c "wusa.exe `"C:\windows\temp\$kbname.msu`" /quiet /log /norestart"}
                    if($LASTEXITCODE -in 0,2359302){write-output "Update Already Installed"; $rebootneeded = $true}
                    elseif($LASTEXITCODE -in 2,3010) {write-output "Update succeeded. Reboot Needed"; $rebootneeded = $true}
                    elseif($LASTEXITCODE -in 1630,-2145124329){write-output "Update not applicable to OS"}
                    else {Write-Output "KB $kbname finished with exit code $LASTEXITCODE"}
                    if($kbname -ne 'kb0000'){
                    $checkhotfix2 = Get-HotFix -id $kbname
                    if(!$checkhotfix2) {write-output "UPDATE FAILED TO INSTALL $kbname"}}
                    Clear-Variable kbname
                    remove-item "C:\windows\temp\$name"
                }
            }

if($rebootneeded -eq $true) {write-output "Server needs reboot to finish updates"}

Set-ExecutionPolicy RemoteSigned -Force
[Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12
Install-PackageProvider -Name NuGet -RequiredVersion 2.8.5.201 -Force
Install-Module -Name PSWindowsUpdate -Force
Import-Module -Name PSWindowsUpdate -force
Get-WindowsUpdate -AcceptAll -Install -IgnoreReboot