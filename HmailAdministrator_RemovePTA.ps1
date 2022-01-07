# Create COM Object

$MailObj = New-Object -ComObject HMailServer.Application;
$MailObj.Authenticate("Administrator","$hmailpass")

$MailObj.Connect()



#Add domain to HMail and make active w/ friendly name

$settings = $MailObj.settings

$settings.AllowSMTPAuthPlain = $false

$settings.ServiceSMTP = $true

$settings.ServiceIMAP = $false



<#
$count= $MailObj.Settings.SecurityRanges.Count


$i = 0


do {
$Range = $MailObj.settings.SecurityRanges.Item($i)


if($Range.lowerip -ne '127.0.0.1') {
$Range.RequireSSLTLSForAuth = $true}


$range.AllowIMAPConnections = $false

$range.AllowSMTPConnections = $true

#$Range.RequireAuthForDeliveryToLocal = $true

#$Range.RequireAuthForDeliveryToRemote = $false

$Range.Save()

$i+=1
}

until ($i -eq $count)


#>

Restart-Service "hMailServer"