set-executionpolicy bypass -force
import-module webadministration;
set-location C:\windows\system32

$thumbprint = ''
$pfxPassword='';
$pfxPath = ""
$domain = ""

# Find number of bad certs
$OriginalCert = (gci -path cert:\ -recurse | ? { $_.Subject -like "CN=*.$domain*" -and $_.Thumbprint -ne "$thumbprint" } | Measure-Object).Count


function Import-PfxCertificate { 

    param([String]$certPath,[String]$certRootStore = 'localmachine',[String]$certStore = 'My',$pfxPass = $null) 
    $pfx = new-object System.Security.Cryptography.X509Certificates.X509Certificate2 

    if ($pfxPass -eq $null) 
    {
        $pfxPass = read-host "Password" -assecurestring
    } 

    $pfx.import($certPath,$pfxPass,"Exportable,PersistKeySet") 

    $store = new-object System.Security.Cryptography.X509Certificates.X509Store($certStore,$certRootStore) 
    $store.open("MaxAllowed") 
    $store.add($pfx) 
    $store.close() 
}




if($OriginalCert)
{
	# found an outdated cert
	Write-Output "Old certificate detected! Deleting now!";
	certutil.exe -delstore my "*.$domain.com"
		
}

Write-Output "Importing new cert";

Import-PfxCertificate -certPath $pfxpath -certStore "My" -pfxPass $pfxpassword


remove-webbinding;
dir iis:\sslbindings\ | ? { $_.Port -eq 443 } | remove-item -Force;

New-WebBinding -Name 'Default Web Site' -IPAddress '*' -Port 80 -Protocol http;
New-WebBinding -name "Default Web Site" -IP "*" -Port 443 -protocol https;


#move-item Cert:\LocalMachine\My\$thumbprint cert:\LocalMachine\WebHosting;
Get-ChildItem cert:\LocalMachine\My\$thumbprint | select -First 1 | New-Item IIS:\SslBindings\0.0.0.0!443;

if(dir iis:\SSLBindings\ | ? { $_.Port -eq 443 })
{
	Write-Output "SSL binding was default website was successfully made. Proceeding to attach cert!";
}
else
{
	Write-Output "[CRITICAL ERROR] 443 PORT BINDING MISSING AFTER ATTEMPTING TO ADD!!!";
	exit;
}

<#
$certificate = Get-ChildItem cert:\ -Recurse | Where-Object { $_.Subject -like 'CN=*.hostedrmm*'} | Select-Object -First 1;
$AddSSLCertToWebBinding = (Get-WebBinding 'Default Web Site' -Port $443 -Protocol "https").AddSslCertificate($certificate.thumbprint, 'MY'); " 
& net stop iisadmin&net start iisadmin&net stop w3svc&net start w3svc
#>

Write-Output "Proceeding to restart services!"

Restart-Service IISAdmin -Force;
Restart-service w3svc -FORCE;

$NewCertDetected = (gci -path cert:\ -recurse | ? { $_.Subject -like 'CN=*.hostedrmm*' -and $_.Thumbprint -eq "$thumbprint" } | Measure-Object).Count

if($NewCertDetected -ge 1)
{
	Write-Output "[COMPLETE] Successfully bound SSL cert!"
}

