New-SPWOPIBinding -ServerName <WacServerName> -AllowHTTP
Get-SPWOPIZone
Set-SPWOPIZone -zone "internal-http"
(Get-SPSecurityTokenServiceConfig).AllowOAuthOverHttp
$config = (Get-SPSecurityTokenServiceConfig)
$config.AllowOAuthOverHttp = $true
$config.Update()
(Get-SPSecurityTokenServiceConfig).AllowOAuthOverHttp

#EXCEL SOAP API
$Farm = Get-SPFarm 
$Farm.Properties.Add("WopiLegacySoapSupport", "<URL>/x/_vti_bin/ExcelServiceInternal.asmx"); 
$Farm.Update();
