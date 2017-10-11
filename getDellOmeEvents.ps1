Function Send-EMail {
    # Use as follow:
    # Send-EMail -EmailTo "Myself@gmail.com" -Body "Test Body" -Subject "Test Subject" 
    Param (
        [Parameter(`
            Mandatory=$true)]
        [String[]]$EmailTo, #This gives a default value to the $EmailFrom command
        [Parameter(`
            Mandatory=$true)]
        [String]$Subject,
        #[Parameter(`
        #    Mandatory=$true)]
        #[String]$Body,
        [Parameter(`
            Mandatory=$false)]
        [String]$EmailFrom="dell.ome@cba.com.au",  #This gives a default value to the $EmailFrom command
        [Parameter(`
            mandatory=$false)]
        [String]$Attachment,
        [Parameter(`
            mandatory=$false)]
        [String]$Password="<<INSERT PASSWORD HERE>>",
        [Parameter(`
            mandatory=$false)]
        [String]$Username="acoe_oc_ome_smtp"
    )

        $SMTPServer = "securemailrelay.cba.com.au" 
        $SMTPMessage = New-Object System.Net.Mail.MailMessage($EmailFrom,$EmailTo,$Subject,$messageBody)
        $SMTPMessage.IsBodyHtml = $true
        $SMTPMessage.Body| ConvertTo-Html

        if ($Attachment -ne $null) {
            $SMTPattachment = New-Object System.Net.Mail.Attachment($Attachment)
            $SMTPMessage.Attachments.Add($SMTPattachment)
        }

        $SMTPClient = New-Object Net.Mail.SmtpClient($SmtpServer, 25) 
        $SMTPClient.EnableSsl = $false 
        $SMTPClient.Credentials = New-Object System.Net.NetworkCredential($Username, $Password);
        $SMTPClient.Send($SMTPMessage)
        Remove-Variable -Name SMTPClient
        Remove-Variable -Name Password

} #End Function Send-EMail

function Ignore-SelfSignedCerts
 {
     try
     {

        Write-Host "Adding TrustAllCertsPolicy type." -ForegroundColor White
         Add-Type -TypeDefinition  @"
         using System.Net;
         using System.Security.Cryptography.X509Certificates;
         public class TrustAllCertsPolicy : ICertificatePolicy
         {
              public bool CheckValidationResult(
              ServicePoint srvPoint, X509Certificate certificate,
              WebRequest request, int certificateProblem)
              {
                  return true;
             }
         }
"@
         
         Write-Host "TrustAllCertsPolicy type added." -ForegroundColor White
       }
     catch
     {
         Write-Host $_ -ForegroundColor "Yellow"
     }

     [System.Net.ServicePointManager]::CertificatePolicy = New-Object TrustAllCertsPolicy
 }

function getAlertFiltersCriticalWarning($baseuri) {

    # Get all of the Alert fitler Ids so that you can select only the Critical and Warning alerts
   
    $filterXml = Invoke-RestMethod -Method Get -UseDefaultCredentials -Uri "$baseuri/AlertFilters"
    # DEBUG COMMENT: print the Alert filters
    # $filterXml.GetAllAlertFiltersResponse.GetAllAlertFiltersResult.AlertFilter | Format-Table -Property Id, IsEnabled, IsReadOnly, Name, Type

    # Store the Critical and Warning alert filters
   
    $criticalFilter = $filterXml.GetAllAlertFiltersResponse.GetAllAlertFiltersResult.AlertFilter|where {$_.name -contains "Critical Alerts"}
    $script:criticalFilterId = $criticalFilter.Id
    $warningFilter = $filterXml.GetAllAlertFiltersResponse.GetAllAlertFiltersResult.AlertFilter|where {$_.name -contains "Warning Alerts"}
    $script:warningFilterId = $warningFilter.Id
}

function get24hCriticalAlerts($baseuri) {
    
    # Now to get all the Critical Alerts and display them in a table
   
    $script:criticalAlertsXml = Invoke-RestMethod -Method Get -UseDefaultCredentials -Uri "$baseuri/AlertFilters/$criticalFilterId/Alerts"
    #$criticalAlertsXml.GetAlertsForFilterResponse.GetAlertsForFilterResult.Alert | Format-Table -Property DeviceIdentifier, DeviceName, DeviceNodeId, DeviceServiceTag, DeviceSystemModelType, DeviceType, DeviceTypeName, EventCategory, EventSource, Id, IsIdrac, IsInband, Message, OSName, Package, SNMPEnterpriseOID, SNMPGenericTrapID, SNMPSpecificTrapID, Severity, SourceName, Status, Time 

    # Initialise the empty array
    $criticalAlerts24h = @()

    foreach ($alert in $criticalAlertsXml.GetAlertsForFilterResponse.GetAlertsForFilterResult.Alert) {
        #(Get-Date -Format s $timeTwin) -gt (Get-Date -Format s ((Get-Date).AddHours(-24)))
        $criticalAlerts24h += $alert | where { (Get-Date -Format s $alert.Time) -gt (Get-Date -Format s ((Get-Date).AddHours(-24))) }
    }

    $criticalAlerts24h

    # DEBUG COMMENT: print all of the Last 24 Hours Critical Alerts
    # $criticalAlerts24h | Format-Table -Property Time, DeviceIdentifier, DeviceName, DeviceNodeId, DeviceServiceTag, DeviceSystemModelType, DeviceType, DeviceTypeName, EventCategory, EventSource, Id, IsIdrac, IsInband, Message, OSName, Package, SNMPEnterpriseOID, SNMPGenericTrapID, SNMPSpecificTrapID, Severity, SourceName, Status

}

function get24hWarningAlerts($baseuri) {
    
    # Now to get all the Critical Alerts and display them in a table
    $script:warningAlertsXml = Invoke-RestMethod -Method Get -UseDefaultCredentials -Uri "$baseuri/AlertFilters/$warningFilterId/Alerts"
    #$warningAlertsXml.GetAlertsForFilterResponse.GetAlertsForFilterResult.Alert | Format-Table -Property DeviceIdentifier, DeviceName, DeviceNodeId, DeviceServiceTag, DeviceSystemModelType, DeviceType, DeviceTypeName, EventCategory, EventSource, Id, IsIdrac, IsInband, Message, OSName, Package, SNMPEnterpriseOID, SNMPGenericTrapID, SNMPSpecificTrapID, Severity, SourceName, Status, Time 

    # Initialise the empty array
    $warningAlerts24h = @()

    foreach ($alert in $warningAlertsXml.GetAlertsForFilterResponse.GetAlertsForFilterResult.Alert) {
        # (Get-Date -Format s $timeTwin) -gt (Get-Date -Format s ((Get-Date).AddHours(-24)))
        $warningAlerts24h += $alert | where { (Get-Date -Format s $alert.Time) -gt (Get-Date -Format s ((Get-Date).AddHours(-24))) }
    }

    $warningAlerts24h

    # DEBUG COMMENT: print all of the Last 24 Hours Warning Alert
    # $warningAlerts24h | Format-Table -Property Time, DeviceIdentifier, DeviceName, DeviceNodeId, DeviceServiceTag, DeviceSystemModelType, DeviceType, DeviceTypeName, EventCategory, EventSource, Id, IsIdrac, IsInband, Message, OSName, Package, SNMPEnterpriseOID, SNMPGenericTrapID, SNMPSpecificTrapID, Severity, SourceName, Status

}

# ------------------ MAIN SECTION --------------------------
#
# Set up some global variables needed for Dell OME and authentication
    # Set up the username and password for Basic Authentication (if needed)
    # In this case, the Dell OME authentication is locally set to the user running it (which migh need to change)
    # $user = "PAM01-PRD01\username"
    # $pass = "password"
    # $pair = "${user}:${pass}
    # $bytes = [System.Text.Encoding]::ASCII.GetBytes($pair)
    # $base64 = [System.Convert]::ToBase64String($bytes)
    # $basicAuthValue = "Basic $base64"
    # $script:headers = @{ Authorization = $basicAuthValue }

$baseuri = "https://ocmome01.pam01-prd01.iams.cba:2607/api/OME.svc"

# Set up the ignore self-signed certificates

Ignore-SelfSignedCerts

# Get all of the Alert Filter Id so that we can select only specific Alerts

getAlertFiltersCriticalWarning($baseuri)

# Initialise the empty arrays

$criticalAlerts24h = @()
$warningAlerts24h = @()

# Get all the Critical Alerts

$criticalAlerts24h = get24hCriticalAlerts($baseuri)

# Get all the Warning Alerts

$warningAlerts24h = get24hWarningAlerts($baseuri)

# Send email with Critical Alerts

if ($criticalAlerts24h -ne $null) {
    $messageBodyTop = $criticalAlerts24h | Group-Object -Property DeviceName, EventSource -NoElement | Sort-Object -Descending count | Select-Object Count, Name | ConvertTo-Html
    $messageBodyBottom = $criticalAlerts24h | Select-Object Time, DeviceName, SourceName, Severity, StatusDeviceTypeName, EventCategory, EventSource, Message, OSName, Package, Status | ConvertTo-Html
    $script:messageBody = $messageBodyTop + $messageBodyBottom
} else {
    $script:messageBody = "Boom! No Critical Alerts for the last 24 Hours!"
}

$criticalAlerts24h | Export-Csv "C:\Users\wonigkwr-adm\last24CriticalAlerts.csv" -NoTypeInformation

Send-EMail -EmailTo "storms.cba.privatecloud@emc.com" -Subject "Critical Alerts for Last 24 Hours - Received from Dell OpenManage Essentials" -Attachment "C:\Users\wonigkwr-adm\last24CriticalAlerts.csv"

# Send email with Warning Alerts

if ($warningAlerts24h -ne $null) {
    $messageBodyTop = $warningAlerts24h | Group-Object -Property DeviceName, EventSource -NoElement | Sort-Object -Descending count | Select-Object Count, Name | ConvertTo-Html
    $messageBodyBottom = $warningAlerts24h | Select-Object Time, DeviceName, SourceName, Severity, StatusDeviceTypeName, EventCategory, EventSource, Message, OSName, Package, Status | ConvertTo-Html
    $script:messageBody = $messageBodyTop + $messageBodyBottom
} else {
    $script:messageBody = "Boom! No Warning Alerts for the last 24 Hours!"
}

$warningAlerts24h | Export-Csv "C:\Users\wonigkwr-adm\last24WarningAlerts.csv" -NoTypeInformation

Send-EMail -EmailTo "storms.cba.privatecloud@emc.com" -Subject "Warning Alerts for Last 24 Hours - Received from Dell OpenManage Essentials" -Attachment "C:\Users\wonigkwr-adm\last24WarningAlerts.csv"