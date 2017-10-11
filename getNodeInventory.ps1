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
        [Parameter(`
            Mandatory=$true)]
        [Object[]]$Body,
        [Parameter(`
            Mandatory=$false)]
        [String]$EmailFrom="dell.ome@cba.com.au",  #This gives a default value to the $EmailFrom command
        [Parameter(`
            mandatory=$false)]
        [String[]]$Attachment,
        [Parameter(`
            mandatory=$false)]
        [String]$Password="<<INSERT PASSWORD HERE>>",
        [Parameter(`
            mandatory=$false)]
        [String]$Username="acoe_oc_ome_smtp"
    )

        $SMTPServer = "securemailrelay.cba.com.au" 
        $SMTPMessage = New-Object System.Net.Mail.MailMessage($EmailFrom,$EmailTo,$Subject,$Body)
        $SMTPMessage.IsBodyHtml = $true
        # $SMTPMessage.Body| ConvertTo-Html
        if ($Attachment -ne $null) {
            foreach ($Attach in $Attachment) {
                $SMTPattachment = New-Object System.Net.Mail.Attachment($Attach)
                $SMTPMessage.Attachments.Add($SMTPattachment)
            }
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

function getDeviceGroups($baseuri) {
    # Get Device Groups
    $filterXml = Invoke-RestMethod -Method Get -UseDefaultCredentials -Uri "$baseuri/DeviceGroups"
    # DEBUG COMMENT: print the Device Groups
    # $filterXml.GetDeviceGroupsResponse.GetDeviceGroupsResult.DeviceGroup | Format-Table -Property Description, DeviceCount, Id, Name, RollupHealth, Type
    
    # Store the Id for the DeviceGroups which have device allocated to them
    $deviceGroupArray = $filterXml.GetDeviceGroupsResponse.GetDeviceGroupsResult.DeviceGroup|where {$_.DeviceCount -gt 0}
    #$deviceGroupArray | Format-Table -Property Description, DeviceCount, Id, Name, RollupHealth, Type

    $deviceGroupArray
    #$script:criticalFilterId = $criticalFilter.Id
    #$warningFilter = $filterXml.GetAllAlertFiltersResponse.GetAllAlertFiltersResult.AlertFilter|where {$_.name -contains "Warning Alerts"}
    #$script:warningFilterId = $warningFilter.Id
}

function getAllDevicesInGroup($groupId, $baseuri) {
    # Get Device Groups
    $filterXml = Invoke-RestMethod -Method Get -UseDefaultCredentials -Uri "$baseuri/DeviceGroups/$groupId/Devices"
    # DEBUG COMMENT: print the Device Groups
    # $filterXml.GetDeviceGroupsResponse.GetDeviceGroupsResult.DeviceGroup | Format-Table -Property Description, DeviceCount, Id, Name, RollupHealth, Type
    
    # Store the Id for the DeviceGroups which have device allocated to them
    $deviceGroupSummaryArray = $filterXml.GetDevicesResponse.GetDevicesResult.Device
    $deviceGroupSummaryArray
}

function getDeviceProcessor($deviceId, $baseuri) {
    # Get device inventory
    $filterXml = Invoke-RestMethod -Method Get -UseDefaultCredentials -Uri "$baseuri/Devices/$deviceId/Processor"
    # DEBUG COMMENT: print the Device Groups
    # $filterXml.GetDeviceGroupsResponse.GetDeviceGroupsResult.DeviceGroup | Format-Table -Property Description, DeviceCount, Id, Name, RollupHealth, Type
    
    # Store the Id for the DeviceGroups which have device allocated to them
    $deviceProcessor = $filterXml.DeviceProcessorResponse.DeviceProcessorResult.Processor
    $deviceProcessor
}

function getDeviceMemory($deviceId, $baseuri) {
    # Get device memory information
    $filterXml = Invoke-RestMethod -Method Get -UseDefaultCredentials -Uri "$baseuri/Devices/$deviceId/Memory"
    # DEBUG COMMENT: print the Device Groups
    # $filterXml.GetDeviceGroupsResponse.GetDeviceGroupsResult.DeviceGroup | Format-Table -Property Description, DeviceCount, Id, Name, RollupHealth, Type
    
    # Store the Id for the DeviceGroups which have device allocated to them
    $deviceMemory = $filterXml.DeviceMemoryResponse.DeviceMemoryResult.TotalMemory
    $deviceMemory
}

# Set up some global variables needed for Dell OME and authentication
# Set up the username and password for Basic Authentication (if needed)
# In this case, the Dell OME authentication is locally set to the user running it (which migh need to change)
# $user = "PAM01-PRD01\username"
# $pass = "password"
# $pair = "${user}:${pass}s
# $bytes = [System.Text.Encoding]::ASCII.GetBytes($pair)
# $base64 = [System.Convert]::ToBase64String($bytes)
# $basicAuthValue = "Basic $base64"
# $script:headers = @{ Authorization = $basicAuthValue }

$baseuri = "https://ocmome01.pam01-prd01.iams.cba:2607/api/OME.svc"
        
# Set up the ignore self-signed certificates

Ignore-SelfSignedCerts

# Set up the device object

$deviceArray =@()

# Get all of the the Device Groups so we can select "RAC" only

$groupObjects = getDeviceGroups $baseuri
$racGroupId = $groupObjects |where {$_.Name -eq "RAC"}

# Get all of the the Devices in the Device Group by the Id of the Device Group

$racAllDevices = getAllDevicesInGroup $racGroupId.Id $baseuri

foreach ($racAllDevice in $racAllDevices) {

    $deviceObject = New-Object -TypeName PSObject
    
    #Processor information
    $procObject = getDeviceProcessor $racAllDevice.Id $baseuri
    $deviceObject | Add-Member -Name 'Hostname' -MemberType NoteProperty -Value $racAllDevice.Name
    $deviceObject | Add-Member -Name 'Model' -MemberType NoteProperty -Value $racAllDevice.SystemModel
    $deviceObject | Add-Member -Name 'ServiceTag' -MemberType NoteProperty -Value $racAllDevice.ServiceTag
    $deviceObject | Add-Member -Name 'ExpressServiceCode' -MemberType NoteProperty -Value $racAllDevice.ExpressServiceCode
    $deviceObject | Add-Member -Name 'procBrand' -MemberType NoteProperty -Value $procObject.Brand[0]
    $deviceObject | Add-Member -Name 'procQuantity' -MemberType NoteProperty -Value $procObject.Count
    $deviceObject | Add-Member -Name 'procCores' -MemberType NoteProperty -Value $procObject.Cores[0]
    $deviceObject | Add-Member -Name 'procModel' -MemberType NoteProperty -Value $procObject.Model[0]
    
    #Memory information
    $memObject = getDeviceMemory $racAllDevice.Id $baseuri
    $deviceObject | Add-Member -Name 'memSize' -MemberType NoteProperty -Value $memObject

    $deviceArray += $deviceObject
}

# Set the counters for each type of Node (ineffective I know, we'll fix this later)

$countNodeTypeABurwood = 0
$countNodeTypeANorwest = 0
$countNodeTypeBBurwood = 0
$countNodeTypeBNorwest = 0
$countNodeTypeCBurwood = 0
$countNodeTypeCNorwest = 0

# Count the amount of Node Type A, B or C

foreach ($device in $deviceArray) {
    
    if (($device.Hostname -like "*ocb*") -And ($device.procCores -eq 22) -And ($device.memSize -eq 786432)) {
        $countNodeTypeABurwood += 1
    } elseif (($device.Hostname -like "*ocn*") -And ($device.procCores -eq 22) -And ($device.memSize -eq 786432)) {
        $countNodeTypeANorwest += 1
    } elseif (($device.Hostname -like "*ocb*") -And ($device.procCores -eq 22) -And ($device.memSize -eq 1572864)) {
        $countNodeTypeCBurwood += 1
    } elseif (($device.Hostname -like "*ocn*") -And ($device.procCores -eq 22) -And ($device.memSize -eq 1572864)) {
        $countNodeTypeCNorwest += 1
    } elseif (($device.Hostname -like "*ocb*") -And ($device.procCores -eq 6) -And ($device.memSize -eq 786432)) {
        $countNodeTypeBBurwood += 1
    } elseif (($device.Hostname -like "*ocn*") -And ($device.procCores -eq 6) -And ($device.memSize -eq 786432)) {
        $countNodeTypeBNorwest += 1
    } else {
        echo "Unknown Type"
    }
}

# Create Object with the values in for each Node Type

$nodeArray = @( @{NodeType = "A"; Datacentre="Burwood"; Count=$countNodeTypeABurwood},
                @{NodeType = "A"; Datacentre="Norwest"; Count=$countNodeTypeANorwest},
                @{NodeType = "B"; Datacentre="Burwood"; Count=$countNodeTypeBBurwood},
                @{NodeType = "B"; Datacentre="Norwest"; Count=$countNodeTypeBNorwest},
                @{NodeType = "C"; Datacentre="Burwood"; Count=$countNodeTypeCBurwood}, 
                @{NodeType = "C"; Datacentre="Norwest"; Count=$countNodeTypeCNorwest} )

# Create the headers for the HTML message

$Header = @"
<style>
TABLE {border-width: 1px;border-style: solid;border-color: black;border-collapse: collapse;}
TH {border-width: 1px;padding: 3px;border-style: solid;border-color: black;background-color: #6495ED; font-family:courier; font-size:12pt;}
TD {border-width: 1px;padding: 3px;border-style: solid;border-color: black;}
BODY {font-family:courier; font-size:10pt;}
</style>
<title>
Device Inventory - Breakdown per node type
</title>
"@

$Pre = @"
<div style='margin:  0px auto; BACKGROUND-COLOR:Black;Color:White;font-weight:bold;FONT-SIZE:  16pt;TEXT-ALIGN: center;'>
Node Type Breakdown
</div>
</p>   
"@

$Post = @"
</div>
</p>   
"@

$HTML = $nodeArray.ForEach({[PSCustomObject]$_}) | Select-Object -Property NodeType, Datacentre, Count | ConvertTo-Html -Head $Header -PreContent $Pre -PostContent $Post

$Pre = @"
<div style='margin:  0px auto; BACKGROUND-COLOR:Black;Color:White;font-weight:bold;FONT-SIZE:  16pt;TEXT-ALIGN: center;'>
Dell OpenManager Essentials - Physical Infrastructure Details
</div>
</p>   
"@

$Post = @"
<div style='margin:  0px auto; BACKGROUND-COLOR:Black;Color:White;font-weight:bold;FONT-SIZE:  16pt;TEXT-ALIGN: center;'>
</div>
</p>   
"@


$HTML2 = $deviceArray | Sort-Object -Property Hostname| Select-Object -Property Hostname, Model, ServiceTag, ExpressServiceCode, procBrand, procQuantity, procCores, procModel, memSize  | ConvertTo-Html -Head $Header -PreContent $Pre -PostContent $Post

$HTML = $HTML + $HTML2

# $HTML | Out-File NodeTypesInvenotry.html
# Invoke-Item NodeTypesInvenotry.html

# Create the CSV file for the Node Type counts

$nodeArray.ForEach({[PSCustomObject]$_}) | Select-Object -Property NodeType, Datacentre, Count | Export-Csv "C:\Users\wonigkwr-adm\nodeTypes.csv" -NoTypeInformation

$deviceArray | Sort-Object -Property Hostname| Select-Object -Property Hostname, Model, ServiceTag, ExpressServiceCode, procBrand, procQuantity, procCores, procModel, memSize | Export-Csv "C:\Users\wonigkwr-adm\nodeInformation.csv" -NoTypeInformation

# Send the email with the morning information

Send-EMail -EmailTo "storms.cba.privatecloud@emc.com" -Subject "Node Type Information" -Body $HTML -Attachment "C:\Users\wonigkwr-adm\nodeTypes.csv","C:\Users\wonigkwr-adm\nodeInformation.csv"
