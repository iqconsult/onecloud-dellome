Function Set-CellColor
{   <#
    .SYNOPSIS
        Function that allows you to set individual cell colors in an HTML table
    .DESCRIPTION
        To be used inconjunction with ConvertTo-HTML this simple function allows you
        to set particular colors for cells in an HTML table.  You provide the criteria
        the script uses to make the determination if a cell should be a particular 
        color (property -gt 5, property -like "*Apple*", etc).
        
        You can add the function to your scripts, dot source it to load into your current
        PowerShell session or add it to your $Profile so it is always available.
        
        To dot source:
            .".\Set-CellColor.ps1"
            
    .PARAMETER Property
        Property, or column that you will be keying on.  
    .PARAMETER Color
        Name or 6-digit hex value of the color you want the cell to be
    .PARAMETER InputObject
        HTML you want the script to process.  This can be entered directly into the
        parameter or piped to the function.
    .PARAMETER Filter
        Specifies a query to determine if a cell should have its color changed.  $true
        results will make the color change while $false result will return nothing.
        
        Syntax
        <Property Name> <Operator> <Value>
        
        <Property Name>::= the same as $Property.  This must match exactly
        <Operator>::= "-eq" | "-le" | "-ge" | "-ne" | "-lt" | "-gt"| "-approx" | "-like" | "-notlike" 
            <JoinOperator> ::= "-and" | "-or"
            <NotOperator> ::= "-not"
        
        The script first attempts to convert the cell to a number, and if it fails it will
        cast it as a string.  So 40 will be a number and you can use -lt, -gt, etc.  But 40%
        would be cast as a string so you could only use -eq, -ne, -like, etc.  
    .PARAMETER Row
        Instructs the script to change the entire row to the specified color instead of the individual cell.
    .INPUTS
        HTML with table
    .OUTPUTS
        HTML
    .EXAMPLE
        get-process | convertto-html | set-cellcolor -Propety cpu -Color red -Filter "cpu -gt 1000" | out-file c:\test\get-process.html

        Assuming Set-CellColor has been dot sourced, run Get-Process and convert to HTML.  
        Then change the CPU cell to red only if the CPU field is greater than 1000.
        
    .EXAMPLE
        get-process | convertto-html | set-cellcolor cpu red -filter "cpu -gt 1000 -and cpu -lt 2000" | out-file c:\test\get-process.html
        
        Same as Example 1, but now we will only turn a cell red if CPU is greater than 100 
        but less than 2000.
        
    .EXAMPLE
        $HTML = $Data | sort server | ConvertTo-html -head $header | Set-CellColor cookedvalue red -Filter "cookedvalue -gt 1"
        PS C:\> $HTML = $HTML | Set-CellColor Server green -Filter "server -eq 'dc2'"
        PS C:\> $HTML | Set-CellColor Path Yellow -Filter "Path -like ""*memory*""" | Out-File c:\Test\colortest.html
        
        Takes a collection of objects in $Data, sorts on the property Server and converts to HTML.  From there 
        we set the "CookedValue" property to red if it's greater then 1.  We then send the HTML through Set-CellColor
        again, this time setting the Server cell to green if it's "dc2".  One more time through Set-CellColor
        turns the Path cell to Yellow if it contains the word "memory" in it.
        
    .EXAMPLE
        $HTML = $Data | sort server | ConvertTo-html -head $header | Set-CellColor cookedvalue red -Filter "cookedvalue -gt 1" -Row
        
        Now, if the cookedvalue property is greater than 1 the function will highlight the entire row red.
        
    .NOTES
        Author:             Martin Pugh
        Twitter:            @thesurlyadm1n
        Spiceworks:         Martin9700
        Blog:               www.thesurlyadmin.com
          
        Changelog:
            1.5             Added ability to set row color with -Row switch instead of the individual cell
            1.03            Added error message in case the $Property field cannot be found in the table header
            1.02            Added some additional text to help.  Added some error trapping around $Filter
                            creation.
            1.01            Added verbose output
            1.0             Initial Release
    .LINK
        http://community.spiceworks.com/scripts/show/2450-change-cell-color-in-html-table-with-powershell-set-cellcolor
    #>

    [CmdletBinding()]
    Param (
        [Parameter(Mandatory,Position=0)]
        [string]$Property,
        [Parameter(Mandatory,Position=1)]
        [string]$Color,
        [Parameter(Mandatory,ValueFromPipeline)]
        [Object[]]$InputObject,
        [Parameter(Mandatory)]
        [string]$Filter,
        [switch]$Row
    )
    
    Begin {
        Write-Verbose "$(Get-Date): Function Set-CellColor begins"
        If ($Filter)
        {   If ($Filter.ToUpper().IndexOf($Property.ToUpper()) -ge 0)
            {   $Filter = $Filter.ToUpper().Replace($Property.ToUpper(),"`$Value")
                Try {
                    [scriptblock]$Filter = [scriptblock]::Create($Filter)
                }
                Catch {
                    Write-Warning "$(Get-Date): ""$Filter"" caused an error, stopping script!"
                    Write-Warning $Error[0]
                    Exit
                }
            }
            Else
            {   Write-Warning "Could not locate $Property in the Filter, which is required.  Filter: $Filter"
                Exit
            }
        }
    }
    
    Process {
        ForEach ($Line in $InputObject)
        {   If ($Line.IndexOf("<tr><th") -ge 0)
            {   Write-Verbose "$(Get-Date): Processing headers..."
                $Search = $Line | Select-String -Pattern '<th ?[a-z\-:;"=]*>(.*?)<\/th>' -AllMatches
                $Index = 0
                ForEach ($Match in $Search.Matches)
                {   If ($Match.Groups[1].Value -eq $Property)
                    {   Break
                    }
                    $Index ++
                }
                If ($Index -eq $Search.Matches.Count)
                {   Write-Warning "$(Get-Date): Unable to locate property: $Property in table header"
                    Exit
                }
                Write-Verbose "$(Get-Date): $Property column found at index: $Index"
            }
            If ($Line -match "<tr( style=""background-color:.+?"")?><td")
            {   $Search = $Line | Select-String -Pattern '<td ?[a-z\-:;"=]*>(.*?)<\/td>' -AllMatches
                $Value = $Search.Matches[$Index].Groups[1].Value -as [double]
                If (-not $Value)
                {   $Value = $Search.Matches[$Index].Groups[1].Value
                }
                If (Invoke-Command $Filter)
                {   If ($Row)
                    {   Write-Verbose "$(Get-Date): Criteria met!  Changing row to $Color..."
                        If ($Line -match "<tr style=""background-color:(.+?)"">")
                        {   $Line = $Line -replace "<tr style=""background-color:$($Matches[1])","<tr style=""background-color:$Color"
                        }
                        Else
                        {   $Line = $Line.Replace("<tr>","<tr style=""background-color:$Color"">")
                        }
                    }
                    Else
                    {   Write-Verbose "$(Get-Date): Criteria met!  Changing cell to $Color..."
                        $Line = $Line.Replace($Search.Matches[$Index].Value,"<td style=""background-color:$Color"">$Value</td>")
                    }
                }
            }
            Write-Output $Line
        }
    }
    
    End {
        Write-Verbose "$(Get-Date): Function Set-CellColor completed"
    }
}

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
        [String]$Attachment,
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

        if ($Attachment -ne "") {
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

function getDeviceGroups($baseuri) {
    # Define the lookup hash to enumerate the types of alerts
    $healthHash = @{"0" = "None"; "2" = "Unknown"; "4" = "Normal"; "8" = "Warning"; "16" = 'Critical'}

    # Get Device Groups
    $filterXml = Invoke-RestMethod -Method Get -UseDefaultCredentials -Uri "$baseuri/DeviceGroups"
    # DEBUG COMMENT: print the Device Groups
    # $filterXml.GetDeviceGroupsResponse.GetDeviceGroupsResult.DeviceGroup | Format-Table -Property Description, DeviceCount, Id, Name, RollupHealth, Type
    
    # Store the Id for the DeviceGroups which have device allocated to them
    $deviceGroupArray = $filterXml.GetDeviceGroupsResponse.GetDeviceGroupsResult.DeviceGroup|where {($_.DeviceCount -gt 0) -and ($_.Name -ne "Unknown") -and ($_.Name -ne "All Devices")}

    #$deviceGroupArray | Format-Table -Property Description, DeviceCount, Id, Name, RollupHealth, Type
    $deviceGroupArray = $deviceGroupArray | ForEach-Object {
        $testVar = $_.RollupHealth
        $_.RollupHealth = $healthHash.$testvar
        $_
    }
    $deviceGroupArray
    #$script:criticalFilterId = $criticalFilter.Id
    #$warningFilter = $filterXml.GetAllAlertFiltersResponse.GetAllAlertFiltersResult.AlertFilter|where {$_.name -contains "Warning Alerts"}
    #$script:warningFilterId = $warningFilter.Id
}

function getDeviceGroupSummary($groupId) {
    # Get Device Groups
    $filterXml = Invoke-RestMethod -Method Get -UseDefaultCredentials -Uri "$baseuri/DeviceGroups/$groupId/Summary"
    # DEBUG COMMENT: print the Device Groups
    # $filterXml.GetDeviceGroupsResponse.GetDeviceGroupsResult.DeviceGroup | Format-Table -Property Description, DeviceCount, Id, Name, RollupHealth, Type
    
    # Store the Id for the DeviceGroups which have device allocated to them
    $deviceGroupSummaryArray = $filterXml.GetDeviceGroupSummaryResponse.GetDeviceGroupSummaryResult
    $deviceGroupSummaryArray
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
$groupObjects = getDeviceGroups($baseuri)

$Header = @"
<style>
TABLE {border-width: 1px;border-style: solid;border-color: black;border-collapse: collapse;}
TH {border-width: 1px;padding: 3px;border-style: solid;border-color: black;background-color: #6495ED; font-family:courier; font-size:12pt;}
TD {border-width: 1px;padding: 3px;border-style: solid;border-color: black;}
BODY {font-family:courier; font-size:10pt;}
</style>
<title>
Device Groups - Health Summary
</title>
"@

$Pre = @"
<div style='margin:  0px auto; BACKGROUND-COLOR:Black;Color:White;font-weight:bold;FONT-SIZE:  16pt;TEXT-ALIGN: center;'>
Dell OpenManager Essentials - Health Rollup Information
</div>
</p>   
"@

$Post = @"
</p>
<div style='margin:  0px auto; BACKGROUND-COLOR:Black;Color:White;font-weight:bold;FONT-SIZE:  16pt;TEXT-ALIGN: center;'>
Dell OpenManager Essentials - Physical Infrastructure Health
</div>
</p>   
"@


$HTML = $groupObjects | Select-Object Name, DeviceCount, RollupHealth | ConvertTo-Html -Head $Header -PreContent $Pre -PostContent $Post | Set-CellColor -Property RollupHealth -Color red -Filter "RollupHealth -like '*Critical*'" 
$HTML = $HTML | Set-CellColor -Property RollupHealth -Color green -Filter "RollupHealth -like '*Normal*'"
$HTML = $HTML | Set-CellColor -Property RollupHealth -Color grey -Filter "RollupHealth -like '*Unknown*'"
$HTML = $HTML | Set-CellColor -Property RollupHealth -Color orange -Filter "RollupHealth -like '*Warning*'" 
# Invoke-Item test.html

foreach ($deviceGroupId in $groupObjects) {
    $deviceName = $deviceGroupId.Name
    $Pre = @"
<p><strong>Device Group Name: $deviceName</p>
"@
    $HTML2 = getDeviceGroupSummary($deviceGroupId.Id) | Select-Object DeviceCount, CriticalCount, WarningCount, NormalCount, UnknownCount | ConvertTo-Html -PreContent $Pre | Set-CellColor CriticalCount -Color red -Filter "CriticalCount -gt 0" -Row
    $HTML = $HTML + $HTML2
}

# For debugging purposes (and testing), uncomment the below and it will show the output in a HTML IE frame
# $HTML | Out-File test.html
# Invoke-Item test.html

Send-EMail -EmailTo "storms.cba.privatecloud@emc.com" -Subject "Infrastructure Health Rollup" -Body $HTML
Send-EMail -EmailTo "wwonigkeit@iqconsult.com.au" -Subject "Infrastructure Health Rollup" -Body $HTML

#
# ------------------ END MAIN SECTION --------------------------
