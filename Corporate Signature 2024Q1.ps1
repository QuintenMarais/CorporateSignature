# Connecting to Azure Parameters
$tenantID = ""
$applicationID = ""
$clientKey = ""
# URL of the website where the config file is available
$configUrl  = "https://config.txt"

$username = ""
$password = ""
$credential = New-Object System.Management.Automation.PSCredential ($username, (ConvertTo-SecureString $password -AsPlainText -Force))


#############################################
# Get Local Stuff Ready
#############################################
clear

Write-Host
Write-Host "Get basic Outlook Information $(Get-Date -Format 'yyyy-MM-ddTHH:mm:ssK')"
$OutlookProfiles = @()
$OutlookUseNewOutlook = $null

    if ($(Get-Command -Name 'Get-AppPackage' -ErrorAction SilentlyContinue -WarningAction SilentlyContinue)) 
            {
                $NewOutlook = Get-AppPackage -Name 'Microsoft.OutlookForWindows' -ErrorAction SilentlyContinue -WarningAction SilentlyContinue
            } 
    else 
            {
                $NewOutlook = $null
            }

    Write-Host 'Outlook'
    $OutlookRegistryVersion = [System.Version]::Parse(((((((Get-ItemProperty 'Registry::HKEY_CLASSES_ROOT\Outlook.Application\CurVer' -ErrorAction SilentlyContinue).'(default)' -ireplace [Regex]::Escape('Outlook.Application.'), '') + '.0.0.0.0')) -ireplace '^\.', '' -split '\.')[0..3] -join '.'))

    try {
        # [Microsoft.Win32.RegistryView]::Registry32 makes sure view the registry as a 32 bit application would
        # This is independent from the bitness of the PowerShell process, while Get-ItemProperty always uses the bitness of the PowerShell process
        # Covers:
        #   Office x86 on Windows x86
        #   Office x86 on Windows x64
        #   Any PowerShell process bitness
        $OutlookFilePath = Get-ChildItem (((([Microsoft.Win32.RegistryKey]::OpenBaseKey([Microsoft.Win32.RegistryHive]::ClassesRoot, [Microsoft.Win32.RegistryView]::Registry32)).OpenSubKey("CLSID\$((Get-ItemProperty 'Registry::HKEY_CLASSES_ROOT\Outlook.Application\CLSID' -ErrorAction Stop).'(default)')\LocalServer32")).GetValue('') -split ' \/')[0].Split([IO.Path]::GetInvalidPathChars()) -join '') -ErrorAction Stop
    } catch {
        try {
            # [Microsoft.Win32.RegistryView]::Registry64 makes sure we view the registry as a 64 bit application would
            # This is independent from the bitness of the PowerShell process, while Get-ItemProperty always uses the bitness of the PowerShell process
            # Covers:
            #   Office x64 on Windows x64
            #   Any PowerShell process bitness
            $OutlookFilePath = Get-ChildItem (((([Microsoft.Win32.RegistryKey]::OpenBaseKey([Microsoft.Win32.RegistryHive]::ClassesRoot, [Microsoft.Win32.RegistryView]::Registry64)).OpenSubKey("CLSID\$((Get-ItemProperty 'Registry::HKEY_CLASSES_ROOT\Outlook.Application\CLSID' -ErrorAction Stop).'(default)')\LocalServer32")).GetValue('') -split ' \/')[0].Split([IO.Path]::GetInvalidPathChars()) -join '') -ErrorAction Stop
        } catch {
            $OutlookFilePath = $null
        }
    }

    if ($OutlookFilePath) {
        try {
            $OutlookBitnessInfo = GetBitness -fullname $OutlookFilePath
            $OutlookFileVersion = [System.Version]::Parse((((($OutlookBitnessInfo.'File Version'.ToString() + '.0.0.0.0')) -ireplace '^\.', '' -split '\.')[0..3] -join '.'))
            $OutlookBitness = $OutlookBitnessInfo.Architecture
            Remove-Variable -Name 'OutlookBitnessInfo'
        } catch {
            $OutlookBitness = 'Error'
            $OutlookFileVersion = $null
        }
    } else {
        $OutlookBitness = $null
        $OutlookFileVersion = $null
    }

    if ($OutlookRegistryVersion.major -eq 0) {
        $OutlookRegistryVersion = $null
    } elseif ($OutlookRegistryVersion.major -gt 16) {
        Write-Host "    Outlook version $OutlookRegistryVersion is newer than 16 and not yet known. Please inform your administrator. Exit." -ForegroundColor Red
        exit 1
    } elseif ($OutlookRegistryVersion.major -eq 16) {
        $OutlookRegistryVersion = '16.0'
    } elseif ($OutlookRegistryVersion.major -eq 15) {
        $OutlookRegistryVersion = '15.0'
    } elseif ($OutlookRegistryVersion.major -eq 14) {
        $OutlookRegistryVersion = '14.0'
    } elseif ($OutlookRegistryVersion.major -lt 14) {
        Write-Host "    Outlook version $OutlookRegistryVersion is older than Outlook 2010 and not supported. Please inform your administrator. Exit." -ForegroundColor Red
        exit 1
    }


    if ($null -ne $OutlookRegistryVersion) {
        try {
            $OutlookDefaultProfile = (Get-ItemProperty "hkcu:\software\microsoft\office\$($OutlookRegistryVersion)\Outlook" -ErrorAction Stop -WarningAction SilentlyContinue).DefaultProfile

            $OutlookProfiles = @(@((Get-ChildItem "hkcu:\SOFTWARE\Microsoft\Office\$($OutlookRegistryVersion)\Outlook\Profiles" -ErrorAction Stop -WarningAction SilentlyContinue).PSChildName) | Where-Object { $_ })

            if ($OutlookDefaultProfile -and ($OutlookDefaultProfile -iin $OutlookProfiles)) {
                $OutlookProfiles = @(@($OutlookDefaultProfile) + @($OutlookProfiles | Where-Object { $_ -ine $OutlookDefaultProfile }))
            }
        } catch {
            $OutlookDefaultProfile = $null
            $OutlookProfiles = @()
        }

        $OutlookIsBetaversion = $false

        if (
            ((Get-Item 'registry::HKEY_LOCAL_MACHINE\Software\Microsoft\Office\ClickToRun\Configuration' -ErrorAction SilentlyContinue -WarningAction SilentlyContinue).Property -contains 'UpdateChannel') -and
            ($OutlookFileVersion -ge '16.0.0.0')
        ) {
            $x = (Get-ItemProperty 'registry::HKEY_LOCAL_MACHINE\Software\Microsoft\Office\ClickToRun\Configuration' -ErrorAction Stop -WarningAction SilentlyContinue).'UpdateChannel'

            if ($x -ieq 'http://officecdn.microsoft.com/pr/5440FD1F-7ECB-4221-8110-145EFAA6372F') {
                $OutlookIsBetaversion = $true
            }

            if ((Get-Item "registry::HKEY_LOCAL_MACHINE\Software\Policies\Microsoft\Office\$($OutlookRegistryVersion)\Common\OfficeUpdate" -ErrorAction SilentlyContinue -WarningAction SilentlyContinue).Property -contains 'UpdateBranch') {
                $x = (Get-ItemProperty "registry::HKEY_LOCAL_MACHINE\Software\Policies\Microsoft\Office\$($OutlookRegistryVersion)\Common\OfficeUpdate" -ErrorAction SilentlyContinue -WarningAction SilentlyContinue).'UpdateBranch'

                if ($x -ieq 'InsiderFast') {
                    $OutlookIsBetaversion = $true
                }
            }
        }

        $OutlookDisableRoamingSignatures = 0

        foreach ($RegistryFolder in (
                "registry::HKEY_CURRENT_USER\Software\Microsoft\Office\$($OutlookRegistryVersion)\Outlook\Setup",
                "registry::HKEY_LOCAL_MACHINE\Software\Microsoft\Office\$($OutlookRegistryVersion)\Outlook\Setup",
                "registry::HKEY_CURRENT_USER\Software\Policies\Microsoft\Office\$($OutlookRegistryVersion)\Outlook\Setup",
                "registry::HKEY_LOCAL_MACHINE\Software\Policies\Microsoft\Office\$($OutlookRegistryVersion)\Outlook\Setup"
            )) {

            $x = (Get-ItemProperty $RegistryFolder -ErrorAction SilentlyContinue).'DisableRoamingSignaturesTemporaryToggle'

            if (($x -in (0, 1)) -and ($OutlookFileVersion -ge '16.0.0.0')) {
                $OutlookDisableRoamingSignatures = $x
            }

            $x = (Get-ItemProperty $RegistryFolder -ErrorAction SilentlyContinue).'DisableRoamingSignatures'

            if (($x -in (0, 1)) -and ($OutlookFileVersion -ge '16.0.0.0')) {
                $OutlookDisableRoamingSignatures = $x
            }
        }

        if ($NewOutlook -and ($((Get-ItemProperty "registry::HKEY_CURRENT_USER\Software\Microsoft\Office\$($OutlookRegistryVersion)\Outlook\Preferences" -ErrorAction SilentlyContinue).'UseNewOutlook') -eq 1)) {
            $OutlookUseNewOutlook = $true
        } else {
            $OutlookUseNewOutlook = $false
        }
    } else {
        $OutlookDefaultProfile = $null
        $OutlookDisableRoamingSignatures = $null
        $OutlookIsBetaVersion = $null

        if ($NewOutlook) {
            $OutlookUseNewOutlook = $true
        } else {
            $OutlookUseNewOutlook = $false
        }
    }

    Write-Host "    Registry version: $OutlookRegistryVersion"
    Write-Host "    File version: $OutlookFileVersion"
    if (($OutlookFileVersion -lt '16.0.0.0') -and ($EmbedImagesInHtml -eq $true)) {
        Write-Host '      Outlook 2013 or earlier detected.' -ForegroundColor Yellow
        Write-Host '      Consider parameter ''EmbedImagesInHtml false'' to avoid problems with images in templates.' -ForegroundColor Yellow
        Write-Host '      Microsoft supports Outlook 2013 until April 2023, older versions are already out of support.' -ForegroundColor Yellow
    }
    Write-Host "    Bitness: $OutlookBitness"
    Write-Host "    Default profile: $OutlookDefaultProfile"
    Write-Host "    DisableRoamingSignatures: $OutlookDisableRoamingSignatures"
    Write-Host "    UseNewOutlook: $OutlookUseNewOutlook"
    Write-Host '  New Outlook'
    Write-Host "    Version: $($NewOutlook.Version)"
    Write-Host "    Status: $($NewOutlook.Status)"
    Write-Host "    UseNewOutlook: $OutlookUseNewOutlook"



# Get the email address associated with the current user's Outlook account

$outlookEmail = Get-ItemProperty -Path "Registry::HKEY_CURRENT_USER\Software\Microsoft\Office\$($OutlookRegistryVersion)\Outlook\Profiles\Outlook\9375CFF0413111d3B88A00104B2A6676\00000002" -Name "Account Name"

# Output the email address
if ($outlookEmail -ne $null) {

    $emailaddress = $($outlookEmail.'Account Name')
} else {
    Write-Output "Outlook email address not found."
}

Write-Output "Outlook Email Address: $($outlookEmail.'Account Name')"
##########################################################
# Get the Signiture folder details
##########################################################

Write-Host
    Write-Host "Get Outlook signature file path(s) $(Get-Date -Format 'yyyy-MM-dd')"
    $SignaturePath = 
    $x = (Get-ItemProperty "hkcu:\software\microsoft\office\$($OutlookRegistryVersion)\common\general" -ErrorAction SilentlyContinue).'Signatures'
    Push-Location ((Join-Path -Path ($env:AppData) -ChildPath 'Microsoft'))
    $x = ($ExecutionContext.SessionState.Path.GetUnresolvedProviderPathFromPSPath($x))
    $SignaturePath = $x
    Write-Host "Signatures are located here : $SignaturePath"
    Write-Host

##############################################
# Graph Part
##############################################

# Authenticate to Microsoft Grpah
Write-Host
Write-Host "Authenticating to Microsoft Graph via REST method"
 
$url = "https://login.microsoftonline.com/$tenantId/oauth2/token"
$resource = "https://graph.microsoft.com/"
$restbody = @{
         grant_type    = 'client_credentials'
         client_id     = $applicationID
         client_secret = $clientKey
         resource      = $resource
}
     
 # Get the return Auth Token
$token = Invoke-RestMethod -Method POST -Uri $url -Body $restbody
     

# Pack the token into a header for future API calls
$header = @{
          'Authorization' = "$($Token.token_type) $($Token.access_token)"
         'Content-type'  = "application/json"
}
 

# Build the Base URL for the API call

$url = "https://graph.microsoft.com/v1.0/users/$emailaddress"
# Call the REST-API
$userPurpose = Invoke-RestMethod -Method GET -headers $header -Uri $url


##############################################
# Get the latest confg file
##############################################
Write-Host
Write-Host "Getting the latest Config File"

try {
    # Download the configuration file
    $configScript = Invoke-WebRequest -Uri $configUrl -Credential $credential | Select-Object -ExpandProperty Content 

    # Execute the configuration script to set variables
    Invoke-Expression $configScript

    # Now you can use the variables declared in the configuration script
#    Write-Host "Source: $Source"
#    Write-Host "ExcludeList: $($ExcludeList -join ', ')"
#    Write-Host "Version: $Version"

    # Add your script logic here using the variables from the configuration file
} catch {
    Write-Host "Failed to download or execute the configuration script: $_"
}
##############################################
# Download the zip File
##############################################
Write-Host
Write-Host "Downloading the latest Signature"
try {
    $tempFolder = [System.IO.Path]::GetTempPath()
    $downloadPath = Join-Path -Path $tempFolder -ChildPath $Source
    Invoke-WebRequest -Uri "https://automation.za.logicalis.com/CorporateSignature/$Source" -Credential $credential -OutFile $downloadPath  



  # Add your script logic here using the variables from the configuration file and the downloaded files
} catch {
  Write-Host "Failed to download, unzip, or execute the configuration script: $_"
}
##############################################
# Unzip the downloaded file to the temporary folder
##############################################
Write-Host
Write-Host "Expanding the Signature"

$unzipFolder = $SignaturePath
Expand-Archive -Path $downloadPath -DestinationPath $unzipFolder -Force

Write-Host
Write-Host "Creating your Custom Signature"



# Remove leading zeros
$MobileNumber = $userPurpose.mobilePhone.TrimStart('0')

# Add country code and format
$mobilePhone = "+27 " + $MobileNumber.Substring(0, 2) + " " + $MobileNumber.Substring(2, 3) + " " + $MobileNumber.Substring(5, 4)



$displayName = $userPurpose.displayName
$jobTitle = $userPurpose.jobTitle
$OfficeNumber = "+27 11 111 1111"
$SignatureName = "Corporate_Signature_2024"


(Get-Content "$SignaturePath\Corporate_Signature_2024.htm") -Replace '__YOURNAME__', $displayName | Set-Content "$SignaturePath\Corporate_Signature_2024.htm"
(Get-Content "$SignaturePath\Corporate_Signature_2024.htm") -Replace '__YOURTITLE__', $jobTitle | Set-Content "$SignaturePath\Corporate_Signature_2024.htm"
(Get-Content "$SignaturePath\Corporate_Signature_2024.htm") -Replace '__YOURCELLNUMBER__', $mobilePhone | Set-Content "$SignaturePath\Corporate_Signature_2024.htm"
(Get-Content "$SignaturePath\Corporate_Signature_2024.htm") -Replace '__YOURLOGICALIS__', "Logicalis South Africa" | Set-Content "$SignaturePath\Corporate_Signature_2024.htm"
(Get-Content "$SignaturePath\Corporate_Signature_2024.htm") -Replace '__YOURMOBILENUMBER__', $OfficeNumber | Set-Content "$SignaturePath\Corporate_Signature_2024.htm"

(Get-Content "$SignaturePath\Corporate_Signature_2024.txt") -Replace '__YOURNAME__', $displayName | Set-Content "$SignaturePath\Corporate_Signature_2024.txt"
(Get-Content "$SignaturePath\Corporate_Signature_2024.txt") -Replace '__YOURTITLE__', $jobTitle | Set-Content "$SignaturePath\Corporate_Signature_2024.txt"
(Get-Content "$SignaturePath\Corporate_Signature_2024.txt") -Replace '__YOURCELLNUMBER__', $mobilePhone | Set-Content "$SignaturePath\Corporate_Signature_2024.txt"
(Get-Content "$SignaturePath\Corporate_Signature_2024.txt") -Replace '__YOURLOGICALIS__', "Logicalis South Africa" | Set-Content "$SignaturePath\Corporate_Signature_2024.txt"
(Get-Content "$SignaturePath\Corporate_Signature_2024.txt") -Replace '__YOURMOBILENUMBER__', $OfficeNumber | Set-Content "$SignaturePath\Corporate_Signature_2024.txt"

##########################################################
# Setting as Default Signature
##########################################################

if ($DisableRoamingSignatures -eq $True) {
    Write-Host
    Write-Host "Setting Disable Roaming Signatures"
    $null = "HKCU:\Software\Microsoft\Office\$($OutlookRegistryVersion)\Outlook\Setup" | ForEach-Object { if (Test-Path $_) { Get-Item $_ } else { New-Item $_ -Force } } | New-ItemProperty -Name 'DisableRoamingSignatures' -Type DWORD -Value 1 -Force
}

if ($DisableRoamingSignaturesTemporaryToggle -eq $True) {
    Write-Host
    Write-Host "Setting Disable Roaming Signatures Temporary Toggle"
    $null = "HKCU:\Software\Microsoft\Office\$($OutlookRegistryVersion)\Outlook\Setup" | ForEach-Object { if (Test-Path $_) { Get-Item $_ } else { New-Item $_ -Force } } | New-ItemProperty -Name 'DisableRoamingSignaturesTemporaryToggle' -Type DWORD -Value 1 -Force
}

if ($SetDefaultNewSignature -eq  $True) {
    Write-Host "Setting Default New signature..."
    # Creating new registry keys for signature
    $null = get-item -path HKCU:\\Software\\Microsoft\\Office\\$OutlookRegistryVersion\\Common\\MailSettings | new-Itemproperty -name NewSignature -value $SignatureName -propertytype string -force
}
if ($SetDefaultReplySignature -eq  $True) {
    Write-Host "Setting Default Reply signature..."
    # Creating new registry keys for signature
    
    $null = get-item -path HKCU:\\Software\\Microsoft\\Office\\$OutlookRegistryVersion\\Common\\MailSettings | new-Itemproperty -name ReplySignature -value $SignatureName -propertytype string -force
}

Write-Host
Write-Host "We're all done"