[CmdletBinding()]
param (
    [string]$Path,
    [string]$Version = 'master'
)

$localpath = $(Join-Path -Path (Split-Path -Path $profile) -ChildPath '\Modules\HttpEnabledSsrsPsTools')

try
{
    if ($Path.length -eq 0)
    {
        if ($PSCommandPath.Length -gt 0)
        {
            $path = Split-Path $PSCommandPath
            if ($path -match "github")
            {
                $path = $localpath
            }
        }
        else
        {
            $path = $localpath
        }
    }
}
catch
{
    $path = $localpath
}

if ($path.length -eq 0)
{
    $path = $localpath
}

if ((Get-Command -Module HttpEnabledSsrsPsTools).count -ne 0)
{
    Write-Output "Removing existing HttpEnabledSsrsPsTools Module..."
    Remove-Module HttpEnabledSsrsPsTools -ErrorAction Stop
}

$url = "https://github.com/schwarrior/HttpEnabledSsrsPsTools/archive/$Version.zip"

$temp = ([System.IO.Path]::GetTempPath()).TrimEnd("\")
$zipfile = "$temp\HttpEnabledSsrsPsTools.zip"

if (!(Test-Path -Path $path))
{
    try
    {
        Write-Output "Creating directory: $path..."
        New-Item -Path $path -ItemType Directory | Out-Null
    }
    catch
    {
        throw "Can't create $Path. You may need to Run as Administrator!"
    }
}
else
{
    try
    {
        Write-Output "Deleting previously installed module..."
        Remove-Item -Path "$path\*" -Force -Recurse
    }
    catch
    {
        throw "Can't delete $Path. You may need to Run as Administrator!"
    }
}

Write-Output "Downloading archive from ReportingServiceTools GitHub..."
[Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12
try
{
    Invoke-WebRequest $url -OutFile $zipfile
}
catch
{
    #try with default proxy and usersettings
    Write-Output "...Probably using a proxy for internet access. Trying default proxy settings..."
    (New-Object System.Net.WebClient).Proxy.Credentials = [System.Net.CredentialCache]::DefaultNetworkCredentials
    Invoke-WebRequest $url -OutFile $zipfile -ErrorAction Stop
}

# Unblock if there's a block
Unblock-File $zipfile -ErrorAction SilentlyContinue

# Keep it backwards compatible
Write-Output "Unzipping archive..."
$shell = New-Object -COM Shell.Application
$zipPackage = $shell.NameSpace($zipfile)
$destinationFolder = $shell.NameSpace($temp)
$destinationFolder.CopyHere($zipPackage.Items())
Move-Item -Path "$temp\HttpEnabledSsrsPsTools-$Version\*" $path
Write-Output "HttpEnabledSsrsPsTools has been successfully downloaded to $path!"

Write-Output "Cleaning up..."
Remove-Item -Path "$temp\HttpEnabledSsrsPsTools-$Version"
Remove-Item -Path $zipfile

Write-Output "Importing HttpEnabledSsrsPsTools Module..."
Import-Module "$path\HttpEnabledSsrsPsTools\HttpEnabledSsrsPsTools.psd1" -Force
Write-Output "HttpEnabledSsrsPsTools Module was successfully imported!"

Get-Command -Module HttpEnabledSsrsPsTools
Write-Output "`n`nIf you experience any function missing errors after update, please restart PowerShell or reload your profile."
