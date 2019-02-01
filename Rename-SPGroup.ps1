<#
.SYNOPSIS
Changes the display name of AD groups based on a CSV file.

.NOTES
The user executing this script must be an SCA on all Site Collections.
Import-Csv assumes the 1st row is the header and therefore the attribute name

#>

# csv file to import
$inputCSV = "f:\DET-AD-Changes.csv"
# folder to write files to
$outputDirectory = "F:\"
$debug = $false

# load PowerShell add-in for SharePoint
$ver = $host | select version
if($ver.Version.Major -gt 1){ $host.Runspace.ThreadOptions = "ReuseThread" }
if((Get-PSSnapin "Microsoft.SharePoint.PowerShell" -ErrorAction SilentlyContinue) -eq $null){ Add-PSSnapin "Microsoft.SharePoint.PowerShell" }

function Rename-SPGroup
{
    [cmdletbinding(SupportsShouldProcess=$True)]

    param(
        [Parameter(Mandatory=$true)]
        [ValidateNotNullOrEmpty()]
        [string]$ExistingName,

        [Parameter(Mandatory=$true)]
        [ValidateNotNullOrEmpty()]
        [string]$NewName
    )

    Begin
    {
        if((Get-PSSnapin "Microsoft.SharePoint.PowerShell" -ErrorAction SilentlyContinue) -eq $null){ Add-PSSnapin "Microsoft.SharePoint.PowerShell" }
        $spVersion = (Get-SPFarm).BuildVersion
    }

    Process
    {
        ForEach($site in Get-SPSite -Limit ALL)
        {
            Write-Verbose "Processing Site Collection $($site.Url)"
            Write-Verbose "--Processing $($ExistingName)"
            if($spVersion.Major -eq 14)
            {
                $group = $site.OpenWeb().SiteUsers | ?{ $_.DisplayName -like "*$($ExistingName)" }
            }
            elseif($spVersion.Major -ge 15)
            {
                $web = $site.RootWeb
                $group = $web.SiteUsers | ?{ $_.DisplayName -match $ExistingName }
            }
            else
            {
                Write-Host "Just which version of SharePoint are you on? I can't figure it out. Aborting..."
            }

            if($group)
            {
                if($PSCmdlet.ShouldProcess($ExistingName,"Change Display Name to $($NewName)"))
                {
                    Write-Verbose "----Changing display name of group $($ExistingName) to $($NewName)"
                    if($spVersion.Major -ge 15)
                    {
                        Set-SPUser -Identity $group -Web $web -DisplayName $newName -Confirm:$false
                    }
                    elseif($spVersion.Major -eq 14)
                    {
                        $group.DisplayName = $newName
                        $group.Update()
                    }
                }
            }
            else
            {
                Write-Verbose "----Group $($ExistingName) not found in this Site Collection...skipping"
            }
        }
    }

    End
    {
        # do nothing
    }
}

$start = Get-Date
# start PowerShell transcript as record of actions
$timestamp = Get-Date -UFormat "%Y%m%d%H%M%S"
$transcript = $outputDirectory.TrimEnd("\") + "\grouprename_" + $timestamp + ".txt"
Start-Transcript -Path $transcript -Force

$rows = Import-Csv $inputCSV

# process all rows in file

ForEach($row in $rows)
{
        $i++
        $group= $null
        $oldName = $row.'Old Name'.TrimStart("*")
        $newName = $row.'New Name'.TrimStart("*")
        if($debug){ Write-Host "Row $($i) - $($oldName)" -ForegroundColor Yellow }
        
        # skip this row if group names are missing
        if(!$oldName -or !$newName)
        {
            if($debug){ Write-Host "Skipping row because missing a name" -ForegroundColor Yellow }
            Continue
        }
        
        If($newName -notmatch "remove")
        {
            $splat = @{
                ExistingName = $oldName
                NewName = $newName
            }

            if($debug)
            {
                $splat.Add("WhatIf", $true)
                $splat.Add("Verbose", $true)
            }
            Rename-SPGroup @splat
        }
        elseif($debug)
        {
            Write-Host "Group $($oldName) was removed...skipping" -ForegroundColor Yellow
        }
    }


Write-Output ""
$timetaken = (Get-Date) - $start
Write-Output "Action completed in $($timetaken.Hours) hours and $($timetaken.Minutes) minutes."
Stop-Transcript
