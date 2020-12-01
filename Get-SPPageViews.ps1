function Get-SPPageViews
{
    <#
    Usage: Get-SPPageViews -RootSiteUrl "http://hej.lab.roblab.com" -OutputFilepath "F:\pspageviews_$(Get-date -Format "yyyyMMddHHmm").csv" -IncludeSites -IncludeWebs -DeleteExistingFile
    #>
	[CmdletBinding(
		RemotingCapability = "PowerShell",
		HelpUri = "",
		ConfirmImpact = "None", #None, Loew, Medium, or High
		DefaultParameterSetName = "",
		SupportsShouldProcess = $false,
		SupportsTransactions = $false,
		SupportsPaging = $false
	)]

	[OutputType([string])]

	Param(
		[string]$RootSiteUrl,
		[string]$Scope,
		[switch]$IncludeSites,
		[switch]$IncludeWebs,
		[switch]$DeleteExistingFile,
		[switch]$SuppressHeader,
		[string]$OutputFilepath
	)

	Begin
	{
		$ver = $host | select version
		if($ver.Version.Major -gt 1){ $host.Runspace.ThreadOptions = "ReuseThread" }
		if((Get-PSSnapin "Microsoft.SharePoint.PowerShell" -ErrorAction SilentlyContinue) -eq $null){ Add-PSSnapin "Microsoft.SharePoint.PowerShell" }

		$tDate = Get-Date
		# Rename csv file if it exists
		if((Test-Path $OutputFilepath) -and ($DeleteExistingFile.IsPresent))
		{
			$stamp = Get-Date -UFormat "%Y%m%d%H%M%S"
			$fileObj = Get-ChildItem $OutputFilepath
			$fileExt = $fileObj.Extension
			$fileB4Ext = $fileObj.Name.Replace($fileObj.Extension,'')
			$oldFile = $fileObj.Name

			Write-Verbose "Renaming file $($fileObj.Name) to $($fileB4Ext)-$($stamp)$($fileExt)"
			Rename-Item "$($fileObj.Name)" "$($fileB4Ext)-$($stamp)$($fileExt)"
		}

		# Write header row to file
		if(!$SuppressHeader.IsPresent)
		{
			$OutputHeader = "Web App,Site Collection,Scope,Name,URL,Most Recent Day with Usage,Hits - All Time,Unique Users - All Time,Hits - Most recent Day with Usage,Unique Users - Most recent day with Usage,Current Date,Monthly Hits,Daily Hits,Month,Day,Size (GB)"
			$OutputHeader | Out-File $OutputFilepath -Append
		}

		# Get Web App for root site
		$RootSite = Get-SPSite $RootSiteUrl -ErrorAction SilentlyContinue
		if(!$RootSite)
		{
			$abort = $true
			Write-Host "Cannot find Site!" -ForegroundColor Red -BackgroundColor White
		}
		else
		{
			$WebApp = $RootSite.WebApplication
		}

        #$webApplications = Get-SPWebApplication

		# Get Search Service Applications
		$inFarm = $null
        $ssaEndPoint =  ([System.web.httputility]::UrlDecode((Get-SPServiceApplicationProxy | ?{ $_.TypeName -like "Search*" } | select ServiceEndpointUri).ServiceEndpointUri.absoluteuri)).Split("=")[2].split("/")[2].split(":")[0]
        if(Get-SPServer | ?{ $_.Address -eq $ssaEndPoint })
        {
            Write-Verbose "In farm"
            $inFarm = $true
            $SearchApp = Get-SPEnterpriseSearchServiceApplication # assumes only one SSA
	        if(!$SearchApp)
	        {
		        $abort = $true
		        Write-Host "Cannot find SSA!" -ForegroundColor Red -BackgroundColor White
	        }
        }
        else
        {
            Write-Verbose "Not in farm"
            $inFarm = $false
            $Session = New-PSSession -ComputerName $ssaEndPoint
            if(!$Session)
	        {
		        $abort = $true
		        Write-Host "Cannot connect to remote SSA endpoint!" -ForegroundColor Red -BackgroundColor White
	        }
            else
            {
                Invoke-Command -Session $Session -ScriptBlock { if((Get-PSSnapin "Microsoft.SharePoint.PowerShell" -ErrorAction SilentlyContinue) -eq $null){ Add-PSSnapin "Microsoft.SharePoint.PowerShell" } }
            }
        }
	}

	process
	{
		if($abort){ return "Aborted" }
		# Loop thru all Site Collections in Web App
		foreach($Site in $WebApp.Sites)
		{
			# Site Collection title
			$SiteColTitle = $Site.RootWeb.Title.Replace(",","") # remove commas from title, since comma used as delimiter
			# Export site analytics if -IncludeSites
			if($IncludeSites.IsPresent)
			{
				$Scope = "Site"
				$SiteTitle = $Site.RootWeb.Title.Replace(",","") # remove commas from title, since comma used as delimiter
				$SiteUrl = $Site.Url
				$UsageData = $SearchApp.GetRollupAnalyticsItemData(1,[System.Guid]::Empty,$Site.ID,[System.Guid]::Empty)
				$LastProcessingTime = $UsageData.LastProcessingTime
				$CurrentDate = $UsageData.CurrentDate
				$TotalHits = $UsageData.TotalHits
				$TotalUniqueUsers = $UsageData.TotalUniqueUsers
				$LastProcessingHits = $UsageData.LastProcessingHits
				$LastProcessingUniqueUsers = $Usage.LastProcessingUniqueUsers
				try{ $HitCountforMonth = $UsageData.GetHitCountforMonth($tDate) } catch{ $HitCountforMonth = "N/A" }
				try{ $HitCountforDay = $UsageData.GetHitCountforDay($tDate) } catch{ $HitCountforDay = "N/A" }
                
                $size = [math]::Round($site.Usage.Storage/1GB,4)

				# Write data to file
				$OutputString = "$($WebApp.Name),$($SiteColTitle),$($Scope),$($SiteTitle),$($SiteUrl),$($LastProcessingTime),$($TotalHits),$($TotalUniqueUsers),$($LastProcessingHits),$($LastProcessingUniqueUsers),$($CurrentDate),$($HitCountforMonth),$($HitCountforDay),$($tDate.Year.ToString('0000'))$($tDate.Month.ToString('00')),$($tDate.ToString('d')),$($size)"
				$OutputString | Out-File $OutputFilepath -Append
			}

			# Export web analytics if -IncludeWebs
			if($IncludeWebs.IsPresent)
			{
				$Webs = $Site.OpenWeb().GetSubwebsForCurrentUser()
				foreach($Web in $Webs)
				{
					$Scope = "Web"
					$SiteTitle = $Web.Title.Replace(",","") # remove commas from title, since comma used as delimiter
					$SiteUrl = $Web.Url
					$UsageData = $SearchApp.GetRollupAnalyticsItemData(1,[System.Guid]::Empty,$Site.ID,$Web.ID)
					$LastProcessingTime = $UsageData.LastProcessingTime
					$CurrentDate = $UsageData.CurrentDate
					$TotalHits = $UsageData.TotalHits
					$TotalUniqueUsers = $UsageData.TotalUniqueUsers
					$LastProcessingHits = $UsageData.LastProcessingHits
					$LastProcessingUniqueUsers = $Usage.LastProcessingUniqueUsers

					# Write data to file
					$OutputString = "$($WebApp.Name),$($SiteColTitle),$($Scope),$($SiteTitle),$($SiteUrl),$($LastProcessingTime),$($TotalHits),$($TotalUniqueUsers),$($LastProcessingHits),$($LastProcessingUniqueUsers),$($CurrentDate),$($HitCountforMonth),$($HitCountforDay),$($tDate.Year.ToString("0000"))$($tDate.Month.ToString("00")),$($tDate.ToString("d"))"
					$OutputString | Out-File $OutputFilepath -Append
				}
			}
		}
	}

	end
	{
		$Site.Dispose()
	}
}
