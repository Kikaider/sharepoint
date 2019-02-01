<#
.SYNOPSIS
Adds a new Sinlge Line of Text field as a Site Column, then adds the new column to each end user List in the Site Collection.

#>

$SiteCollectionURL = "http://hej.lab.roblab.com"
$AddToDefaultView = $true
$ColumnName = "NewColumnOne"
$ColumnDescription = ""
$ColumnDisplayName = ""

# load PowerShell add-in for SharePoint
$ver = $host | select version
if($ver.Version.Major -gt 1){ $host.Runspace.ThreadOptions = "ReuseThread" }
if($null -eq (Get-PSSnapin "Microsoft.SharePoint.PowerShell" -ErrorAction SilentlyContinue)){ Add-PSSnapin "Microsoft.SharePoint.PowerShell" }

function Add-SiteColumn
{
    [cmdletbinding(SupportsShouldProcess=$True)]

    param(
        [Parameter(Mandatory=$true)]
        [ValidateNotNullOrEmpty()]
        [string]$ColumnName,

        [Parameter(Mandatory=$false)]
        [string]$ColumnDescription,

        [Parameter(Mandatory=$false)]
        [string]$ColumnDisplayName,

        [Parameter(Mandatory=$false)]
        [string]$ColumnGroup = "Custom",

        [Parameter(Mandatory=$true)]
        [ValidateNotNullOrEmpty()]
        [string]$SiteCollectionURL,

        [Parameter(Mandatory=$false)]
        [Switch]$Hidden,

        [Parameter(Mandatory=$false)]
        [Switch]$Required,

        [Parameter(Mandatory=$false)]
        [Switch]$DoNotShowInDisplayForm,

        [Parameter(Mandatory=$false)]
        [Switch]$DoNotShowInEditForm,

        [Parameter(Mandatory=$false)]
        [Switch]$DoNotShowInListSettings,

        [Parameter(Mandatory=$false)]
        [Switch]$DoNotShowInNewForm
    )

    Begin
    {
        if($null -eq (Get-PSSnapin "Microsoft.SharePoint.PowerShell" -ErrorAction SilentlyContinue)){ Add-PSSnapin "Microsoft.SharePoint.PowerShell" }
        #$spVersion = (Get-SPFarm).BuildVersion
        try {
            $site = Get-SPSite -Identity $SiteCollectionURL
        }
        catch {
            Write-Error -Message "Unable to find Site Collection specified." -RecommendedAction "Please verify the -SiteCollectionURL parameter and try again."
            return $false
        }
        $site = Get-SPSite -Identity $SiteCollectionURL
        If($site)
        {
            $web = $site.RootWeb
        }
        else {
            Write-Error -Message "Unable to find Site Collection specified." -RecommendedAction "Please verify the -SiteCollectionURL parameter and try again."
            return $false
        }
        
        if($ColumnDisplayName){$cColumnDisplayName = $ColumnDisplayName}else{$cColumnDisplayName = $ColumnName}
        if($Hidden){$cHidden = "TRUE"}else{$cHidden = "FALSE"}
        if($Required){$cRequired = "TRUE"}else{$cRequired = "FALSE"}
        if($DoNotShowInDisplayForm){$cDoNotShowInDisplayForm = "TRUE"}else{$cDoNotShowInDisplayForm = "FALSE"}
        if($DoNotShowInEditForm){$cDoNotShowInEditForm = "TRUE"}else{$cDoNotShowInEditFormd = "FALSE"}
        if($DoNotShowInListSettings){$cDoNotShowInListSettings = "TRUE"}else{$cDoNotShowInListSettings = "FALSE"}
        if($DoNotShowInNewForm){$cDoNotShowInNewForm = "TRUE"}else{$cDoNotShowInNewForm = "FALSE"}
    }

    Process
    {
        
        $fieldXML = "<Field Type='Text'
        Name='$($ColumnName)'
        Description='$($ColumnDescription)'
        DisplayName='$($cColumnDisplayName)'
        Group='$($ColumnGroup)'
        Hidden='$($cHidden)'
        Required='$($cRequired)'
        ShowInDisplayForm='$($cDoNotShowInDisplayForm)'
        ShowInEditForm='$($cDoNotShowInEditForm)'
        ShowInListSettings='$($cDoNotShowInListSettings)'
        ShowInNewForm='$($cDoNotShowInNewForm)'></Field>"
        
        try {
            $web.Fields.AddFieldAsXml($fieldXML)
            Write-Verbose "Successfullyy added Site Column $($ColumnName) to $($SiteCollectionURL)."
            return $True
        }
        catch {
            Write-Error -Message "Unable to add Site Column." -RecommendedAction "Please verify the parameters and try again."
            return $false
        }
    }

    End
    {
        $web.Dispose()
        $site.Dispose()
    }
}

Function Add-SiteColumnToList
{
    [cmdletbinding(SupportsShouldProcess=$True)]

    param(
        [Parameter(Mandatory=$true)]
        [ValidateNotNullOrEmpty()]
        [string]$SiteURL,    
        
        [Parameter(Mandatory=$true)]
        [ValidateNotNullOrEmpty()]
        [string]$ListName,

        [Parameter(Mandatory=$true)]
        [ValidateNotNullOrEmpty()]
        [string]$ColumnName,

        [Parameter(Mandatory=$false)]
        [string]$FieldType = "Text",

        [Parameter(Mandatory=$false)]
        [bool]$IsRequired = $False,

        [Parameter(Mandatory=$false)]
        [Switch]$AddToDefaultView
    )

    begin{
        $ErrorActionPreference = "Stop"
    }
    
    process{
        Try{
            $List = (Get-SPWeb $SiteURL).Lists.TryGetList($ListName)
            
            if($null -ne $List)
            {
                if(!$List.Fields.ContainsField($ColumnName))
                {      
                    $List.Fields.Add($ColumnName,$FieldType,$IsRequired)
                    $List.Update()
    
                    #Update the default view to include the new column
                    if($AddToDefaultView){
                        $View = $List.DefaultView
                        $View.ViewFields.Add($ColumnName)
                        $View.Update()
                    }
                    Write-Host "Column '$ColumnName' Added to the List!" -ForegroundColor Green
                }
                else
                {
                    Write-Host "Column '$ColumnName' already Exists in the List" -ForegroundColor Red
                }
            }
            else
            {
                Write-Host "List '$ColumnName' doesn't exists!" -ForegroundColor Red
            }        
        }
        catch {
            Write-Host $_.Exception.Message -ForegroundColor Red
        }
    }

    end
    {
        $ErrorActionPreference = "Continue"
    }
}

#add new Site Columns to Site Collection
$c1 = Add-SiteColumn -ColumnName $ColumnName -SiteCollectionURL $SiteCollectionURL

#add new Site Columns to each List in Site Collection
if($c1)
{
   $lists = Get-SPSite $SiteCollectionURL -Limit All |
    Select -ExpandProperty AllWebs |
    Select -ExpandProperty Lists |
    Where { -not $_.hidden -and $_.AllowDeletion -and -not $_.IsApplicationList } |
    Select ParentWebUrl, Title
   
   ForEach($list in $lists)
   {
        if($AddToDefaultView)
        {
            Add-SiteColumnToList -SiteURL ($SiteCollectionURL + $list.ParentWebUrl) -ListName $list.Title -ColumnName $ColumnName -AddToDefaultView
        }
        else
        {
            Add-SiteColumnToList -SiteURL ($SiteCollectionURL + $list.ParentWebUrl) -ListName $list.Title -ColumnName $ColumnName
        }
   }
}
