Param (
    [Parameter(Mandatory = $false)][string]$UpdatesLaterThan = "2022-07"
)

$requiredUpdatesList = New-Object System.Collections.ArrayList

function addItemToRequiredUpdatesList {
    Param (
	[Parameter(Mandatory = $true)][string]$UpdateTitleEn,
	[Parameter(Mandatory = $true)][string]$KBNumber,
        [Parameter(Mandatory = $false)][boolean]$Preview = $false, #This is for future use if need arise
        [Parameter(Mandatory = $false)][boolean]$OutOfBand = $false #This is for future use if need arise
    )
    
    $updateDate = $UpdateTitleEn.Split("")[0]

    if ($updateDate -gt $UpdatesLaterThan) {
        $patchInfo = New-Object System.Object
        $patchInfo | Add-Member -MemberType NoteProperty -Name 'UpdateTitleEn' -Value $UpdateTitleEn
        $patchInfo | Add-Member -MemberType NoteProperty -Name 'KBNumber' -Value $KBNumber
        $patchInfo | Add-Member -MemberType NoteProperty -Name 'Date' -Value $updateDate
        $patchInfo | Add-Member -MemberType NoteProperty -Name 'Preview' -Value $Preview
        $patchInfo | Add-Member -MemberType NoteProperty -Name 'OutOfBand' -Value $OutOfBand
        $requiredUpdatesList.add($patchInfo) > $null
    }
}

function populateListRequiredUpdates {
    #2022-08
    addItemToRequiredUpdatesList -UpdateTitleEn "2022-08 Cumulative Update for Windows 10" -KBNumber "KB5016616"
    addItemToRequiredUpdatesList -UpdateTitleEn "2022-08 Cumulative Update Preview for Windows 10" -KBNumber "KB5016688" -Preview:$true
    #2022-09
    addItemToRequiredUpdatesList -UpdateTitleEn "2022-09 Cumulative Update for Windows 10" -KBNumber "KB5017308"
    addItemToRequiredUpdatesList -UpdateTitleEn "2022-09 Cumulative Update Preview for Windows 10" -KBNumber "KB5017380" -Preview:$true
    #2022-10
    addItemToRequiredUpdatesList -UpdateTitleEn "2022-10 Cumulative Update for Windows 10" -KBNumber "KB5018410"
    addItemToRequiredUpdatesList -UpdateTitleEn "2022-10 Cumulative Update for Windows 10" -KBNumber "KB5020435" -OutOfBand:$true
    addItemToRequiredUpdatesList -UpdateTitleEn "2022-10 Cumulative Update Preview for Windows 10" -KBNumber "KB5018482" -Preview:$true
    addItemToRequiredUpdatesList -UpdateTitleEn "2022-10 Cumulative Update for Windows 10" -KBNumber "KB5020953" -OutOfBand:$true
    #2022-11
    addItemToRequiredUpdatesList -UpdateTitleEn "2022-11 Cumulative Update for Windows 10" -KBNumber "KB5019959"
    addItemToRequiredUpdatesList -UpdateTitleEn "2022-11 Cumulative Update Preview for Windows 10" -KBNumber "KB5020030" -Preview:$true
    #2022-12
    addItemToRequiredUpdatesList -UpdateTitleEn "2022-12 Cumulative Update for Windows 10" -KBNumber "KB5021233"
    #2023-01
    addItemToRequiredUpdatesList -UpdateTitleEn "2023-01 Cumulative Update for Windows 10" -KBNumber "KB5022282"
    addItemToRequiredUpdatesList -UpdateTitleEn "2023-01 Cumulative Update Preview for Windows 10" -KBNumber "KB5019275" -Preview:$true
    #2023-02
    addItemToRequiredUpdatesList -UpdateTitleEn "2023-02 Cumulative Update for Windows 10" -KBNumber "KB5022834"
    addItemToRequiredUpdatesList -UpdateTitleEn "2023-02 Cumulative Update Preview for Windows 10" -KBNumber "KB5022906" -Preview:$true
    #2023-03
    addItemToRequiredUpdatesList -UpdateTitleEn "2023-03 Cumulative Update for Windows 10" -KBNumber "KB5023696"
    addItemToRequiredUpdatesList -UpdateTitleEn "2023-03 Cumulative Update Preview for Windows 10" -KBNumber "KB5023773" -Preview:$true
    #2023-04
    addItemToRequiredUpdatesList -UpdateTitleEn "2023-04 Cumulative Update for Windows 10" -KBNumber "KB5025221"
    #2023-05
    addItemToRequiredUpdatesList -UpdateTitleEn "2023-05 Cumulative Update for Windows 10" -KBNumber "KB5026361"
    #2023-06
    addItemToRequiredUpdatesList -UpdateTitleEn "2023-06 Cumulative Update for Windows 10" -KBNumber "KB5027215"
}

#"Borrowed" from Timothy Tibbetts:
#https://www.majorgeeks.com/content/page/how_to_check_your_windows_update_history_with_powershell.html
function Get-WuaHistory {
    $session = (New-Object -ComObject 'Microsoft.Update.Session')
    $history = $session.QueryHistory("",0,1000) | ForEach-Object {
        if ($_.ResultCode -eq 2) { #Result code 2 is installed updates without errors

            $_ | Add-Member -MemberType NoteProperty -Name Result -Value "Installed"
	    $_ | Add-Member -MemberType NoteProperty -Name UpdateId -Value $_.UpdateIdentity.UpdateId
	    $_ | Add-Member -MemberType NoteProperty -Name RevisionNumber -Value $_.UpdateIdentity.RevisionNumber
	    $_ | Add-Member -MemberType NoteProperty -Name Product -Value $Product -PassThru
        }
    }

    #Remove null records and only return the fields we want
    $history |
    Where-Object {![String]::IsNullOrWhiteSpace($_.title)} |
    Select-Object Result, Date, Title, SupportUrl, UpdateId, RevisionNumber
}


populateListRequiredUpdates

#Make KBs into regex
$regexKB = $requiredUpdatesList.KBNumber -join '.*|.*'
$regexKB = ".*" + $regexKB + ".*"

$updateList = Get-WuaHistory
$foundRequiredUpdatesInstalled = $updateList | where { $_.Title | ? { $_ -match $regexKB } }

#return $requiredUpdatesList
return $foundRequiredUpdatesInstalled
