# Delete messages from mailbox

# GLOBALS
$global:SEARCH_NAME = ""

function print-helpmenu{
        Write-Host "?/help:`t`t`tPrint this menu."
        Write-Host "exit:`t`t`tExit this application."
        Write-Host "new-search:`t`tStart a new search. You will be guided through setting up a search to execute."
        <#Write-Host "get-search:`t`tGets results of current search session."#>
        Write-Host "preview-results:`t`tPreviews results of the search. Outputs in JSON blob."
        Write-Host "purge-emails:`t`tPurges the results of the current search."
        Write-Host "remove-search:`t`tRemoves the current search."
        Write-Host "set-search:`t`tIf you already know the name of your search, you can use this command to continue where you left off from a previous session."
}

function new-search{
        Write-Host "Starting new search...`n"
        $searchname = Read-Host "Enter a name for your search"
        $exchangelocation = Read-Host "Enter email address(es) (comma-separated) to search"
        $query = Read-Host "Query"

        Write-Host "Building command: New-ComplianceSearch -Name `"$($searchname)`" -ExchangeLocation $($exchangelocation) -ContentMatchQuery `"$($query)`""
        $yn = Read-Host "Does this look correct? (y/n)"
        if($yn -eq "y"){
                New-ComplianceSearch -Name "$($searchname)" -ExchangeLocation $($exchangelocation) -ContentMatchQuery "$($query)"
                Write-Host "`nStarting search: `"$($searchname)`""
                Start-ComplianceSearch -Identity "$($searchname)"
                $global:SEARCH_NAME = $searchname
        }
        else{
                Write-Host "Cancelling search builder. Exiting to main menu."
        }
}

function set-search{
        $global:SEARCH_NAME = "$(Read-Host "Enter the name of your search")"
        Write-Host "Search set to: $($global:SEARCH_NAME)"
}
<# TODO: Fix output
function get-search{
        Write-Host "Retrieving latest search...`n"
        Get-ComplianceSearch -Identity "$($global:SEARCH_NAME)"
        # To be implemented once script is working
        $num_results = Get-ComplianceSearch -Identity "$($global:SEARCH_NAME)" | Format-List -Property Items
        Write-Host "Number of items: $($num_results)"
}#>

function preview-results{
        New-ComplianceSearchAction -SearchName "$($global:SEARCH_NAME)" -Preview
        Get-ComplianceSearchAction -Identity "$($global:SEARCH_NAME)_Preview" | Format-List -Property Results
}

# TODO: Write-host used for testing purposes.
function purge-emails{
        Write-Host "New-ComplianceSearchAction -SearchName `"$($global:SEARCH_NAME)`" -Purge -PurgeType SoftDelete"
        Write-Host "Get-ComplianceSearchAction -Identity `"$($global:SEARCH_NAME)_Purge"
        Write-Host "Get-ComplianceSearchAction -Identity `"$($global:SEARCH_NAME)_Purge`" | Format-List -Property Results"
}

function remove-search{
        Write-Host "Deleting compliance search!"
        Remove-ComplianceSearch -Identity "$($SEARCH_NAME)"
        $global:SEARCH_NAME = ""
}

Write-Host "Welcome to Mail Sniper!"
Write-Host "Please type '?' or 'help' for more a list of options`n"
Write-Host "Connecting to the eDiscovery module..."
try{
        Connect-Ediscovery
}
catch [System.Management.Automation.RuntimeException] {
        Write-Host "Authentication cancelled. Exiting!`n";
        Exit;
}
Write-Host "Connected!"

while(1 -eq 1){
        $in = Read-Host "$($global:SEARCH_NAME)>"
        if($in -eq "exit"){exit;}
        elseif($in -eq "query help"){start-process "https://docs.microsoft.com/en-us/exchange/client-developer/web-service-reference/querystring-querystringtype?redirectedfrom=MSDN";}
        elseif(($in -eq "help") -or ($in -eq "?")){print-helpmenu;}
        elseif($in -eq "new-search"){new-search;}
        <#elseif($in -eq "get-search"){get-search;}#>
        elseif($in -eq "preview-results"){preview-results;}
        elseif($in -eq "purge-emails"){purge-emails;} 
        elseif($in -eq "remove-search"){remove-search;} 
        elseif($in -eq "set-search"){set-search;} 
        else{write-host "Please type an appropriate command. Use '?'/'help' for a list of commands.";}
}