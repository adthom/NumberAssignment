$SharePointUrl = Get-AutomationVariable -Name SharePointUrl

$Credentials = Get-AutomationPSCredential -Name "PhoneNumberListAdmin"
Connect-PnPOnline -Url $SharePointUrl -Credentials $Credentials
Connect-MicrosoftTeams -Credential $Credentials

$List = Get-PnPList -Identity PhoneNumberManagement
$CAML = @"
<View>
  <Query>
    <Where>
      <And>
        <IsNotNull>
          <FieldRef Name='AssignedTo'/>
        </IsNotNull>
        <Eq>
          <FieldRef Name='AvailableToAssign'/>
          <Value Type='Boolean'>1</Value>
        </Eq>
      </And>
    </Where>
  </Query>
</View>
"@
# Check if user has number assigned in the list
$NumbersToClear = @(Get-PnPListItem -List $List -Query $CAML)

Write-Output "$($NumbersToClear.Count) numbers found that need to be removed from users."

foreach ($Number in $NumbersToClear) {
    $User = Get-CsOnlineUser -Filter { LineUri -eq "tel:$($Number['Title'])" -or LineUri -eq "$($Number['Title'])" } | Select-Object -First 1
    if ($null -eq $User) {
        $User = Get-CsOnlineUser -Identity $Number['AssignedTo'].Email
    }
    if ($null -ne $User -or ($User.LineUri -ne "tel:$($Number['Title'])" -and $User.LineUri -ne "$($Number['Title'])")) {
        Write-Output "$($Number['Title']) is currently assigned to $($User.UserPrincipalName)"
        $Updated = $false
        if ($User.VoicePolicy -eq "BusinessVoice") {
            Set-CsOnlineVoiceUser -Identity $User.Identity -TelephoneNumber $null
            $Updated = $true
        }
        elseif ($User.OnPremLineUriManuallySet) {
            Set-CsUser -Identity $User.Identity -OnPremLineURI $null
            $Updated = $true
        }
        else {
            Write-Warning "$($User.UserPrincpalName) has a number assigned in on-premises Active Directory, cannot update here!"
        }

        if ($Updated) {
            Set-PnPListItem -List $List -Identity $Number -Values @{ 'AssignedTo' = "" } | Out-Null
            Write-Output "Number cleared from $($User.UserPrincipalName)"
        }
    }
    else {
        Write-Warning "$($Number['AssignedTo'].Email) is not actually assigned number $($Number['Title']), clearing from list."
        Set-PnPListItem -List $List -Identity $Number -Values @{ 'AssignedTo' = "" } | Out-Null
    }
}

Disconnect-PnPOnline
Disconnect-MicrosoftTeams