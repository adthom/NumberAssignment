$SharePointUrl = Get-AutomationVariable -Name SharePointUrl

$Credentials = Get-AutomationPSCredential -Name "PhoneNumberListAdmin"
Connect-PnPOnline -Url $SharePointUrl -Credentials $Credentials
Connect-MicrosoftTeams -Credential $Credentials

# only target TeamsOnly provisioned users with valid PhoneSystem License assigned
$Users = @(Get-CsOnlineUser -Filter { Enabled -eq $true -and VoicePolicy -eq 'HybridVoice' -and OnPremLineUriManuallySet -eq $false -and LineUri -eq $null } | Where-Object { $_.InterpretedUserType.EndsWith('TeamsOnlyUser') })
# $Users += Get-CsOnlineUser -Filter { Enabled -eq $true -and VoicePolicy -eq 'BusinessVoice' -and LineUri -eq $null }

Write-Output "$($Users.Count) users found without assigned numbers."

$List = Get-PnPList -Identity PhoneNumberManagement

foreach ($User in $Users) {
    $UserId = $User.UserPrincipalName
    $Country = $User.UsageLocation
    $NumberType = if ($User.VoicePolicy -eq "HybridVoice") { "DirectRouting" } # else { "CallingPlan" }

    $PnPUser = Get-PnPUser -Identity "i:0#.f|membership|$($UserId)"
    if ($null -ne $PnPUser) {
        $CAML = @"
        <View>
          <Query>
            <Where>
              <And>
                <Eq>
                  <FieldRef Name='AssignedTo'/>
                  <Value Type='Text'>$($PnPUser.Title)</Value>
                </Eq>
                <Eq>
                  <FieldRef Name='AvailableToAssign'/>
                  <Value Type='Boolean'>0</Value>
                </Eq>
              </And>
            </Where>
          </Query>
        </View>
"@
        $Assigned = Get-PnPListItem -List $List -Query $CAML -PageSize 1 | Select-Object -First 1
    }
    else {
        $Assigned = $null
    }

    if ($null -ne $Assigned -and ($null -eq $Assigned["NumberType"] -or $Assigned["NumberType"] -ne $NumberType)) {
        # invalid number type, clear from list, get new number
        Set-PnPListItem -List $List -Identity $Assigned -Values @{ 'AssignedTo' = ""; 'AvailableToAssign' = $true } | Out-Null
        $Assigned = $null
    }

    if ($null -eq $Assigned) {
        $CAML = @"
<View>
  <Query>
    <Where>
      <And>
        <And>
          <IsNull>
            <FieldRef Name='AssignedTo'/>
          </IsNull>
          <Eq>
            <FieldRef Name='AvailableToAssign'/>
            <Value Type='Boolean'>1</Value>
          </Eq>
        </And>
        <And>
          <Eq>
            <FieldRef Name='NumberType'/>
            <Value Type='Text'>$NumberType</Value>
          </Eq>
          <Eq>
            <FieldRef Name='Country'/>
            <Value Type='Text'>$Country</Value>
          </Eq>
        </And>
      </And>
    </Where>
  </Query>
</View>
"@
        $Unassigned = Get-PnPListItem -List $List -Query $CAML -PageSize 1 | Select-Object -First 1
        if ($null -ne $Unassigned) {
            Set-PnPListItem -List $List -Identity $Unassigned -Values @{ 'AssignedTo' = $UserId; 'AvailableToAssign' = $false } | Out-Null
            $Assigned = $Unassigned
        }
    }

    if ($null -ne $Assigned) {
        $EmailParams = @{
            SmtpServer                 = 'smtp.office365.com'
            Port                       = '587'
            UseSSL                     = $true
            Credential                 = $Credentials
            From                       = $Credentials.UserName
            To                         = $User.WindowsEmailAddress
            Subject                    = "Your new Phone Number has been assigned"
            Body                       = "$($User.DisplayName)-`r`nYou have been assigned a new phonenumber in Microsoft Teams.`r`n`r`nYour new number is now $($Assigned['Title'])."
            DeliveryNotificationOption = 'OnFailure', 'OnSuccess'
        }

        if ($NumberType -eq "DirectRouting") {
            $Number = "tel:" + $Assigned['Title']
            Set-CsUser -Identity $UserId -OnPremLineURI $Number -EnterpriseVoiceEnabled $true -HostedVoiceMail $true
            Send-MailMessage @EmailParams
            Write-Output "$($Assigned['Title']) assigned to $UserId"
        } 
        # else {
        #   # need to set location for CallingPlan
        #   $Location = Get-CsOnlineLisLocation | Select-Object -First 1
        #   Set-CsOnlineVoiceUser -Identity $UserId -TelephoneNumber $Assigned['Title'] -LocationID $Location.LocationId
        #   Send-MailMessage @EmailParams
        #   Write-Output "$($Assigned['Title']) assigned to $UserId"
        # }
    }
    else {
        Write-Warning "No Available numbers for $NumberType in $Country, unable to assign number for $UserId!"
    }
}

Disconnect-PnPOnline
Disconnect-MicrosoftTeams