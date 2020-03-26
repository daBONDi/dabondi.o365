#!powershell

# Copyright: (c) 2020 (David Baumann(@daBONDi) <me@davidbaumann.at>
# GNU General Public License v3.0+ (see COPYING or https://www.gnu.org/licenses/gpl-3.0.txt)

#AnsibleRequires -CSharpUtil Ansible.Basic

$spec = @{
  options = @{
      DisplayName = @{ type = "str"; required = $true }
      Members = @{ type="list"; required = $false; elements="str"; default=@()  }
      Owners = @{ type="list"; required =$false; elements="str"; default=@()}
      o365_admin_username = @{ type="str"; required = $true; no_log = $true}
      o365_admin_password = @{ type="str"; required = $true; no_log = $true}
      MSOGroupMembers = @{ type="list"; required=$false; elements="str"; default=@() }
      MSOGroupOwners = @{ type="list"; required=$false; elements="str"; default=@() }
  }
  supports_check_mode = $true
  no_log = $true
}
$module = [Ansible.Basic.AnsibleModule]::Create($args, $spec)

# Push Params to Vars
$team_name = $module.Params.DisplayName
$members = $module.Params.Members
$owners = $module.params.Owners

# Create Credentials
$o365_admin_username = $module.Params.o365_admin_username
$o365_admin_password = $module.Params.o365_admin_password
$secpasswd = ConvertTo-SecureString $o365_admin_password -AsPlainText -Force
$o365_admin_creds = New-Object System.Management.Automation.PSCredential ($o365_admin_username, $secpasswd)

try{
  Import-Module "MicrosoftTeams" -RequiredVersion 1.0.21
}catch{
  $module.FailJson("Powershell Module MicrosoftTeams missing")
}

Try{
  Connect-MicrosoftTeams -Credential $o365_admin_creds | Out-Null
}catch{
  $module.FailJson("Error on connecting to Teams $_.Exception.Message",$_)
}

$team_object = Get-Team -Displayname $team_name
if($team_object -is [string])
{
  $module.FailJson($team_object)
}

if($team_object -is [array]){
  $module.FailJson("We found Multiple Teams!")
}

# =========================== Initalization Finished

$current_owners = Get-Teamuser -GroupId $team_object.GroupId -Role Owner
$current_members = Get-Teamuser -GroupId $team_object.GroupId -Role Member

function unwrap_teams_user_list($List){
  $result = @()
  foreach($item in $List){
    $result += $item | Select-Object -ExpandProperty User
  }
  return $result
}

function build_list_work_items($CurrentList, $DesiredList){
  # See if we need to Add someone
  $add_list = @()
  $remove_list = @()
  # The Desired List is Empty
  if(-not $DesiredList){
    $remove_list = $CurrentList
    $diff = "Remove all from CurrentList"
  }elseif(-not $CurrentList){
    $add_list = $DesiredList
    $remove_list = @()
    $diff = "Add all from DesiredList"
  }else{
    # We need to Search Difference
    $diff = Compare-Object -ReferenceObject $DesiredList -DifferenceObject $CurrentList

    foreach($item in $diff){
      if($item.SideIndicator -eq "<="){ $add_list += $item.InputObject }
      if($item.SideIndicator -eq "=>"){ $remove_list += $item.InputObject }
    }
  }
  return $add_list, $remove_list, $diff
}

function RemoveUsersFromTeam($Role, $List,$Team){
  $module.Result.changed = $true
  if(-not $module.CheckMode){
    # We do something here
    foreach($item in $List){
      if($Role -eq "Owner"){
        Remove-TeamUser -GroupId $($Team.GroupId) -User $Item -Role $Role | Out-Null
      }else{
        Remove-TeamUser -GroupId $($Team.GroupId) -User $Item -Role $Role | Out-Null
      }
    }
  }
}

function AddUsersToTeam($Role, $List,$Team){
  $module.Result.changed = $true
  if(-not $module.CheckMode){
    # We do something here
    foreach($item in $list){
      try{
        Add-TeamUser -GroupId $($Team.GroupId) -User $Item -Role $Role
      }catch{
        $module.FailJson("Error on Adding Team Member $item from $Role : $($_.Exception.Message)")
      }
    }
  }
}

function Get-MSolGroupId ($GroupName){
  $msol_group_id = Get-MsolGroup -All | Where-Object DisplayName -eq $GroupName | Select-Object -ExpandProperty ObjectId
    if(-not $msol_group_id){
      $module.FailJson( ("Could not find targeted MSOnline Group {0}" -f $GroupName))
    }
  return $msol_group_id
}

function Get-MSolGroupMembers($GroupNameList){
  $result = @()
  foreach($grp in $GroupNameList){
    $group_id = Get-MsolGroupId -GroupName $grp
    $group_members = Get-MsolGroupMember -GroupObjectId $group_id | Select-Object -ExpandProperty EmailAddress | Foreach-Object { $_.ToString() }
    $result +=$group_members
  }
  return $result
}

function  Get-MSOnlineMembers($MemberGroups, $OwnerGroups, $O365Credentials){
  $mso_owners = @()
  $mso_members = @()

  try{
    Import-Module MSOnline
  }catch{
    $module.FailJson("Cannot Import Powershell Module MSOnline, is it Missing on execution system?")
  }
  
  try{
    Connect-MsolService -Credential $O365Credentials
  }catch{
    $module.FailJson("Error on Connection to MSOnline Service")
  }

  $mso_owners =  Get-MSolGroupMembers -GroupNameList $OwnerGroups
  $mso_members = Get-MsolGroupMembers -GroupNameList $MemberGroups

  return  $mso_members, $mso_owners
}


$module.Result.CurrentOwners = $current_owners
$module.Result.CurrentMembers = $current_members

$module.Result.CurrentOwners_unwarp = unwrap_teams_user_list -List $current_owners
$module.Result.CurrentMembers_unwarp = unwrap_teams_user_list -List $current_members

$add_owners = @()
$remove_owners = @()
$add_members = @()
$remove_members = @()

# Add Owners/Members from MSOnline Group
if($module.Params.MSOGroupMembers -or $module.Params.MSOGroupOwners){
  
  $mso_members,$mso_owners = Get-MSOnlineMembers -MemberGroups $module.Params.MSOGroupMembers -OwnerGroups $module.Params.MSOGroupOwners -O365Credentials $o365_admin_creds
  if($mso_owners){
    $owners += $mso_owners
  }
  if($mso_members){
    $members += $mso_members
  }
  $module.Result.mso_owners = $mso_owners
  $module.Result.mso_members = $mso_members
}

$module.Result.desiredOwners = $owners
$module.Result.desiredMembers = $members

$add_owners, $remove_owners,$diff_owner = build_list_work_items -CurrentList $(unwrap_teams_user_list -List $current_owners) -DesiredList $owners
$add_members, $remove_members,$diff_member = build_list_work_items -CurrentList $(unwrap_teams_user_list -List $current_members) -DesiredList $members

$module.Result.owner_add = $add_owners
$module.Result.owner_remove = $remove_owners
$module.Result.owner_diff = $diff_owner
$module.Result.member_add = $add_members
$module.Result.member_remove = $remove_members
$module.Result.member_diff = $diff_member


# Now we got our Lists lets do it
if($add_owners.Count -gt 0){
  AddUsersToTeam -Role "Owner" -List $add_owners -Team $team_object
  
}
if($remove_owners.Count -gt 0){
  RemoveUsersFromTeam -Role "Owner" -List $remove_owners -Team $team_object
}

if($add_members.Count -gt 0){
  AddUsersToTeam -Role "Member" -List $add_members -Team $team_object
}
if($remove_members.Count -gt 0){
  RemoveUsersFromTeam -Role "Member" -List $remove_members -Team $team_object
}

try{
  Disconnect-MicrosoftTeams | Out-Null
}catch{
  $module.FailJson("Error on disconnecting from Teams $_.Exception.Message")
}

$module.ExitJson()
