#!powershell

# Copyright: (c) 2020 (David Baumann(@daBONDi) <me@davidbaumann.at>
# GNU General Public License v3.0+ (see COPYING or https://www.gnu.org/licenses/gpl-3.0.txt)

#AnsibleRequires -CSharpUtil Ansible.Basic

$spec = @{
  options = @{
      MsolGroup = @{ type = "str"; required = $true }
      MsolMembershipGroups = @{ type="list"; required = $false; elements="str"; default=@()}
      MsolMembers = @{ type="list"; elements="str"; required=$false; default=@()}
      AddOnly = @{ type="bool"; required=$false; default=$false}
      o365_admin_username = @{ type="str"; required = $true; no_log = $true}
      o365_admin_password = @{ type="str"; required = $true; no_log = $true}
  }
  supports_check_mode = $true
  no_log = $true
}
$module = [Ansible.Basic.AnsibleModule]::Create($args, $spec)

$o365_admin_username = $module.Params.o365_admin_username
$o365_admin_password = $module.Params.o365_admin_password

$secpasswd = ConvertTo-SecureString $o365_admin_password -AsPlainText -Force
$o365_admin_creds = New-Object System.Management.Automation.PSCredential ($o365_admin_username, $secpasswd)

try{
  Import-Module MSOnline
}catch{
  $module.FailJson("Cannot Import Powershell Module MSOnline, is it Missing on execution system?")
}

try{
  Connect-MsolService -Credential $o365_admin_creds
}catch{
  $module.FailJson("Error on Connection to MSOnline Service")
}

$msol_group_id = Get-MsolGroup | Where-Object DisplayName -eq $module.Params.MsolGroup | Select-Object -ExpandProperty ObjectId
if(-not $msol_group_id){
  $module.FailJson( ("Could not find targeted MSOnline Group {0}" -f $module.params.MsolGroup))
}

# Get Current Group Members from MSOnline
$current_group_members = Get-MsolGroupMember -GroupObjectId $msol_group_id | Select-Object -ExpandProperty ObjectId | Foreach-Object { $_.ToString() }
#$module.Result.currentGroupMembers = $current_group_members


# Build Desired Group Member List
$desiredMembers=@()

# Get Member objects from Ansible Param Member
foreach($m in $module.Params.MsolMembers){
  $desiredMembers += (Get-MsolUser -UserPrincipalName $m | Select-Object -ExpandProperty ObjectId).ToString()
}

# Now Work with MSOL Membership Groups
foreach($msol_group in $module.Params.MsolMembershipGroups){

  $member_msol_group_id = Get-MsolGroup | Where-Object DisplayName -eq $msol_group | Select-Object -ExpandProperty ObjectId
  $member_msol_group_members = Get-MsolGroupMember -GroupObjectId $member_msol_group_id | Select-Object -ExpandProperty ObjectId
  foreach( $msol_user In $member_msol_group_members){
      $desiredMembers += $msol_user
  }
}

# Now we Calculate what we need to change
$add_members = @()
$remove_members =@()
$diff_members = @()

# Remove Duplicates if happen
$desiredMembers = $desiredMembers | Sort-Object -Unique

# Check what we need to Change if so we can go to desired state
if(-not $desiredMembers){
  $remove_members = $current_group_members
}elseif (-not $current_group_members) {
  $add_members = $desiredMembers
}else{
  # We need to Compare
  $diff_members = Compare-Object -ReferenceObject $desiredMembers -DifferenceObject $current_group_members
  foreach($item in $diff_members){
    if($item.SideIndicator -eq "<="){ $add_members += $item.InputObject }
    if($item.SideIndicator -eq "=>"){ $remove_members += $item.InputObject }
  }
}

# Remove Duplicates if happen
$add_members = $add_members | Sort-Object -Unique
$remove_members = $remove_members | Sort-Object -Unique

# Add Members
if($add_members){
  $module.Result.changed = $true
  if(-not $module.CheckMode){
    foreach($m in $add_members){
      try{
        Add-MsolGroupMember -GroupObjectId $msol_group_id -GroupMemberObjectId $m
      }catch{
        $module.FailJson("Error on Adding MsolGroupmember $($_.Exception.Message)")
      }
    }
  }
}

# Remove Members
if($remove_members){
  $module.Result.changed = $true
  if( (-not $module.CheckMode) -and (-not $module.Params.AddOnly) ){
    foreach($m in $remove_members){
      try{
        Remove-MsolGroupMember -GroupObjectId $msol_group_id -GroupMemberObjectId $m
      }catch{
        $module.FailJson("Error on Removing MsolGroupMember $($_.Exception.Message)")
      }
    }
  }
}

$module.ExitJson()