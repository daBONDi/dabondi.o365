#!powershell

# Copyright: (c) 2020 (David Baumann(@daBONDi) <me@davidbaumann.at>
# GNU General Public License v3.0+ (see COPYING or https://www.gnu.org/licenses/gpl-3.0.txt)

#AnsibleRequires -CSharpUtil Ansible.Basic

# Template List
# https://docs.microsoft.com/en-us/microsoftteams/get-started-with-teams-templates
# Education Template educationProfessionalLearningCommunity = EDU_CLASS

$spec = @{
  options = @{
      DisplayName = @{ type = "str"; required = $true }

      # Edu Template 
      Template = @{ type = "str"; default=""; required=$false}
      
      o365_admin_username = @{ type="str"; required = $true; no_log = $true}
      o365_admin_password = @{ type="str"; required = $true; no_log = $true}

      #Params we send to Set Teams
      MailNickName = @{ type= "str"; required =$false; default="" }
      Description = @{ type = "str"; required = $false; default=" " }
      AllowAddRemoveApps = @{ type = "bool"; required = $false; default=$false}
      AllowChannelMentions = @{ type = "bool"; required = $false; default=$true}
      AllowCreateUpdateChannels = @{ type = "bool"; required = $false; default=$false}
      AllowCreateUpdateRemoveConnectors = @{ type = "bool"; required = $false; default=$false}
      AllowCreateUpdateRemoveTabs = @{ type = "bool"; required = $false; default=$false}
      AllowCustomMemes = @{ type = "bool"; required = $false; default=$false}
      AllowDeleteChannels = @{ type = "bool"; required = $false; default=$false}
      AllowGiphy = @{ type = "bool"; required = $false; default=$true}
      AllowGuestCreateUpdateChannels = @{ type = "bool"; required = $false; default=$false}
      AllowGuestDeleteChannels = @{ type = "bool"; required = $false; default=$false}
      AllowOwnerDeleteMessages = @{ type = "bool"; required = $false; default=$true}
      AllowStickersAndMemes = @{ type = "bool"; required = $false; default=$true}
      AllowTeamMentions = @{ type = "bool"; required = $false; default=$true}
      AllowUserDeleteMessages = @{ type = "bool"; required = $false; default=$true}
      AllowUserEditMessages = @{ type = "bool"; required = $false; default=$true}
      Visibility = @{ type = "str"; required = $false; default="Private"; choices ="Private","Public"}
  }
  supports_check_mode = $true
  no_log = $true
}
$module = [Ansible.Basic.AnsibleModule]::Create($args, $spec)


# We Iterate over all Params we not Itterate Over to check if changed
$filtered_params = @('DisplayName','o365_admin_username','o365_admin_password','Template')

# Replace Mail Nick With Generated one
if($module.Params.MailNickName -eq ""){
  $generated_mail_nick_name = $module.Params.DisplayName -Replace '[^a-zA-Z0-9\-_]',''
  $generated_mail_nick_name = $generated_mail_nick_name.Trim(' ')
  $module.Params.MailNickName =  $generated_mail_nick_name.ToLower()
}

$team_name =$module.Params.DisplayName
$team_template = $module.Params.Template
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

$module.Result.requested_team = $team_name

if($team_object -is [array]){
  $module.FailJson("We found Multiple Teams!")
}

# We need to Create the Team
if(-not $team_object){
  $module.Result.changed=$true
  $module.Result.team_created = $true
  if(-not $module.CheckMode){

    if($team_template -eq ""){
      try{
        $team_object = New-Team -DisplayName $team_name -MailNick $module.Params.MailNickName
      }catch{
        $module.FailJson("Error on Creating Team $team_name without Template : $_.Exception.Message",$_)
      }
      
    }else{
      try{
        $team_object = New-Team -DisplayName $team_name -Template $team_template -MailNick $module.Params.MailNickName
      }catch{
        $module.FailJson("Error on Creating Team $team_name with Template $team_template  : $_.Exception.Message",$_)
      }
    }

    
  }else{
    # We need to Break out because we would add team, but it is not existsing on check mode
    $module.ExitJson()
  }
}


if(-not $team_object){
  $module.FailJson("We fail on fetching Team Object, before we can set any propertys")
}

# Now we Build the Update Parameters
$update_command_params = @{}

Foreach($param in $module.Params.GetEnumerator()){
  if(-not ($filtered_params -contains $param.key)){
    # If Param is not null
    if($param.value -ne $null){
      $add=$false
      # String Property
      if($param.value -is [string]){
        if(-not ($param.value.equals( $($team_object | Select-Object -ExpandProperty $param.key) ) ) ){

          # Edge Case - Visibility = HiddenMembership , we cannot change from Hidden to Private
          if($team_object.Visibility.equals("HiddenMembership") -and $param.key.equals("Visibility") ){
            $add = $false
            $module.warn("Group Visibility is HiddenMembership we cannot change Visbility When Membership is hidden!")
          }else{
            $add = $true
          }
          
        }
      }elseif ($param.value -is [bool]){
        if($param.value -ne $($team_object | Select-Object -ExpandProperty $param.key) ){
          $add = $true
        }
      }else{
        $module.warn("We could not detect Type for $($param.key):$($param.value)")
      }

      # We add it to Global Params
      if($add){
        $update_command_params.add($param.key,$param.value)
      }
    }
  }
}


$module.Result.updated_propertys = $update_command_params
# Now we Apply it
if($update_command_params.Count){
  if(-not $module.CheckMode){
    Set-Team -GroupId $($team_object.GroupId) @update_command_params | Out-Null
  }
  $module.Result.changed=$true
}


try{
  Disconnect-MicrosoftTeams | Out-Null
}catch{
  $module.FailJson("Error on disconnecting from Teams $_.Exception.Message",$_)
}

$module.ExitJson()