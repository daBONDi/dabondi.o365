#!powershell

# Copyright: (c) 2020 (David Baumann(@daBONDi) <me@davidbaumann.at>
# GNU General Public License v3.0+ (see COPYING or https://www.gnu.org/licenses/gpl-3.0.txt)

#AnsibleRequires -CSharpUtil Ansible.Basic

# Note: the command will return immediately, but the Teams application will not reflect the update immediately. The Teams application may need to be open for up to an hour before changes are reflected.

# We need Teams Powershell Test Environment module_defaults
# https://www.undocumented-features.com/2019/12/02/add-posh-test-gallery-to-add-microsoft-teams-test-powershell-module/
# Register-PackageSource -Name PoshTestGallery -Location https://www.poshtestgallery.com/api/v2/ -ProviderName PowerShellGet
# Install-Module -Name MicrosoftTeams -RequiredVersion 1.0.21 -Repository PoshTestGallery

$spec = @{
  options = @{
      DisplayName = @{ type="str"; required = $true }
      PictureText = @{ type="str"; required = $true}

      # For Color Names check https://docs.microsoft.com/en-us/dotnet/api/system.windows.media.brushes?redirectedfrom=MSDN&view=netframework-4.8
      BackgroundColorName = @{ type="str"; required = $false; default="ForestGreen"}
      ForegroundColorName = @{ type="str"; required = $false; default="White"}

      FontSize = @{ type="int"; required = $false; default=22}
      o365_admin_username = @{ type="str"; required = $true; no_log = $true}
      o365_admin_password = @{ type="str"; required = $true; no_log = $true}
  }
  supports_check_mode = $true
  no_log = $true
}

$default_image_height = 100
$default_image_width = 100

$image_width = $default_image_height
$image_height = $default_image_width

$module = [Ansible.Basic.AnsibleModule]::Create($args, $spec)

$backgroundColorName = $module.Params.BackgroundColorName
$foreGroundColorName  = $module.Params.ForegroundColorName
$fontSize  = $module.Params.FontSize
$text  = $module.Params.PictureText

# Generate Image
$image_path = "$($env:Temp)\$([guid]::NewGuid()).png"

$module.Result.ImagePath = $image_path


Add-Type -AssemblyName System.Drawing
$bmp = new-object System.Drawing.Bitmap $image_width,$image_height 
$font = new-object System.Drawing.Font Consolas,$fontSize
$brushBg = [System.Drawing.SolidBrush]::new([System.Drawing.Color]::FromName($backgroundColorName))
$brushFg = [System.Drawing.SolidBrush]::new([System.Drawing.Color]::FromName($foreGroundColorName))
$graphics = [System.Drawing.Graphics]::FromImage($bmp) 
$graphics.FillRectangle($brushBg,0,0,$bmp.Width,$bmp.Height)

$rect = [System.Drawing.RectangleF]::FromLTRB(0, 0, $image_width, $image_height)
$format = [System.Drawing.StringFormat]::GenericDefault
$format.Alignment = [System.Drawing.StringAlignment]::Center
$format.LineAlignment = [System.Drawing.StringAlignment]::Center

$graphics.DrawString($text,$font,$brushFg,$rect,$format) 
$graphics.Dispose() 
$bmp.Save($image_path) 

$team_name =$module.Params.DisplayName
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

if(-not $team_object){
  $module.FailJson("We could not Find Team with Displayname $team_name")
}

# Set Image
try{
  Set-TeamPicture -GroupId $($team_object.GroupId) -ImagePath $image_path | Out-Null
}catch{
  $module.FailJson("Error on Setting Team Picture", $_)
}finally{
  If(Test-Path $image_path){
    Remove-Item -Path $image_path -Force
  }
  
}

# We always need to Set an Image
$module.Result.changed = $true

try{
  Disconnect-MicrosoftTeams | Out-Null
}catch{
  $module.FailJson("Error on disconnecting from Teams $_.Exception.Message",$_)
}

$module.ExitJson()