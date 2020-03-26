#!/usr/bin/python
# -*- coding: utf-8 -*-

# Copyright: (c) 2020 (David Baumann(@daBONDi) <me@davidbaumann.at>
# GNU General Public License v3.0+ (see COPYING or https://www.gnu.org/licenses/gpl-3.0.txt)

ANSIBLE_METADATA = {'status': ['preview'],
                    'supported_by': 'Community',
                    'version': '1.0'}

DOCUMENTATION = r'''
---
module: o365_team_generate_picture
short_description: Generate a Picture for the Teams Team
description:
  - Generate a Picute for the Teams Team and assign it
  - It takes a Time until is visible(< 5 Hours)
requirements:
  - Powershell Module MicrosoftTeams
options:
  DisplayName:
    description:
      - DisplayName of the Teams Team
    required: yes
    type: str
  
  PictureText:
    description:
      - Text in the Picture
    required: yes
    type: str

  BackgroundColorName:
    description:
      - Name of the .Net Framework Color to use in Background
      - https://docs.microsoft.com/en-us/dotnet/api/system.windows.media.brushes?redirectedfrom=MSDN&view=netframework-4.8
    required: yes
    type: str
  
  ForegroundColorName:
    description:
      - Name of the .Net Framework Color to use for the Text
      - https://docs.microsoft.com/en-us/dotnet/api/system.windows.media.brushes?redirectedfrom=MSDN&view=netframework-4.8
    required: yes
    type: str

  FontSize:
    description:
      - Size of the Font to use for the Picute
    required: no
    type: int
    default: 22

  o365_admin_username:
    description:
    - Username for connecting with the Powershell Module MSOnline to O365
    type: str
    required: yes

  o365_admin_password:
    description:
    - Password for connecting with the Powershell Module MSOnline to O365
    type: str
    required: yes

author: David Baumann(@daBONDi)

notes:
- https://www.undocumented-features.com/2019/12/02/add-posh-test-gallery-to-add-microsoft-teams-test-powershell-module/
- Register-PackageSource -Name PoshTestGallery -Location https://www.poshtestgallery.com/api/v2/ -ProviderName PowerShellGet
- Install-Module -Name MicrosoftTeams -RequiredVersion 1.0.21 -Repository PoshTestGallery

'''

EXAMPLES = r'''
- name: Ensure Team has a nice Class Picture
  o365_team_generate_picture:
    DisplayName: "{{ item.DisplayName }}"
    PictureText: "{{ item.ShortCutName }}"
    o365_admin_username: "{{ o365user }}"
    o365_admin_password: "{{ o365password }}"
    BackgroundColorName: "{{ item.BackgroundColorName }}"
  with_items: "{{ teams }}"
  loop_control:
    label: "{{ item.DisplayName }}"
'''

RETURN = r'''

'''

