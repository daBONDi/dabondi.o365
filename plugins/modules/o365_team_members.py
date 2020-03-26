#!/usr/bin/python
# -*- coding: utf-8 -*-

# Copyright: (c) 2020 (David Baumann(@daBONDi) <me@davidbaumann.at>
# GNU General Public License v3.0+ (see COPYING or https://www.gnu.org/licenses/gpl-3.0.txt)

ANSIBLE_METADATA = {'status': ['preview'],
                    'supported_by': 'Community',
                    'version': '1.0'}

DOCUMENTATION = r'''
---
module: o365_team_members
short_description: Ensure Teams Team Membership
description:
  - Ensure Members an Owners for the Team
  - You can use direct Membership, or use a MSOnline Group and the Members of the MSOnline Group will be added
requirements:
  - Powershell Module MicrosoftTeams
  - Powershell Module MSOnline
options:
  DisplayName:
    description:
      - DisplayName of the Teams Team
    required: yes
    type: str
  
  Members:
    description:
      - List of direct Members (E-Mail/UPN of O365 User)
    required: no
    type: list

  Owners:
    description:
      - List of direct Owners (E-Mail/UPN of O365 User)
    required: no
    type: list
  
  MSOGroupMembers:
    description:
      - List of MSOnline Groups(Security) where there Members will be added as Members to the Team
    required: no
    type: list

  MSOGroupOwners:
    description:
      - List of MSOnline Groups(Security) where there Members will be added as Owners to the Team
    required: no
    type: list

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
- name: Ensure Team Members are in Team
  o365_team_members:
    DisplayName: "{{ item.DisplayName }}"
    Members: "{{ item.Members | default(omit) }}"
    Owners: "{{ item.Owners + o365.teams.default_owners }}"
    MSOGroupMembers: "{{ item.MSOGroupMembers }}"
    MSOGroupOwners: "{{ item.MSOGroupOwners }}"
    o365_admin_username: "{{ o365user }}"
    o365_admin_password: "{{ o365password }}"

  with_items: "{{ teams }}"
  loop_control:
    label: "{{ item.DisplayName }}"
'''

RETURN = r'''

'''

