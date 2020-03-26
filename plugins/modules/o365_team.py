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
short_description: Ensure a specific Team exists in Microsoft Teams
description:
  - Ensure a specific Team exists in Microsoft Teams
requirements:
  - Powershell Module MicrosoftTeams
options:
  DisplayName:
    description:
      - DisplayName of the Teams Team
    required: yes
    type: str
  
  Template:
    description:
      - Template to use, Currently there is only one Template 'EDU_Class'
    type: str
    default: ''

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

  MailNickName:
    description:
      - The MailNickName parameter specifies the alias for the associated Office 365 Group. This value will be used for the mail enabled object and will be used as PrimarySmtpAddress for this Office 365 Group. The value of the MailNickName parameter has to be unique across your tenant.
    type: str
    default: ''
  
  Description:
    description:
      - Team description. Team Description Characters Limit - 1024.
    type: str
    default: ''

  AllowAddRemoveApps:
    description:
      - Boolean value that determines whether or not members (not only owners) are allowed to add apps to the team.
    type: bool
    default: false

  AllowChannelMentions:
    description:
      - Boolean value that determines whether or not channels in the team can be @ mentioned so that all users who follow the channel are notified.
    type: bool
    default: true

  AllowCreateUpdateChannels:
    description:
      - Setting that determines whether or not members (and not just owners) are allowed to create channels.
    type: bool
    default: false

  AllowCreateUpdateRemoveConnectors:
    description:
      - Setting that determines whether or not members (and not only owners) can manage connectors in the team.
    type: bool
    default: false

  AllowCreateUpdateRemoveTabs:
    description:
      - Setting that determines whether or not members (and not only owners) can manage tabs in channels.
    type: bool
    default: false

  AllowCustomMemes:
    description:
      - Setting that determines whether or not members can use the custom memes functionality in teams.
    type: bool
    default: false

  AllowDeleteChannels:
    description:
      - Setting that determines whether or not members (and not only owners) can delete channels in the team.
    type: bool
    default: false

  AllowGiphy:
    description:
      - Setting that determines whether or not members (and not only owners) can delete channels in the team.
    type: bool
    default: true

  AllowGuestCreateUpdateChannels:
    description:
      - Setting that determines whether or not guests can create channels in the team.
    type: bool
    default: false

  AllowGuestDeleteChannels:
    description:
      - Setting that determines whether or not guests can delete in the team.
    type: bool
    default: false
  
  AllowOwnerDeleteMessages:
    description:
      - Setting that determines whether or not owners can delete messages that they or other members of the team have posted.
    type: bool
    default: true

  AllowStickersAndMemes:
    description:
      - Setting that determines whether stickers and memes usage is allowed in the team.
    type: bool
    default: true

  AllowTeamMentions:
    description:
      - Setting that determines whether the entire team can be mentioned (which means that all users will be notified)
    type: bool
    default: true

  AllowUserDeleteMessages:
    description:
      - Setting that determines whether or not members can delete messages that they have posted.
    type: bool
    default: true

  AllowUserEditMessages:
    description:
      - Setting that determines whether or not users can edit messages that they have posted.
    type: bool
    default: true


  Visibility:
    description:
      - Team visibility valid values are Private and Public
    type: str
    choices: [Private,Public]
    default: 'Private'

Author: David Baumann(@daBONDi)

notes:
- There is currently only 1 Template EDU_CLASS defined in the Powershell Module
- https://docs.microsoft.com/en-us/microsoftteams/get-started-with-teams-templates
'''

EXAMPLES = r'''
- name: Ensure Team is created for Class
  o365_team:
    DisplayName: "{{ item.DisplayName }}"
    Template: "{{ team_template }}"
    Description: "{{ item.Description }}"
    o365_admin_username: "{{ o365user }}"
    o365_admin_password: "{{ o365password }}"
  with_items: "{{ teams }}"
  loop_control:
    label: "{{ item.DisplayName }}"
'''

RETURN = r'''
requested_team: 
  description: The name of the Team we requested
  type: str

team_created: 
  description: If the Team is created or modifed
  type: bool

update_propertys: 
  description: A list of Values which Properties changed
  type: list
'''

