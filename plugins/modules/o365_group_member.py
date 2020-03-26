#!/usr/bin/python
# -*- coding: utf-8 -*-

# Copyright: (c) 2020 (David Baumann(@daBONDi) <me@davidbaumann.at>
# GNU General Public License v3.0+ (see COPYING or https://www.gnu.org/licenses/gpl-3.0.txt)

ANSIBLE_METADATA = {'status': ['preview'],
                    'supported_by': 'Community',
                    'version': '1.0'}

DOCUMENTATION = r'''
---
module: o365_group_member
short_description: Manage Office 365 Group Membership with MSOnline Powershell Module
description:
  - Manage Office 365 Group Membership with MSOnline Powershell Module
  - Can add Azure AD Securitygroup Members into an Office 365 Group
requirements:
  - Powershell Module MSOnline
options:
  MsolGroup:
    description:
      - Name of the MsolGroup(Azure AD Group/Office 365 Group) to change
    required: yes
    type: str

  MsolMembershipGroups:
    description:
      - All Members in this MsolGroups will be Members of the Group defined with MsolGroup
    required: no
    type: list
    default: []

  MsolMembers:
    description:
      - List of Members(E-Mail Addresses) for the Group MsolGroup
    required: no
    default: []

  AddOnly:
    description:
      - Add only Members don't remove Members if they are not defined
    type: bool
    default: no

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
'''

EXAMPLES = r'''

- o365_group_member:
    o365_admin_username: o365user
    o365_admin_password: o365password
    MsolGroup: 'o365-allow-group-creation'
    MsolMembers:
      - 'Sync_FJ-V-USYNC1_52ed24809b42@josephinumat.onmicrosoft.com'
    MsolMembershipGroups:
      - "usg-service-o365-teacher"
      - "usg-service-o365-employe"
'''

RETURN = r'''

'''

