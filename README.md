# Autotask Contacts Sync Group Setup
This PowerShell script automates the creation of a dynamic security group in Microsoft 365 for Autotask Contacts synchronization.

## Features
- Installs and updates required Microsoft Graph PowerShell modules
- Creates or manages dynamic security group with specific membership rules
- Handles application consent flow
- Generates email notification with tenant and group details
- Includes error handling and user confirmations

## Prerequisites
- PowerShell 5.1 or higher
- Microsoft Graph PowerShell SDK
- Microsoft 365 admin credentials

## Required Permissions
- Group.ReadWrite.All
- Organization.Read.All

## Usage
1. Run the script
2. Follow prompts to authenticate
3. Confirm group creation/deletion if needed
4. Complete app consent process
5. Email will be generated automatically

## Notes
- Group members are filtered based on specific service plan IDs
- Script checks for existing groups to prevent duplicates
- Automatically formats email with tenant details
