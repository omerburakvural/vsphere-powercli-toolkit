# vSphere PowerCLI Toolkit

A curated collection of PowerShell and VMware PowerCLI scripts for common vSphere administration and operational automation tasks.

The scripts in this repository are sanitized and generalized versions of practical automation patterns. They do not contain production credentials, internal hostnames, IP addresses, customer information, or employer-specific configuration.

## Current Scripts

### VMRemove

Removes a virtual machine from a vSphere environment using VMware PowerCLI.

Before using the script:

- Review all parameters carefully.
- Test in a non-production environment.
- Confirm the target virtual machine.
- Use an account with the minimum required permissions.
- Ensure that backups or recovery options are available.

## Requirements

- PowerShell 7 or Windows PowerShell 5.1
- VMware PowerCLI
- Network access to a VMware vCenter Server
- Appropriate vSphere permissions

Install VMware PowerCLI:

```powershell
Install-Module VMware.PowerCLI -Scope CurrentUser
