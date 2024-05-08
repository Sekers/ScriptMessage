# INTERNAL TESTING SCRIPT

#################################################
# Import General Testing Configuration Settings #
#################################################

# $Config = Get-Content -Path "$PSScriptRoot\Config\config_general.json" | ConvertFrom-Json

#############
# DEBUGGING #
#############

$ErrorActionPreference = 'Stop'
# [string]$VerbosePreference = $Config.Debugging.VerbosePreference # Use 'Continue' to Enable Verbose Messages and Use 'SilentlyContinue' to reset back to default.
# [bool]$LogDebugInfo = $Config.Debugging.LogDebugInfo # Writes Extra Information to the log if $true.

##############################
# Configure Script Messaging #
##############################

# Import ScriptMessage Module
Import-Module Script-Message

# Set ScriptMessage MessagingConfiguration Path
Set-ScriptMessageConfigFilePath -Path "$PSScriptRoot\Config\config_scriptmessage.json"

###############################
# PUT TESTING CODE AFTER HERE #
###############################
