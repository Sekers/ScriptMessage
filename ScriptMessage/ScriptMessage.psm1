﻿# Type Definitions

# Public Enum
# Name: MessagingService
enum MessagingService {
    MgGraph
}

# Public Enum
# Name: MessageType
enum MessageType {
    Mail
    Chat
}

# Public Class
class MessageServiceType {
    [MessagingService[]]$Service
    [MessageType[]]$Type
}

# Import Private Functions
$ScriptMessageFunctions = @(Get-ChildItem -Path $PSScriptRoot\Private\*.ps1)
Foreach($ScriptMessageFunction in $ScriptMessageFunctions)
{
    Write-Verbose "Importing $ScriptMessageFunction"
    Try
    {
        . $ScriptMessageFunction.fullname
    }
    Catch
    {
        Write-Error -Message "Failed to import function $($ScriptMessageFunction.fullname): $_"
    }
}

# Import Public Functions
$ScriptMessageFunctions = @(Get-ChildItem -Path $PSScriptRoot\Public\*.ps1)
Foreach($ScriptMessageFunction in $ScriptMessageFunctions)
{
    Write-Verbose "Importing $ScriptMessageFunction"
    Try
    {
        . $ScriptMessageFunction.fullname
    }
    Catch
    {
        Write-Error -Message "Failed to import function $($ScriptMessageFunction.fullname): $_"
    }
}

# Import Services
$ScriptServices = @(Get-ChildItem -Path $PSScriptRoot\Services\*.ps1)
Foreach($ScriptService in $ScriptServices)
{
    Write-Verbose "Importing $ScriptService"
    Try
    {
        . $ScriptService.fullname
    }
    Catch
    {
        Write-Error -Message "Failed to import function $($ScriptService.fullname): $_"
    }
}
