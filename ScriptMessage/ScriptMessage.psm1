# Module Variables
# Module scope rather than global, so callers can neither see nor change them, and each import starts them fresh.

# The configuration file path set by Set-ScriptMessageConfigFilePath, read by Get-ScriptMessageConfig.
$script:ScriptMessageConfigFilePath = $null

# Each service's context from Get-ScriptMessageContext. Disconnect-ScriptMessage removes a service's entry.
$script:ScriptMessageCachedServiceContext = [PSCustomObject]@{}

# Each file in Services/ registers itself here as it is dot-sourced below.
$script:ScriptMessageServiceTable = [ordered]@{}

# Aliases

# Type Definitions

# Public Enum
# Name: MessagingService
enum MessagingService {
    MicrosoftGraph
}

# Public Enum
# Name: MessageType
enum MessageType {
    Mail
    Chat
}

# Public Enum
# Name: MailType
enum MailType {
    OneOnOne
    Group
}

# Public Enum
# Name: ChatType
enum ChatType {
    OneOnOne
    Group
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
