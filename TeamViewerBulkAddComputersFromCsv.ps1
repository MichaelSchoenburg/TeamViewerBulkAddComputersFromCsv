<#
.SYNOPSIS
    MscTeamViewerInterface

.DESCRIPTION
    Long description

.SYNTAX


.PARAMETERS


.EXAMPLE
    PS C:\> <example usage>
    Explanation of what the example does

.INPUTS
    Inputs (if any)

.OUTPUTS
    Output (if any)

.RELATED LINKS
    GitHub: https://github.com/MichaelSchoenburg/MscTeamViewerInterface

.NOTES
    Author: Michael Schönburg
    Version: v1.0
    Last Edit: 07.04.2022
    
    This projects code loosely follows the PowerShell Practice and Style guide, as well as Microsofts PowerShell scripting performance considerations.
    Style guide: https://poshcode.gitbook.io/powershell-practice-and-style/
    Performance Considerations: https://docs.microsoft.com/en-us/powershell/scripting/dev-cross-plat/performance/script-authoring-considerations?view=powershell-7.1

.REMARKS
    To see the examples, type: "get-help Get-HotFix -examples".
    For more information, type: "get-help Get-HotFix -detailed".
    For technical information, type: "get-help Get-HotFix -full".
    For online help, type: "get-help Get-HotFix -online"
#>

#region INITIALIZATION
<# 
    Libraries, Modules, ...
#>

#endregion INITIALIZATION
#region DECLARATIONS
<#
    Declare local variables and global variables
#>

# $ApiTokenUser = 'MichaelSchoenburg'
if ($PSScriptRoot) {
    Set-Location -Path $PSScriptRoot
} else {
    Set-Location -Path "C:\Users\mschoenburg\GIT\MscTeamViewerInterface"
}
$ApiTokenUser = 'Onboarding'
$ApiToken = (Get-Content .\TeamViewerData.ini | Select -Skip 1 | ConvertFrom-StringData).$ApiTokenUser 
$ApiTokenSec = ConvertTo-SecureString -String $ApiToken -AsPlainText -Force

$Csv = Import-Csv .\TeamViewer-Passwoerter.csv -Delimiter ';'

$prefixLengthG = 11
$prefixLengthIt = 12

#endregion DECLARATIONS
#region FUNCTIONS
<# 
    Declare Functions
#>

function Write-ConsoleLog {
    <#
    .SYNOPSIS
    Logs an event to the console.
    
    .DESCRIPTION
    Writes text to the console with the current date (US format) in front of it.
    
    .PARAMETER Text
    Event/text to be outputted to the console.
    
    .EXAMPLE
    Write-ConsoleLog -Text 'Subscript XYZ called.'
    
    Long form
    .EXAMPLE
    Log 'Subscript XYZ called.
    
    Short form
    #>
    [alias('Log')]
    [CmdletBinding()]
    param (
        [Parameter(Mandatory = $true,
        Position = 0)]
        [string]
        $Text
    )

    # Save current VerbosePreference
    $VerbosePreferenceBefore = $VerbosePreference

    # Enable verbose output
    $VerbosePreference = 'Continue'

    # Write verbose output
    Write-Verbose "$( Get-Date -Format 'MM/dd/yyyy HH:mm:ss' ) - $( $Text )"

    # Restore current VerbosePreference
    $VerbosePreference = $VerbosePreferenceBefore
}

#endregion FUNCTIONS
#region EXECUTION
<# 
    Script entry point
#>

Connect-TeamViewerApi -ApiToken $ApiTokenSec # Install-Module TeamViewerPS

$Groups = Get-TeamViewerGroup

for ($i = 0; $i -lt $Csv.Count; $i++) {
    # Write-Progress -Activity 'TeamViewer Stuff' -CurrentOperation 'Checking if device already exists' -PercentComplete ( $i/$CSV.Count*100 ) # Doesn't work in VS Code
    
    # TeamViewer ID
    $TvId = $null
    # Cut out all the empty spaces
    $TvId = $Csv[$i].'Geraete-ID' -replace '\s',''

    # Group
    $Group = $null
    $g = $Csv[$i].Adresse
    $Group = $g.Substring($prefixLengthG, ($g.Length - $prefixLengthG))
    $InternalCustId = $g.Substring(0,8) # Internal Customer ID
    $TvGroup = $Groups.where({$_.Name -like "*$( $InternalCustId )*"})

    # Name
    $Name = $null
    $ItInv = $Csv[$i].'IT-Inventar'
    $ItInvCut = $ItInv.Substring($prefixLengthIt, ($ItInv.Length - $prefixLengthIt))
    if (($ItInvCut -contains '|') -or ($ItInvCut -contains '(') -or ($ItInvCut -contains '#')) {
        $Name = $ItInvCut
    } else {
        # Adresskontakt doesn't have no ID in front of it, 
        # so it must be compared to the variable where the ID has been cut off already

        # Check if Adresskontakt is an actual person or just the Address/Firm
        $Contact = $Csv[$i].Adresskontakt
        if ($Group -ne $Contact) { 
            $Name = "$( $ItInvCut ) | $( $Contact )"
        } else {
            $Name = $ItInvCut
        }
    }
    # Maximal length is 50
    if ($Name.Length -gt 50) {
        $Name = $Name.Substring(0,50)
    }

    # Password
    $PW = $null
    $password = $Csv[$i].Passwort
    if ($password) {
        $PW = ConvertTo-SecureString -String $password -AsPlainText -Force   
    }

    if (Get-TeamViewerDevice -TeamViewerId $TvId) {
        Write-Host "$( Get-Date -Format 'MM/dd/yyyy HH:mm:ss' ) - [$( $i )/$($Csv.Count)] ALREADY EXISTS" -ForegroundColor Magenta
        Write-Host "TeamViewer ID =         $( $TvId )" -ForegroundColor Magenta
        Write-Host "Group =                 $( $Group )" -ForegroundColor Magenta
        Write-Host "Name =                  $( $Name )" -ForegroundColor Magenta
        Write-Host "Password =              $( $PW )" -ForegroundColor Magenta
    } else {
        Write-Host "$( Get-Date -Format 'MM/dd/yyyy HH:mm:ss' ) - [$( $i )/$($Csv.Count)] PROCESSING" -ForegroundColor Cyan
        Write-Host "TeamViewer ID =         $( $TvId )" -ForegroundColor Cyan
        Write-Host "Group =                 $( $Group )" -ForegroundColor Cyan
        Write-Host "InternalCustId =        $( $InternalCustId )" -ForegroundColor Cyan
        Write-Host "TeamViewer Group ID =   $( $TvGroup.Id )" -ForegroundColor Cyan
        Write-Host "TeamViewer Group Name = $( $TvGroup.Name )" -ForegroundColor Cyan
        Write-Host "Name =                  $( $Name )" -ForegroundColor Cyan
        Write-Host "Contact =               $( $Contact )" -ForegroundColor Cyan
        Write-Host "Password =              $( $PW )" -ForegroundColor Cyan
        
        if ($password) {
            New-TeamViewerDevice -TeamViewerId $TvId -Group $TvGroup.Id -Name $Name -Password $PW
        } else {
            New-TeamViewerDevice -TeamViewerId $TvId -Group $TvGroup.Id -Name $Name
        }
    }
}

Disconnect-TeamViewerApi

#endregion EXECUTION
