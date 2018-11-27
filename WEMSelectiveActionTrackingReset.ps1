<#
.SYNOPSIS
This is a simple Powershell script to manage the registry caching of tracked WEM actions. 
Idea being that it can be published a simple WEM application and utilised in support scenarios saving users from trauling through the registry

.DESCRIPTION
Script parses the relevent Norskale\VirtualAll registry locations which hold individual action tracking
Aimed to address scenarios where you don't want to kill the entire action tracking
Address challenges such as run once requirements or changes to requirements, USV changes which require a reset of the USV action tracking etc

V1.0 - James Kindon - Initial release
#>

# Define Reset Cache Function
function ResetActionTracking {
    [cmdletbinding()]
    param (
        [Parameter(Mandatory)]
        [ValidateSet('Applications',
        'EnvironmentalSettings',
        'EnvironmentVariables',
        'ExternalTasks',
        'FileAssociations',
        'FileSystemOperations',
        'Groups',
        'IniFilesOperations',
        'InitTasks',
        'NetworkDrives',
        'Ports',
        'Printers',
        'RegistryValues',
        'SystemUtilitiesSettings',
        'SystemUtilitiesSettingsGroups',
        'UserDsns',
        'UserSelfAssignedApplications',
        'UsvUserConfigurationSettings',
        'VirtualDrives',
        'UserPrinters',
        'All')]
        [string]
        $ActionSet
    )
    #Get Path details
    $RootExecTrackingPath = 'HKCU:\SOFTWARE\VirtuAll Solutions\VirtuAll User Environment Manager\Agent\Tasks Exec Cache'
    
    if ($ActionSet -ne "All" -and $ActionSet -ne "Groups" -and $ActionSet -ne "UserSelfAssignedApplications" -and $ActionSet -ne "SystemUtilitiesSettingsGroups" -and $ActionSet -ne "UserPrinters") {
        
        ConfirmAndKill
        
    } elseif ($ActionSet -eq "All") {

        ConfirmAndKillAll

    } elseif ($ActionSet -eq "Groups") {
        
        ConfirmAndKillReg -UniqueRegKey "LastGroupsList"

    } elseif ($ActionSet -eq "UserSelfAssignedApplications") {
        
        ConfirmAndKillReg -UniqueRegKey "AssignedApplicationsList"

    } elseif ($ActionSet -eq "SystemUtilitiesSettingsGroups") {
        
        ConfirmAndKillReg -UniqueRegKey "SystemUtilitiesKnownGroupList"

    } elseif ($ActionSet -eq "UserPrinters") {

        ConfirmAndKillPrinters

    }
}

# Define Show-Menu Function
# https://www.business.com/articles/powershell-interactive-menu/
function Show-Menu {
    param (
        [string]$Title = 'Workspace Environment Management Selective Action Tracking Reset Utility'
    )
    Clear-Host
    Write-Host "====== $Title ======"  -ForegroundColor Cyan
    Write-Host "======            Selectively remove problematic WEM Action Tracking            ======"  -ForegroundColor Cyan
     
    Write-Host "Enter '1' to reset Applications Tracking"
    Write-Host "Enter '2' to reset Environmental Settings Tracking"
    Write-Host "Enter '3' to reset Environment Variables Tracking"
    Write-Host "Enter '4' to reset External Tasks Tracking"
    Write-Host "Enter '5' to reset File Associations Tracking"
    Write-Host "Enter '6' to reset File System Operations Tracking"
    Write-Host "Enter '7' to reset Groups Tracking"
    Write-Host "Enter '8' to reset Ini Files Operations Tracking"
    Write-Host "Enter '9' to reset Init Tasks Tracking"
    Write-Host "Enter '10' to reset Network Drives Tracking"
    Write-Host "Enter '11' to reset Ports Tracking"
    Write-Host "Enter '12' to reset Printers Tracking"
    Write-Host "Enter '13' to reset Registry Values Tracking"
    Write-Host "Enter '14' to reset System Utilities Settings Tracking"
    Write-host "Enter '15' to reset System Utilities Settings Groups Tracking"
    Write-Host "Enter '16' to reset User DSN Tracking"
    Write-Host "Enter '17' to reset User Self Assigned Applications Tracking"
    Write-Host "Enter '18' to reset Usv User Configuration Settings (Folder Redirection) Tracking"
    Write-Host "Enter '19' to reset Virtual Drives Tracking"
    Write-Host "Enter '20' to reset User Initiated Default Printer Assignment Tracking"
    Write-Host "Enter '21' to reset All values (All WEM Processed items) Tracking" -ForegroundColor Yellow
    Write-Host "Enter '22' to refresh the WEM cache."
    Write-Host "Enter 'Q' to quit."
}

# Define Test-Registry Function
# https://stackoverflow.com/questions/5648931/test-if-registry-value-exists
Function Test-RegistryValue($regkey, $name) {
    $exists = Get-ItemProperty -Path "$regkey" -Name "$name" -ErrorAction SilentlyContinue
    If (($exists -ne $null) -and ($exists.Length -ne 0)) {
        Return $true
    }
    Return $false
}

function ConfirmAndKill {
    $TrackingSet = Get-ChildItem -Path $RootExecTrackingPath\$ActionSet -ErrorAction SilentlyContinue
    #Check Counts
    if ($TrackingSet.Count -gt "0") {
        Write-host "There are $($ActionSet) tracking records availble to remove" -ForegroundColor Cyan
        #Remove Items
        $ConfirmDelete = Read-Host "Would you like to remove WEM $($ActionSet) tracking records? Enter Y or N"
        while ("Y", "N" -notcontains $ConfirmDelete) {
            $ConfirmDelete = Read-Host "Enter Y or N"
        }
        if ($ConfirmDelete -eq "Y") {
            Write-Warning "Deleting WEM $($ActionSet) tracking records"
            $TrackingSet | Remove-Item -Force
        }
        elseif ($ConfirmDelete -eq "N") {
            Write-Host "No WEM $($ActionSet) Tracking Deleted"
        }
    }
    elseif ($TrackingSet.Count -eq "0") {
        Write-Warning "There are no $($ActionSet) records to remove"
    }
}

function ConfirmAndKillReg {
    [cmdletbinding()]
    param (
        [Parameter(Mandatory)]
        [string]
        $UniqueRegKey
    )
    if (Test-RegistryValue -regkey $RootExecTrackingPath\$ActionSet -name $UniqueRegKey) {
        Write-host "There are $($ActionSet) tracking records availble to remove" -ForegroundColor Cyan
        #Remove Items
        $ConfirmDelete = Read-Host "Would you like to remove WEM $($ActionSet) tracking records? Enter Y or N"
        while ("Y", "N" -notcontains $ConfirmDelete) {
            $ConfirmDelete = Read-Host "Enter Y or N"
        }
        if ($ConfirmDelete -eq "Y") {
            Write-Warning "Deleting WEM $($ActionSet) tracking records"
            Remove-ItemProperty -Path $RootExecTrackingPath\$ActionSet -Name $UniqueRegKey
        }
        elseif ($ConfirmDelete -eq "N") {
            Write-Host "No WEM $($ActionSet) Tracking Deleted"
        }
    }
    elseif (!(Test-RegistryValue -regkey $RootExecTrackingPath\$ActionSet -name $UniqueRegKey)) {
        Write-Warning "There are no $($ActionSet) records to remove"
    }
}

function ConfirmAndKillPrinters {
    $RootExecTrackingPath = 'HKCU:\SOFTWARE\VirtuAll Solutions\VirtuAll User Environment Manager\Agent\User Printers Management'
    if (Test-RegistryValue -regkey $RootExecTrackingPath -name "UserSelectedDefaultPrinter") {
        Write-host "There are $($ActionSet) tracking records availble to remove" -ForegroundColor Cyan
        #Remove Items
        $ConfirmDelete = Read-Host "Would you like to remove WEM $($ActionSet) tracking records? Enter Y or N"
        while ("Y", "N" -notcontains $ConfirmDelete) {
            $ConfirmDelete = Read-Host "Enter Y or N"
        }
        if ($ConfirmDelete -eq "Y") {
            Write-Warning "Deleting WEM $($ActionSet) tracking records"
            Remove-ItemProperty -Path $RootExecTrackingPath -Name "UserSelectedDefaultPrinter"
        }
        elseif ($ConfirmDelete -eq "N") {
            Write-Host "No WEM $($ActionSet) Tracking Deleted"
        }
    }
    elseif (!(Test-RegistryValue -regkey $RootExecTrackingPath -name "UserSelectedDefaultPrinter")) {
        Write-Warning "There are no $($ActionSet) records to remove"
    }
}

function ConfirmAndKillAll {
    $TrackingSet = Get-ChildItem -Path $RootExecTrackingPath -ErrorAction SilentlyContinue
    if ($TrackingSet.Count -gt "0") {
        $ConfirmDelete = Read-Host "Would you like to remove All WEM tracked actions? Enter Y or N"
        while ("Y", "N" -notcontains $ConfirmDelete) {
            $ConfirmDelete = Read-Host "Enter Y or N"
        }
        if ($ConfirmDelete -eq "Y") {
            Write-Warning "Deleting All WEM tracked actions"
            #Remove Entire Set
            $TrackingSet | Remove-Item -Recurse -Force
            Write-Host "Please refresh the WEM agent to re-apply actions"
            Read-Host "Press any Key to Exit"
            Exit
        }
        elseif ($ConfirmDelete -eq "N") {
            Write-Host "No WEM action tracking deleted"
        }
    }
    elseif ($TrackingSet.Count -eq "0") {
        Write-Warning "There are no WEM tracked actions to delete. Please refresh the WEM cache"
    }
}

# Get WEM Broker from Policy
$WEMBroker = (Get-ItemProperty -Path "HKLM:\SOFTWARE\Policies\Norskale\Agent Host\" -Name "BrokerSvcName").BrokerSvcName

# Run Tool
do {
    Show-Menu
    $input = Read-Host "Please make a selection"
    switch ($input) {
        '1' {
            Clear-Host
            'You selected to reset Application tracked actions'
            ResetActionTracking -ActionSet Applications
        } '2' {
            Clear-Host
            'You selected to reset Environmental Settings tracked actions'
            ResetActionTracking -ActionSet EnvironmentalSettings
        } '3' {
            Clear-Host
            'You selected to reset Environment Variables tracked actions'
            ResetActionTracking -ActionSet EnvironmentVariables
        } '4' {
            Clear-Host
            'You selected to reset External Tasks tracked actions'
            ResetActionTracking -ActionSet ExternalTasks    
        } '5' {
            Clear-Host
            'You selected to reset File Associations tracked actions'
            ResetActionTracking -ActionSet FileAssociations
        } '6' {
            Clear-Host
            'You selected to reset File System tracked actions'
            ResetActionTracking -ActionSet FileSystemOperations
        } '7' {
            Clear-Host
            'You selected to reset Groups tracked actions'
            ResetActionTracking -ActionSet Groups
        } '8' {
            Clear-Host
            'You selected to reset Ini Files Operations tracked actions'
            ResetActionTracking -ActionSet IniFilesOperations
        } '9' {
            Clear-Host
            'You selected to reset Init Tasks tracked actions'
            ResetActionTracking -ActionSet InitTasks
        } '10' {
            Clear-Host
            'You selected to reset Network Drives tracked actions'
            ResetActionTracking -ActionSet NetworkDrives
        } '11' {
            Clear-Host
            'You selected to reset Ports tracked actions'
            ResetActionTracking -ActionSet Ports
        } '12' {
            Clear-Host
            'You selected to reset Printers tracked actions'
            ResetActionTracking -ActionSet Printers
        } '13' {
            Clear-Host
            'You selected to reset Registry Values tracked actions'
            ResetActionTracking -ActionSet RegistryValues
        } '14' {
            Clear-Host
            'You selected to reset System Utilities Settings tracked actions'
            ResetActionTracking -ActionSet SystemUtilitiesSettings
        } '15' {
            Clear-Host
            'You selected to reset System Utilities Settings Groups tracked actions'
            ResetActionTracking -ActionSet SystemUtilitiesSettingsGroups
        }'16' {
            Clear-Host
            'You selected to reset User DSN tracked actions'
            ResetActionTracking -ActionSet UserDsns
        } '17' {
            Clear-Host
            'You selected to reset User Self Assigned Applications tracked actions'
            ResetActionTracking -ActionSet UserSelfAssignedApplications
        } '18' {
            Clear-Host
            'You selected to reset Usv User Configuration Settings (Folder Redirection) tracked actions'
            ResetActionTracking -ActionSet UsvUserConfigurationSettings
        } '19' {
            Clear-Host
            'You selected to reset Virtual Drives tracked actions'
            ResetActionTracking -ActionSet VirtualDrives
        } '20' {
            Clear-Host
            Write-Warning 'You selected to reset user self assigned default printer tracked actions'
            ResetActionTracking -ActionSet UserPrinters           
        } '21' {
            Clear-Host
            Write-Warning 'You selected to reset All WEM tracked actions'
            ResetActionTracking -ActionSet All
        } '22' {
            Clear-Host
            'You selected to refresh the WEM cache'
            Start-Process 'C:\Program Files (x86)\Norskale\Norskale Agent Host\AgentCacheUtility.exe' -ArgumentList "-refreshcache -BrokerName $($WEMBroker)"
        } 'q' {
            return
        }
    }
    pause
}
until ($input -eq 'q')



