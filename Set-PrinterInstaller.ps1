#Requires -Version 5


<#
.SYNOPSIS
    Printer Installer for RDS
.DESCRIPTION
  <Brief description of script>
.PARAMETER <Parameter_Name>
    <Brief description of parameter input required. Repeat this attribute if required>
.INPUTS
  None
.OUTPUTS
  Log file stored in Event Logs (Event Viewer > Applications and Services Logs > Update Network Printers > Event ID 4444)
.NOTES
  Version:        1.5
  Author:         Elliot Schot
  Creation Date:  26/09/2018
  Purpose/Change: fixes to use from network drive

  Version:        1.4
  Author:         Elliot Schot
  Creation Date:  19/09/2018
  Purpose/Change: Reformat Code to work with JSON data input

  Version:        1.3
  Author:         Elliot Schot
  Creation Date:  27/07/2018
  Purpose/Change: Reformat Code into functions

  Version:        1.2
  Author:         Elliot Schot
  Creation Date:  27/07/2018
  Purpose/Change: Reformat Code

  Version:        1.1
  Author:         Cristian Gallardo
  Creation Date:  23/04/2018
  Purpose/Change: Set up to run almost silently [https://stackoverflow.com/questions/1802127/how-to-run-a-powershell-script-without-displaying-a-window?utm_medium=organic&utm_source=google_rich_qa&utm_campaign=google_rich_qa]

  Version:        1.0
  Author:         Cristian Gallardo
  Creation Date:  18/04/2018
  Purpose/Change: Initial script development

.EXAMPLE
  Powershell.exe -ExecutionPolicy Bypass -NonInteractive -WindowStyle Hidden c:\Scripts\Set-PrinterInstaller.ps1
#>




#---------------------------------------------------------[Initialisations]--------------------------------------------------------
#Hide powershell window
Function Hide-Powershell () {
    $t = '[DllImport("user32.dll")] public static extern bool ShowWindow(int handle, int state);'
    add-type -name win -member $t -namespace native
    [native.win]::ShowWindow(([System.Diagnostics.Process]::GetCurrentProcess() | Get-Process).MainWindowHandle, 0) | Out-Null
}

#Call Hide Powershell
Hide-Powershell

#File Paths
$ScriptDir = Split-Path -parent $MyInvocation.MyCommand.Path
$PathRDSClientName = Join-Path -Path $ScriptDir -ChildPath "\util\Get-RDSClientName.ps1"
$PathRDSSessionID = Join-Path -Path $ScriptDir -ChildPath "\util\Get-RDSSessionID.ps1"
 
#IMPORTING MODULES
Import-Module $PathRDSClientName
Import-Module $PathRDSSessionID

#Files imports
[string]$Script:DataFile = Join-Path -Path $ScriptDir -ChildPath "\data\PrinterMaps.json"
#----------------------------------------------------------[Declarations]----------------------------------------------------------
#VARIABLES
Function Get-EnviromentVariables ([Boolean] $value) {
    if ($value -eq $TRUE) {
        Return Get-RDSSessionId | Get-RDSClientName
    } else {
        Return "ITC4437E6DFEE00"
    }
}

#$loggedOnUser = (Get-ItemProperty -path ("HKCU:\Volatile Environment\") -name "USERNAME").USERNAME
$username = $Env:USERNAME
$computername = $Env:COMPUTERNAME
#True is for live, False is for testing
$sessionCLIENTNAME = Get-EnviromentVariables($TRUE)


#Local Variables
[array]$Script:LogTrace = @()
[array]$Script:Printers = @()
[array]$Script:CustomPrinter = @()
[array]$Script:CustomPrinterINFO = @()
[string]$Script:GroupName = $Null
[string]$Script:ClientName = $Null
[string]$Script:DefaultPrinter = $Null
[array]$Script:JSON = $Null


#-----------------------------------------------------------[Functions]------------------------------------------------------------

Function Get-JSONData ([String] $stringJSONFile) {
    try {
        $jsonLocal = Get-Content $stringJSONFile | ConvertFrom-Json
        $Script:LogTrace += ("Successfully imported the JSON File: " + $stringJSONFile)
        Return $jsonLocal

    } catch {
        $ErrorMessage = $_.Exception.Message
        $Script:LogTrace += ($ErrorMessage)
        Return $Null
    }

}
Function Remove-OSCNetworkPrinters {
    $Script:LogTrace += ("Started Remove-OSCNetworkPrinters")

    $NetworkPrinters = Get-WmiObject -Class Win32_Printer | Where-Object {$_.Network}

    If ($null -ne $NetworkPrinters) {
        Try {
            Foreach ($NetworkPrinter in $NetworkPrinters) {
                $NetworkPrinter.Delete()
                $DeletedPrinterName = $NetworkPrinter.Name
                $Script:LogTrace += ("Successfully deleted the network printer: " + $DeletedPrinterName)
            }
        }
        Catch {
            $ErrorMessage = $_.Exception.Message
            $Script:LogTrace += ($ErrorMessage)
        }
    }
    Else {
        $Script:LogTrace += ("Cannot find network printer in the current environment.")
    }
    $Script:LogTrace += ("Finished Remove-OSCNetworkPrinters")
    $Script:LogTrace += ("")
    Return
}

Function Set-DefaultPrinter($PrinterName) {
    $Script:LogTrace += ("Started Set-DefaultPrinter")
    $Script:LogTrace += ("Setting default printer to: " + $PrinterName)
    $Printers = Get-WmiObject -Class Win32_Printer

    Try {
        Write-Verbose "Get the specified printer info."
        $Printer = $Printers | Where {$_.Name -eq "$PrinterName"}

        If ($Printer) {
            $Printer.SetDefaultPrinter() | Out-Null

            $Script:LogTrace += ("Successfully set the default printer to " + $PrinterName)
        }
        Else {
            Write-Warning "Cannot find the specified printer."
            $Script:LogTrace += ("Cannot find the specified printer.")
        }
    }
    Catch {
        $ErrorMessage = $_.Exception.Message
        $Script:LogTrace += ($ErrorMessage)
    }
    $Script:LogTrace += ("Finished Set-DefaultPrinter")
    Return
}

Function Invoke-WriteEventLog ($EvtMessage) {
    #$loggedOnUser = (Get-ItemProperty -path ("HKCU:\Volatile Environment\") -name "USERNAME").USERNAME

    $SessionInfo = "$username connected from $sessionCLIENTNAME"
    $EvtMessage = "$SessionInfo `r`n $EvtMessage"

    $logFileExists = Get-EventLog -list | Where-Object {$_.logdisplayname -eq "Update Network Printers Task"} 
    if (! $logFileExists) {
        New-EventLog -LogName "Update Network Printers Task" -Source "Update Network Printers Task"
    }
    Write-EventLog -LogName "Update Network Printers Task" -Source "Update Network Printers Task" -EventID 4444 -EntryType Information -Message $EvtMessage -Category 1 -RawData 10, 20
    exit
}

Function Get-ClientMatch($InputSelectedClient) {
    $Script:LogTrace += ("Started Get-ClientMatch")
    $result = $FALSE
    foreach ($group in $JSON) {
        foreach ($client in $group.Clients) {
            If ($InputSelectedClient -eq $client) {
                $result = $TRUE

                #Log Values
                $Script:LogTrace += ("Found Client Match in: " + $group.GroupName + " group")
                $Script:LogTrace += ("Printers: " + $group.Printers + " added to Printer array")
                $Script:LogTrace += ("Default Printer: " + $group.DefaultPrinter)



                #Assign Global Values
                $Script:GroupName = $group.GroupName
                $Script:ClientName = $InputSelectedClient
                $Script:DefaultPrinter = $group.DefaultPrinter
                $Script:Printers = $group.Printers

                #Get the Custom Printers for this Client
                Get-CustomPrinters($InputSelectedClient) ($group.CustomPrinters)
            }
        }
    }
    $Script:LogTrace += ("Finished Get-ClientMatch")
    $Script:LogTrace += ("")
    Return $result
}

function Get-CustomPrinters ([String] $InputClientName, [array] $InputGroupCustomPrinters) {
    $Script:LogTrace += ("Started Get-CustomPrinters")
    foreach ($CustomGroup in $InputGroupCustomPrinters) {
        If ($CustomGroup.ClientName -eq $client) {
            $Script:CustomPrinter += $CustomGroup.PrinterName
            $Script:LogTrace += ("Custom Printer added: " + $CustomGroup.PrinterName + " to array")
        }
    }
    $Script:LogTrace += ("Finished Get-CustomPrinters")
    Return
}

function Invoke-ProccessPrinters {
    $Script:LogTrace += ("Started Invoke-ProccessPrinters")
    #Connect CustomPrinters (Epson Docket Printers)
    if ($Script:CustomPrinter.count -gt 0) {
        foreach ($printer in $Script:CustomPrinter) {
            Try {
                add-printer -connectionname $printer -ErrorAction Stop
                $Script:LogTrace += ("Successfully added the custom printer: " + $printer)
            }
            Catch {
                $ErrorMessage = $_.Exception.Message
                $Script:LogTrace += ("Failed adding the custom printer: " + $printer)
                $Script:LogTrace += ($ErrorMessage)
            }
        }
    }
    else {
        $Script:LogTrace += ("No custom printers added")
    }


    if ($Script:Printers.count -gt 0) {
        foreach ($printer in $Script:Printers) {
            Try {
                add-printer -connectionname $printer -ErrorAction Stop
                $Script:LogTrace += ("Successfully added the network printer: " + $printer)
            }
            Catch {
                $ErrorMessage = $_.Exception.Message
                $Script:LogTrace += ("Failed adding the network printer: " + $printer)
                $Script:LogTrace += ($ErrorMessage)   
            } 
        }
    }
    else {
        $Script:LogTrace += ("No network printers added")
    }

    #Set Default Printer
    Set-DefaultPrinter($Script:DefaultPrinter)
    $Script:LogTrace += ("Finished Invoke-ProccessPrinters")
    $Script:LogTrace += ("")
    Return
}

Function Start-LogTrace () {
    $Script:LogTrace += ("============================================ Script Logging Started ============================================")
    $Script:LogTrace += ("Client ID is: " + $sessionCLIENTNAME)
}

Function Complete-LogTrace () {
    $Script:LogTrace += ("============================================ Script Logging Finished ============================================")
    $EventLog = $Script:LogTrace | out-string
    Invoke-WriteEventLog($EventLog)
}

Function Main () {
    Start-LogTrace

    $Script:JSON = Get-JSONData($Script:DataFile)

    #If $Script:JSON this is null it means the JSON file failed to import, so it should log said failure then exit
    if ($Script:JSON -eq $Null) {
        $Script:LogTrace += ("$Script:JSON = $Null")

        Complete-LogTrace
    }

    if (Get-ClientMatch($sessionCLIENTNAME)) {
        $Script:LogTrace += ("Client match found in json file")
        $Script:LogTrace += ("")
        #Remove All Printers
        Remove-OSCNetworkPrinters

        #Process stored data
        Invoke-ProccessPrinters
    }
    else {
        $Script:LogTrace += ("No Client match found in JSON file [" + $DataFile + "]")
        $Script:LogTrace += ("")
    }
    
    Complete-LogTrace

}

#-----------------------------------------------------------[Execution]------------------------------------------------------------


#Start Script
Main