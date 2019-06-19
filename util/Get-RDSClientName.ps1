<#
  .SYNOPSIS
    Returns the RDS client name
 
  .DESCRIPTION
    Returns the value of HKCU:\Volatile Environment\<SessionID>\CLIENTNAME
 
  .EXAMPLE
    Get-RDSClientName -SessionId 4
 
  .EXAMPLE
    Get-RDSClientName -SessionId Get-RDSSessionId
 
  .OUTPUTS
    System.String
#>
function Get-RDSClientName
{
  [CmdletBinding()]
  Param
  (
  # Identifies a RDS session ID
    [Parameter(Mandatory = $true, ValueFromPipeline = $true)]
    [System.String] 
    $SessionId
  )
  $returnValue = $null
  $regKey = 'HKCU:\Volatile Environment\{0}' -f $SessionId
  try
  {
    $ErrorActionPreference = 'Stop'
    $regKeyValues = Get-ItemProperty $regKey
    $sessionName = $regKeyValues | ForEach-Object {$_.SESSIONNAME}
    if ($sessionName -ne 'Console')
    {
      $returnValue = $regKeyValues | ForEach-Object {$_.CLIENTNAME}
    }
    else
    {
      Write-Warning 'Console session'
#     $returnValue = $env:COMPUTERNAME
    }
  }
  catch
  {
    $_.Exception | Write-Error
  }
  New-Object psobject $returnValue
}