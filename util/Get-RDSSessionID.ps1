<#
  .SYNOPSIS
    Returns the RDS session ID of a given user.
 
  .DESCRIPTION
    Leverages query.exe session in order to get the given user's session ID.
 
  .EXAMPLE
    Get-RDSSessionId
 
  .EXAMPLE
    Get-RDSSessionId -UserName johndoe
 
  .OUTPUTS
    System.String
#>
function Get-RDSSessionId
{
  [CmdletBinding()]
  Param
  (
  # Identifies a user name (default: current user)
    [Parameter(ValueFromPipeline = $true)]
    [System.String] 
    $UserName = $env:USERNAME
  )
  $returnValue = $null
  try
  {
    $ErrorActionPreference = 'Stop'
    $output = query.exe session $UserName |
      ForEach-Object {$_.Trim() -replace '\s+', ','} |
        ConvertFrom-Csv
    $returnValue = $output.ID
  }
  catch
  {
    $_.Exception | Write-Error
  }
  New-Object psobject $returnValue
}