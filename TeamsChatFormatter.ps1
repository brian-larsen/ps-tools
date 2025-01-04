<#
.Description
  Running this script imports the custom module files which individually load the associated commandlets.
  It automatically calls the proper commandlet to launch the Teams Chat Formatter interface.
#>

$forms = $PSScriptRoot + "\Forms\Forms.psm1"
$generalUtilities = $PSScriptRoot + "\GeneralUtilities\GeneralUtilities.psm1"

Import-Module -Name $forms, $generalUtilities
Format-TeamsChat