$forms = $PSScriptRoot + "\Forms\Forms.psm1"
$generalUtilities = $PSScriptRoot + "\GeneralUtilities\GeneralUtilities.psm1"

Import-Module -Name $forms, $generalUtilities
Show-TeamsFormatter