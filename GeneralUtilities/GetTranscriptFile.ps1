Function Get-TranscriptFile {
    <#
    .DESCRIPTION
        Opens a file Explorer dialog for selection of an existing text file.
    #>
    $fileDialog = New-Object System.Windows.Forms.OpenFileDialog
    $fileDialog.ShowDialog() | Out-Null
    if ($fileDialog.DialogResult.Cancel) {
        return $null
    } else {
        $transcriptName = $fileDialog.FileName
        return $transcriptName
    }
}
Export-ModuleMember -Function Get-TranscriptFile