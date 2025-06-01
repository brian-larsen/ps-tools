<#
.Synopsis
    Set-OutputFile opens a save file dialog for the selection of a path to a new or existing text file.
.Description
    the function Set-OutputFile allows for the path to a text file to be selected. the file can be an existing one,
    or a new file name can be entered. The file itself will not be created but the path will be returned.
    the calling function can then create the file if it does not currently exist.
#>
function Set-OutputFile {
    [CmdletBinding()]
    param()
    $fileDialog = New-Object System.Windows.Forms.SaveFileDialog
    $fileDialog.Filter = "Text files (*.txt)|*.txt|All files (*.*)|*.*"
    $fileDialog.Title = "Select File"

    if ($fileDialog.ShowDialog() -eq "OK")
    {
        return $fileDialog.FileName
    }
    else { return $null }
}
Export-ModuleMember -Function Set-OutputFile