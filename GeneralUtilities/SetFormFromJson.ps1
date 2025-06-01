<#
.Synopsis
    Generates a Windows Form object based on the data provided by the Json configuration file.
.Description
    Modified from original file created by Rod Meaney (see https://devblogs.microsoft.com/powershell-community/simple-form-development-using-powershell/.
    Set-FormFromJson function creates the Windows form based on the elements defined in the Json file pass as an argument.
    Then configuration data from the Json file is converted to hashtables for easy processing.
.Parameter JsonFile
    The configuration file containing each element to be included in the form that is displayed.
.Outputs
    main_form       [System.Windows.Forms.Form]     Returns the form object along with the elements included in the form.
    formElements    [System.Collections.Hashtable]  Returns a hashtable with the individual elements to be included in the form.
#>
function Set-FormFromJson {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory=$true)]$JsonFile
    )

    $formElements = @{}
    $configText = (Get-Content -Path $JsonFile -Raw)

    # Convert text data from Json file to PowerShell hashtable data structures.
    $config = $configText | ConvertFrom-Json | ConvertTo-HashtableV5

    # Main window
    $main_form = New-Object System.Windows.Forms.Form
    $main_form.ClientSize =  (New-Object -TypeName System.Drawing.Size -ArgumentList @($config.Form.Width,$config.Form.Height))
    $main_form.MinimumSize = "$($config.form.MinWidth),$($config.Form.MinHeight)"
    $main_form.text = $config.Form.Text

    if ($config.Form.BackColor) {$main_form.BackColor = Set-ElementColor -color $config.Form.BackColor}
    if ($config.Form.ForeColor) {$main_form.ForeColor = Set-ElementColor -color $config.Form.ForeColor}

    # Add tabs to the form if they are specified in the configuration file.
    if ($config.ContainsKey("Tabs")) {
        $FormTabControl = (New-Object -TypeName System.Windows.Forms.TabControl)
        $FormTabControl.Size = "$($config.Form.width),$($config.Form.Height)"
        $FormTabControl.Location = "0,0"
        $FormTabControl.Anchor = Set-Anchors -AnchorList $config.Tabs.Anchors
        $main_form.Controls.Add($FormTabControl)
        foreach ($tab in $config.Tabs) {
            $Tab1 = (New-Object -TypeName System.Windows.Forms.TabPage)
            $Tab1.DataBindings.DefaultDataSourceUpdateMode = 0
            $Tab1.UseVisualStyleBackColor = $true
            $Tab1.Name = $tab.Name
            $Tab1.Text = $tab.Text
            $Tab1.Location = (New-Object -TypeName System.Drawing.Point -ArgumentList @($tab.x,$tab.y))
            $Tab1.Size = (New-Object -TypeName System.Drawing.Size -ArgumentList @($tab['Size-x'],$tab['Size-y']))
            if ($tab.BackColor) {$Tab1.BackColor = Set-ElementColor -color $tab.BackColor}
            if ($tab.ForeColor) {$Tab1.ForeColor = Set-ElementColor -color $tab.ForeColor}
            $formElements.Add($tab.Name, @{})
            Set-ElementsFromJson -Form $Tab1 -FormHash $formElements[$tab.Name] -Elements $tab.Elements
            $FormTabControl.Controls.Add($Tab1)
        }
    } else {
        Set-ElementsFromJson -Form $main_form -FormHash $formElements -Elements $config.Elements
    }
    return $main_form, $formElements
}
Export-ModuleMember -Function Set-FormFromJson