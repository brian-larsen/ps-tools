<#
.Synopsis
  Generates a Windows Form object based on the data provided by the Json configuration file.
.Description
  Modified from the original file by Rod Meany (see https://devblogs.microsoft.com/powershell-community/simple-form-development-using-powershell/)
  Set-FormFromJson function creates the Windows form based on the elements defined in the Json file passed as an argument.
  The configuration data from the Json file is converted to hashtables for easy processing.
.Parameter JsonFile
  The configuration file containing each element to be included in the form that is displayed.
.Outputs
  main_form     [System.Windows.Forms.Form]     Returns the form object along with the elements included in the form.
  FormElements  [System.Collections.Hashtable]  Returns a hashtable with the individual elements to be included in the form.
#>
function Set-FormFromJson {
  [CmdletBinding()]
  param(
    [Parameter(Mandatory=$true)]$JsonFile
  )

  $FormElements = @{}
  $configText = (Get-Content -Path $JsonFile -Raw)
  # Convert text data from Json file to PowerShell hashtable data structures.
  $config = $configText | ConvertFrom-Json | ConvertTo-HashtableV5

  # Main Window
  $main_form = New-Object System.Windows.Forms.Form
  $main_form.Text = $config.Form.Text
  $main_form.Width = $config.Form.Width
  $main_form.Height = $config.Form.Height
  $main_form.MaximumSize = New-Object System.Drawing.Size($config.Form.Width, $config.Form.Height)
  $main_form.MinimumSize = New-Object System.Drawing.Size($config.Form.Width, $config.Form.Height)
  # $main_form.Autosize = $true

  # Add tabs to the form if they are specified in the configuration file.
  if ($config.ContainsKey("Tabs")) {
    $FormTabControl = NewObject System.Windows.Forms.TabControl
    $FormTabControl.Size = "$($config.Form.Width),$($config.Form.Height)"
    $FormTabControl.Location = "0,0"

    $main_form.Controls.Add($FormTabControl)
    foreach ($tab in $config.Tabs) {
      $Tab1 = New-Object System.Windows.Forms.TabPage
      $Tab1.DataBindings.DefaultDataSourceUpdateMode = 0
      $Tab1.UseVisualStyleBackColor = $true
      $Tab1.Name = $tab.Name
      $Tab1.Text = $tab.Text
      $FormElements.Add($tab.Name, @{})
      Set-ElementsFromJson -Form $Tab1 -FormHash $FormElements[$tab.Name] -Elements $tab.Elements
      $FormTabControl.Controls.Add($Tab1)
    }
  } else {
    Set-ElementsFromJson -Form $main_form -FormHash $FormElements -Elements $config.Elements
  }
  return $main_form, $FormElements
}
Export-ModuleMember -Function Set-FormFromJson