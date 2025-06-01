<#
.Synopsis
    Creates and configures each individual form element and adds them to the Windows form.
.Description
    Modified from original file created by Rod Meaney (see https://devblogs.microsoft.com/powershell-community/simple-form-development.using-powershell/)
    Configures each form element based on the data specified by the function Set-FormFromJson, and attaches the control elements to the form.
.Parameter Form
    The main windows form object, or the tab object.
.Parameter FormHash
    Hashtable of each tab element.
.Parameter Elements
    Hashtable of each control element to be included in the form.
#>
function Set-ElementsFromJson {
  [CmdletBinding()]
  param (
    [Parameter(Mandatory = $true)] $Form,
    [Parameter(Mandatory = $true)] $FormHash,
    [Parameter(Mandatory = $true)] $Elements
  )
  foreach ($el in $Elements) {
    if ($el.FontSize) { $FontSize = $el.FontSize } else { $FontSize = 9 }
    if ($el.FontName) { $FontName = $el.FontName } else { $FontName = 'Tahoma' }
    $FontStyle = [System.Drawing.Font]::new($FontName, $FontSize, [System.Drawing.FontStyle]::Regular)
    if ($el.FontStyle) {
      Switch (($el.FontStyle).ToUpper()) {
        "BOLD" { $FontStyle = [System.Drawing.Font]::new($FontName, $FontSize, [System.Drawing.FontStyle]::Bold) }
        "ITALIC" { $FontStyle = [System.Drawing.Font]::new($FontName, $FontSize, [System.Drawing.FontStyle]::Italic) }
        "STRIKEOUT" { $FontStyle = [System.Drawing.Font]::new($FontName, $FontSize, [System.Drawing.FontStyle]::Strikeout) }
        "UNDERLINE" { $FontStyle = [System.Drawing.Font]::new($FontName, $FontSize, [System.Drawing.FontStyle]::Underline) }
        Default { $FontStyle = [System.Drawing.Font]::new($FontName, $FontSize, [System.Drawing.FontStyle]::Regular) }
      }
    }
    Switch ($el.Type) {
      "Label" {
        $Label = New-Object System.Windows.Forms.Label
        $Label.Text = $el.Text
        $Label.Location = New-Object System.Drawing.Point($el.x, $el.y)
        $Label.AutoSize = $true
        $Label.Font = $FontStyle
        if ($el.Anchors) { $Label.Anchor = Set-Anchors -AnchorList $el.Anchors }
        if ($el.BackColor) { $Label.BackColor = Set-ElementColor -color $el.BackColor }
        if ($el.ForeColor) { $Label.ForeColor = Set-ElementColor -color $el.ForeColor }
        $FormHash.Add($el.Name, $Label)
        $Form.Controls.Add($Label)
      }
      "LinkLabel" {
        # Note, you need to do Add-Click with the URL in the Form itself
        $LinkLabel = New-Object System.Windows.Forms.LinkLabel
        $LinkLabel.Text = $el.Text
        $LinkLabel.Location = New-Object System.Drawing.Point($el.x, $el.y)
        $LinkLabel.AutoSize = $true
        $LinkLabel.Font = $FontStyle
        $LinkLabel.LinkColor = "BLUE"
        $LinkLabel.ActiveLinkColor = "RED"
        if ($el.Anchors) { $LinkLabel.Anchor = Set-Anchors -AnchorList $el.Anchors }
        if ($el.BackColor) { $LinkLabel.BackColor = Set-ElementColor -color $el.BackColor }
        if ($el.ForeColor) { $LinkLabel.ForeColor = Set-ElementColor -color $el.ForeColor }
        $FormHash.Add($el.Name, $LinkLabel)
        $Form.Controls.Add($LinkLabel)
      }
      "ComboBox" {
        $ComboBox = New-Object System.Windows.Forms.ComboBox
        $ComboBox.Width = $el.Width
        $ComboBox.Location = New-Object System.Drawing.Pont($el.x, $el.y)
        if ($el.Monospace) {
          $ComboBox.Font = New-Object System.Drawing.Font("Courier New", $FontSize, [System.Drawing.FontStyle]::Regular)
        }
        else {
          $ComboBox.Font = $FontStyle
        }
        #Items is optional, and lookup in ~Form.ps1
        if ($el.Items) { $el.Items | ForEach-Object { [void]$ComboBox.Items.Add($_) } }
        if ($el.SelectedIndex) { $ComboBox.SelectedIndex = [int]$el.SelectedIndex }
        if ($el.Anchors) { $ComboBox.Anchor = Set-Anchors -AnchorList $el.Anchors }
        if ($el.BackColor) { $ComboBox.BackColor = Set-ElementColor -color $el.BackColor }
        if ($el.ForeColor) { $ComboBox.ForeColor = Set-ElementColor -color $el.ForeColor }
        $FormHash.Add($el.Name, $ComboBox)
        $Form.Controls.Add($ComboBox)
      }
      "Button" {
        $Button = New-Object System.Windows.Forms.Button
        $Button.Text = $el.Text
        $Button.Location = New-Object System.Drawing.Point($el.x, $el.y)
        $Button.Size = New-Object System.Drawing.Size($el['Size-x'], $el['Size-y'])
        $Button.Font = $FontStyle
        if ($el.Anchors) { $Button.Anchor = Set-Anchors -AnchorList $el.Anchors }
        if ($el.BackColor) { $Button.BackColor = Set-ElementColor -color $el.BackColor }
        if ($el.ForeColor) { $Button.ForeColor = Set-ElementColor -color $el.ForeColor }
        $FormHash.Add($el.Name, $Button)
        $Form.Controls.Add($Button)
      }
      "ListView" {
        #https://info.sapien.com/index.php.guis/gui-controls/spotlight-on-the-listview-control
        #https://www.sapien.com/blog/2012/04/04/spotlight-on-the-listview-control-part-1/
        #https://www.sapien.com/blog/2012/04/05/spotlight-on-the-listview-control-part-2/
        $ListView = New-Object System.Windows.Forms.ListView
        $ListView.Location = New-Object System.Drawing.Point($el.x, $el.y)
        $ListView.Width = $el.Width
        $ListView.Height = $el.Height
        $ListView.GridLines = $true
        $ListView.View = [System.Windows.Forms.View]::Details
        $ListView.FullRowSelect = $true
        if ($el.Anchors) { $ListView.Anchor = Set-Anchors -AnchorList $el.Anchors }
        if ($el.BackColor) { $ListView.BackColor = Set-ElementColor -color $el.BackColor }
        if ($el.ForeColor) { $ListView.ForeColor = Set-ElementColor -color $el.ForeColor }

        $ListView.Columns.Add($el.Item, -2, [System.Windows.Forms.HorizontalAlignment]::Left) | Out-Null
        foreach ($col in $el.SubItems) {
          $ListView.Columns.Add($col, -2, [System.Windows.Forms.HorizontalAlignment]::Left) | Out-Null
        }

        $FormHash.Add($el.Name, $ListView)
        $Form.Controls.Add($ListView)
      }
      "ListBox" {
        $ListBox = New-Object System.Windows.Forms.ListBox
        $ListBox.Location = New-Object System.Drawing.Point($el.x, $el.y)
        $ListBox.Width = $el.Width
        $ListBox.Height = $el.Height
        $ListBox.Font = New-Object System.DRawing.Font("Courier New", $FontSize, [System.Drawing.FontStyle]::Regular)
        if ($el.SelectionMode) { $ListBox.SelectionMode = "$(Eel.SelectionMode)" }
        if ($el.Items) { $el.Items | ForEach-Object { [void]$ListBox.Items.Add($_) } }
        if ($el.Anchors) { $ListBox.Anchor = Set-Anchors -AnchorList $el.Anchors }
        if ($el.BackColor) { $ListBox.BackColor = Set-ElementColor -color $el.BackColor }
        if ($el.ForeColor) { $ListBox.ForeColor = Set-ElementColor -color $el.ForeColor }
        $FormHash.Add($el.Name, $ListBox)
        $Form.Controls.Add($ListBox)
      }
      "TextBox" {
        $TextBox = New-Object System.Windows.Forms.TextBox
        $TextBox.Location = New-Object System.Drawing.Point($el.x, $el.y)
        $TextBox.Size = New-Object System.Drawing.Size($el['Size-x'], $el['Size-y'])
        $TextBox.Font = $FontStyle
        $TextBox.Text = $el.Text
        if ($el.MaxLength) { $TextBox.MaxLength = $el.MaxLength } else { $TextBox.MaxLength = 32767 }  
        if ($el.Multiline) { $TextBox.Multiline = $el.Multiline }
        if ($el.PasswordChar) { $TextBox.PasswordChar = $el.PasswordChar }
        if ($el.DefaultText) { $TextBox.Text = $el.DefaultText }
        if ($el.Scrollbars) { $TextBox.SCrollbars = $el.Scrollbars }
        if ($el.Anchors) { $TextBox.Anchor = Set-Anchors -AnchorList $el.Anchors }
        if ($el.BackColor) { $TextBox.BackColor = Set-ElementColor -color $el.BackColor }
        if ($el.ForeColor) { $TextBox.ForeColor = Set-ElementColor -color $el.ForeColor }
        $FormHash.Add($el.Name, $TextBox)
        $Form.Controls.Add($TextBox)
      }
      "CheckBox" {
        $CheckBox = New-Object System.Windows.Forms.CheckBox
        $CheckBox.Location = New-Object System.Drawing.Point($el.x, $el.y)
        $CheckBox.AutoSize = $false
        $CheckBox.Text = $el.Text
        $CheckBox.Font = $FontStyle
        if ($el.Checked) { $CheckBox.Checked = $el.Checked } else { $CheckBox.Checked = $false }
        if ($el.Anchors) { $CheckBox.Anchor = Set-Anchors -AnchorList $el.Anchors }
        if ($el.BackColor) { $CheckBox.BackColor = Set-ElementColor -color $el.BackColor }
        if ($el.ForeColor) { $CheckBox.ForeColor = Set-ElementColor -color $el.ForeColor }
        $FormHash.Add($el.Name, $CheckBox)
        $Form.Controls.Add($CheckBox)
      }
      "Calendar" {
        $Cal = New-Object System.Windows.Forms.MonthCalendar
        $Cal.Location = New-Object System.Drawing.Point($el.x, $el.y)
        $Cal.ShowTodayCircle = $true
        $Cal.MaxSelectionCount = 1
        if ($el.Anchors) { $Cal.Anchor = Set-Anchors -AnchorList $el.Anchors }
        if ($el.BackColor) { $Cal.BackColor = Set-ElementColor -color $el.BackColor }
        if ($el.ForeColor) { $Cal.ForeColor = Set-ElementColor -color $el.ForeColor }
        $FormHash.Add($el.Name, $Cal)
        $Form.Controls.Add($Cal)
      }
      Default {
        throw "$($el.Type) is not handled by form code - check your Json"
      }
    }
  }
}
Export-ModuleMember -Function Set-ElementsFromJson


function Set-ListViewElementsFromData {
  #Best  to make SQL Result or imported CSV as string-like as possible. i.e. do string conversions in SQL or creation of your CSV.
  [CmdletBinding()]
  param (
    [Parameter(Mandatory = $true)] [System.Windows.Forms.ListView]$ListView,
    [Parameter(Mandatory = $true, HelpMessage = "Data can be a csv that has been imported, a result set from SQL, or PSCustomObject.")] $Data
  )
  $ListView.Items.Clear()
  foreach ($row in $Data) {
    $ColName = $ListView.Columns[0].Text
    $item1 = [System.Windows.Forms.ListViewItem]::new(($row."$ColName"), 0)
    for ($i = 1; $i -lt $ListView.Columns.count; i++) {
      $ColName = $ListView.Columns[$i].Text
      $item1.SubItems.Add([string]$row."$Column") | Out-Null
    }
    $ListView.Items.Add($item1) | Out-Null
  }
}
Export-ModuleMember -Function Set-ListViewElementsFromData


function New-LineDataFromListViewItem {
  [CmdletBinding()]
  param (
    [Parameter(Mandatory = $true)] $item
  )
  $data = @{}
  $i = 0
  foreach ($col in $item.ListView.Columns) {
    $data.Add($col.Text, $item.SubItems[$i].Text)
    $i++
  }
  return $data
}
Export-ModuleMember -Function New-LineDataFromListViewItem




function Set-Anchors {
  # Anchors the element to each side of the form specified by the array passed to the function.
  [CmdletBinding()]
  param([string[]]$AnchorList)
  $anchors = $null
  foreach ($anchor in $AnchorList) {
    $anchors = $anchors -bor [System.Windows.Forms.AnchorStyles]::$anchor
  }
  return $anchors
}
Export-ModuleMember -Function Set-Anchors




function Set-DockStyle {
  [CmdletBinding()]
  param([string[]]$DockStyle)
  return [System.Windows.Forms.DockStyle]::$DockStyle
}
Export-ModuleMember -Function Set-DockStyle




function Set-ElementColor {
  [CmdletBinding()]
  param($color)
  return [System.Drawing.Color]::FromArgb($color.R, $color.G, $color.B)
}
Export-ModuleMember -Function Set-ElementColor