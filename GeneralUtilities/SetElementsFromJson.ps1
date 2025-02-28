<#
.Synopsis
  Creates and configures each individual form element and adds them to the Windows form.
.Description
  Modified from the original file by Rod Meany (see https://devblogs.microsoft.com/powershell-community/simple-form-development-using-powershell/)
  Configures each form element based on the data specified by the function Set-FormFromJson, and attaches the control elements to the form.
.Parameter Form
  The main Windows form object.
.Parameter FormHash
  Hashtable of each form tab data.
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
    if ($el.FontSize) { $FontSize = $el.Fontsize } else {$FontSize = 9}
    $FontStyle = [System.Drawing.Font]::new("Tahoma", $FontSize, [System.Drawing.FontStyle]::Regular)
    if ($el.FontStyle) {
      Switch (($el.FontStyle).ToUpper()) {
        "BOLD" { $FontStyle = [System.Drawing.Font]::new("Tahoma", $FontSize, [System.Drawing.FontStyle]::Bold) }
        "ITALIC" { $FontStyle = [System.Drawing.Font]::new("Tahoma", $FontSize, [System.Drawing.FontStyle]::Italic) }
        "STRIKEOUT" { $FontStyle = [System.Drawing.Font]::new("Tahoma", $FontSize, [System.Drawing.FontStyle]::Strikeout) }
        "UNDERLINE" { $FontStyle = [System.Drawing.Font]::new("Tahoma", $FontSize, [System.Drawing.FontStyle]::Underline) }
        Default { $FontStyle = [System.Drawing.Font]::new("Tahoma", $FontSize, [System.Drawing.FontStyle]::Regular) }
      }
    }
    switch ($el.Type) {
      "Label" {
        $Label = New-Object System.Windows.Forms.Label
        $Label.Text = $el.Text
        $Label.Location = New-Object System.Drawing.Point($el.x, $el.y)
        $Label.AutoSize = $true
        $Label.Font = $FontStyle
        if ($el.Anchors) {
          $Label.Anchor = Set-Anchors -anchorList $el.Anchors
        } else {$Label.Anchor = [System.Windows.Forms.AnchorStyles]::None}
        $FormHash.Add($el.Name, $Label)
        $Form.Controls.Add($Label)
      }
      "LinkLabel" {
        #Note, you need to do Add-Click with the URL in the Form iteself
        $LinkLabel = New-Object System.Windows.Forms.LinkLabel
        $LinkLabel.Text = $el.Text
        $LinkLabel.Location = New-Object System.Drawing.Point($el.x, $el.y)
        $LinkLabel.AutoSize = $true
        $LinkLabel.Font = $FontStyle
        $LinkLabel.LinkColor = "BLUE"
        $LinkLabel.ActiveLinkColor = "RED"
        $FormHash.Add($el.Name, $LinkLabel)
        $Form.Controls.Add($LinkLabel)
      }            
      "ComboBox" {
        $ComboBox = New-Object System.Windows.Forms.ComboBox
        $ComboBox.Width = $el.Width
        $ComboBox.Location = New-Object System.Drawing.Point($el.x, $el.y)
        if ($el.Monospace) {
          $ComboBox.Font = New-Object System.Drawing.Font("Courier New", $FontSize, [System.Drawing.FontStyle]::Regular)
        }
        else {
          $ComboBox.Font = $FontStyle
        } 
        #Items is optional, add lookup in ~Form.ps1
        if ($el.Items) { $el.Items | ForEach-Object { [void] $ComboBox.Items.Add($_) } }
        if ($el.SelectedIndex) {$ComboBox.SelectedIndex = [int]$el.SelectedIndex}
        $FormHash.Add($el.Name, $ComboBox)
        $Form.Controls.Add($ComboBox)
      }
      "Button" {
        $Button = New-Object System.Windows.Forms.Button
        $Button.Text = $el.Text
        $Button.Location = New-Object System.Drawing.Size($el.x, $el.y)
        $Button.Size = New-Object System.Drawing.Size($el['Size-x'], $el['Size-y'])
        $Button.Font = $FontStyle
        if ($el.Anchors) {$Button.Anchor = Set-Anchors -anchorList $el.Anchors}
        $FormHash.Add($el.Name, $Button)
        $Form.Controls.Add($Button)
      }
      "ListView" {
        #https://info.sapien.com/index.php/guis/gui-controls/spotlight-on-the-listview-control
        #https://www.sapien.com/blog/2012/04/04/spotlight-on-the-listview-control-part-1/
        #https://www.sapien.com/blog/2012/04/05/spotlight-on-the-listview-control-part-2/
        $ListView = New-Object System.Windows.Forms.ListView
        $ListView.Location = New-Object System.Drawing.Point($el.x, $el.y)
        $ListView.Width = $el.Width
        $ListView.Height = $el.Height  
        $ListView.GridLines = $true
        $ListView.View = [System.Windows.Forms.View]::Details
        $ListView.FullRowSelect = $true

        $ListView.Columns.Add($el.Item, -2, [System.Windows.Forms.HorizontalAlignment]::Left) | Out-Null
        foreach ($col in $el.SubItems) {
          $ListView.Columns.Add($col, -2, [System.Windows.Forms.HorizontalAlignment]::Left) | Out-Null
        }

        $FormHash.Add($el.Name, $ListView)
        $Form.Controls.Add($ListView)
      }
      "ListBox" {
        $listBox = New-Object System.Windows.Forms.ListBox
        $listBox.Location = New-Object System.Drawing.Point($el.x, $el.y)
        $listBox.Width = $el.Width
        $listBox.Height = $el.Height
        $listBox.Font = New-Object System.Drawing.Font("Courier New", $FontSize, [System.Drawing.FontStyle]::Regular)
        if ($el.SelectionMode) { $listBox.SelectionMode = "$($el.SelectionMode)" }
        if ($el.Items) { $el.Items | ForEach-Object { [void] $listBox.Items.Add($_) } } 
        $FormHash.Add($el.Name, $listBox)
        $Form.Controls.Add($listBox)
      }
      "TextBox" {
        $textBox = New-Object System.Windows.Forms.TextBox
        $textBox.Location = New-Object System.Drawing.Point($el.x, $el.y)
        $textBox.Size = New-Object System.Drawing.Size($el['Size-x'], $el['Size-y'])
        $textBox.Font = $FontStyle
        $textBox.Text = $el.Text 
        if ($el.Anchors) {$textBox.Anchor = Set-Anchors -anchorList $el.Anchors}
        if ($el.Multiline) {$textBox.Multiline = $el.Multiline}
        if ($el.PasswordChar) {$textBox.PasswordChar = $el.PasswordChar}
        if ($el.DefaultText) {$textBox.Text = $el.DefaultText}
        if ($el.Scrollbars) {$textBox.Scrollbars = $el.Scrollbars}
        $FormHash.Add($el.Name, $textBox)
        $Form.Controls.Add($textBox)
      }
      "CheckBox" {
        $checkBox = New-Object System.Windows.Forms.CheckBox
        $checkBox.Location = New-Object System.Drawing.Point($el.x, $el.y)
        $checkBox.AutoSize = $false
        $checkBox.Text = $el.Text
        $checkBox.Font = $FontStyle 
        if ($el.Checked) {$checkBox.Checked = $el.Checked} else {$checkBox.Checked = $false}
        $FormHash.Add($el.Name, $checkBox)
        $Form.Controls.Add($checkBox)
      }
      "Calendar" {
        $Cal = New-Object System.Windows.Forms.MonthCalendar
        $Cal.Location = New-Object System.Drawing.Size($el.x, $el.y)
        $Cal.ShowTodayCircle = $true
        $Cal.MaxSelectionCount = 1
        $FormHash.Add($el.Name, $Cal)
        $Form.Controls.Add($Cal)
      }
      Default {
        throw "$($el.Type) is not handled by form code - check your json"
      }
    }
  }
}
Export-ModuleMember -Function Set-ElementsFromJson

function Set-ListViewElementsFromData {
  #BEST to make SQL Result or imported CSV as string-like as possible. i.e. do string conversions in SQL or creation of your csv
  [CmdletBinding()]
  param (
    [Parameter(Mandatory = $true)] [System.Windows.Forms.ListView]$ListView,
    [Parameter(Mandatory = $true, HelpMessage = "Data can be a csv that has been imported, a result set from SQL, or PSCustomObject")] $Data
  )
  $ListView.Items.Clear()
  foreach ($row in $Data) {
    $ColName = $ListView.Columns[0].Text
    $item1 = [System.Windows.Forms.ListViewItem]::new(($row."$ColName"), 0)
    for ($i = 1; $i -lt $ListView.Columns.count; $i++) {
      $ColName = $ListView.Columns[$i].Text
      $item1.SubItems.Add([string]$row."$ColName") | Out-Null 
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
  param([string[]]$anchorList)
  $anchors = 0
  foreach ($anchor in $anchorList) {
    $anchors = $anchors -bor [System.Windows.Forms.AnchorStyles]::$anchor
  }
  return $anchors
}
Export-ModuleMember -Function Set-Anchors