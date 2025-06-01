<#
.Synopsis
    Displays a graphical interface for formatting Microsoft Teams chat transcripts.
.Description
    Show-TeamsFormatter function displays a Windows Form consisting of the elements defined in the json file (ShowTeamsFormatter.json).
    The function takes no arguments, but the formatting logic takes as input unformatted transcript text copied from a Microsoft teams conversation and applies formatting rules.
    The input can either be pasted directly into the form, or uploaded from a text file.
.Notes
    Author: Brian Larsen
#>
function Show-TeamsFormatter {
    [CmdletBinding()]
    param()

    # ===== TOP ====
    # The name of the Json file should match the name of the function.
    $FormJson = $PSCommandPath.Replace(".ps1", ".json")
    $NewForm, $FormElements = Set-FormFromJson $FormJson
    $appVersion = "2.0"

    # Resize and reposition controls when the window is resized
    $NewForm.Add_Resize({
        # Current window width and height
        $CurrentWidth = $NewForm.Width
        $CurrentHeight = $Newform.Height

        # Set new size and location of Textboxes and Input/Output textbox labels.
        $NewTextBoxSize = @((($currentWidth / 2) - 20), ($($currentHeight - 202)))
        $FormElements.tab_Home.textbox_Input.Size = (New-Object -TypeName System.Drawing.Size -ArgumentList $NewTextBoxSize)
        $FormElements.tab_Home.textbox_Output.Size = (New-Object -TypeName System.Drawing.Size -ArgumentList $NewTextBoxSize)
        $NewOutputTextBoxLocation = @(($($currentWidth / 2) - 9), 30)
        $FormElements.tab_Home.textbox_Output.Location = (New-Object -TypeName System.Drawing.Point -ArgumentList $NewOutputTextBoxLocation)
        
        $NewInputLabelX = ($CurrentWidth * .25) - 30
        $NewOutputLabelX = ($CurrentWidth * .75) - 50
        $FormElements.tab_Home.label_InputText.Location = (New-Object -TypeName System.Drawing.Point -ArgumentList @($NewInputLabelX,6))
        $FormElements.tab_Home.label_OutputText.Location = (New-Object -TypeName System.Drawing.Point -ArgumentList @($NewOutputLabelX,6))
    })

    # Set the default output file based on Json configuration file.
    $configData = Get-Content $FormJson -Raw | ConvertFrom-Json
    [string]$script:defaultOutputFile = $configData.Form.$script:defaultOutputFile
    
    # If no file is specified, set the default location to TeamChat.txt in user's local Documents folder
    if ([string]::IsNullOrEmpty($script:defaultOutputFile))
    {
        $script:defaultOutputFile = "$env:userprofile\Documents\Teamschat.txt"
    }

    # Add version number to the window title
    $NewForm.Text = $configData.Form.Text + $appVersion

    # ==== Main tab elements START ====

    # Clear content from the Input box.
    $FormElements.tab_Home.button_ClearInput.Add_Click({
        $FormElements.tab_Home.textBox_Input.Text = ""
        Write-Host "Input text deleted."
    })

    # Upload an existing transcript file.
    $FormElements.tab_Home.button_ImportFile.Add_Click({
        $transcriptFile = Get-TranscriptFile
        if ([string]::IsNullOrEmpty($transcriptFile)) {
            Write-Warning "No input file selected."
        }
        else {
            $FormElements.tab_Home.textbox_Input.Text = @(Get-Content -Path $transcriptFile -Raw)
            Write-Host "File imported."
        }
    })

    # Format the text in the input box and place the result in the output box.
    $FormElements.tab_Home.button_FormatText.Add_Click({
        $content = $FormElements.tab_Home.textbox_Input.Text
        $content = Format-TeamsChat $content

        # Insert the formattedtext in to the output box.
        $FormElements.tab_Home.textbox_Output.Text = $content
        Write-Host "Text formatted."
    })

    # Copy the contents of the Output box to the clipboard
    $FormElements.tab_Home.button_CopyToClipboard.Add_Click({
        if ($FormElements.tab_Home.textbox_Output.Text) {
            Set-Clipboard -Value $FormElements.tab_Home.textbox_Output.Text
            Write-Host "Copied to clipboard."
        }
        else {Write-Warning "Copy to clipboard failed - No content found."}
    })

    # Save the file to the specified location, creating the file if it doesn't exist.
    $FormElements.tab_Home.button_OutputFile.Add_Click({
        $file = $script:defaultOutputFile
        if (Test-Path -Path $file -IsValid)
        {
            if (!(Test-Path -Path $file)) {New-Item -Path $file -ItemType "file" -Force | Out-Null}
            $content = $FormElements.tab_Home.textbox_Output.Text
            $content | Out-File -FilePath $file -Force
            Write-Host "Saved to: $file"
        }
        else {Write-Host "Invalid output file path."}
    })
    # ==== Main tab elements END ====

    # ==== Settings tab elements START ====
    # Display the current download file path
    $FormElements.tab_Settings.label_DefaultOutput.Text = "Output File: `n" + $script:defaultOutputfile

    # Open a save file dialog and select the location and file name.
    # Only text files are allowed.
    $FormElements.tab_Settings.button_SetDefaultOutputFile.Add_Click({
        $outputFile = Set-OutputFile
        if (!($null -eq $outputFile) -and (Test-Path -Path $outputFile -IsValid))
        {
            # Set the default file to the newly specified path.
            $script:defaultOutputFile = $outputFile

            # Update the label displaying the path in the form.
            $FormElements.tab_Settings.label_DefaultOutput.Text = "Output File: `n" + $script:defaultOutputFile

            # Update the Json config file with the newly specified default output path for persistence across sessions.
            $configData = Get-Content $FormJson -Raw | ConvertFrom-Json
            $configData.Form.DefaultOutputFile = $outputFile
            $configData | ConvertTo-Json -Depth 32 | Set-Content -Path $FormJson -Force
            Write-Host "Updated output file path."
        }
    })
    # ==== Settings tab elements END ====

    # ==== BOTTOM ====
    $null = $NewForm.ShowDialog()
}
Export-ModuleMember -Function Show-TeamsFormatter