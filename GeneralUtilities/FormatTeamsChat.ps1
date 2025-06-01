<#
.Synopsis
    Format-TeamsChat produces a more readable version of a Microsoft Teams conversation transcript.
.Description
    Format-TeamsChat matches the raw Microsoft Teams transcript against a set of regular expression rules.
    These rules alter or remove some of the text, producing a more readable version of the conversation.
.Parameter content
    A string element containing the Microsoft Teams conversation messages and their metadata.
#>
function Format-TeamsChat {
    [CmdletBinding()]
    param ($content)
    [string]$displayName = "(?:[\w'-]+ ?){1,2}, (?:\b[\w'-]+\b *){0,2}(?:\(\b[\w+ _&-]+\))"
    [string]$timestamp = "(?:(?:\r\n)*(?:Yesterday|Monday|Tuesday|Wednesday|Thursday|Friday|Saturday|Sunday|\d{1,2}/\d{1,2}| )*\d{1,2}:\d{1,2} (?:AM|PM))?"
    [string]$reply_txt = "(?<ref>" + $displayName + "\r\n(?:\d{1,2}/){2}\d{4} \d{1,2}:\d{1,2} (?:AM|PM))"
    [string]$byLine_txt = "(?:.{1,45})(?:(\.{3}))? by " + $displayName + "(?:, has an attachment.)?"
    [string]$newComment_txt = "(?<header>" + $timestamp + "\r\n" + $displayName + $timestamp  + ")(?:\r\n)*"

    [regex]$reply = $reply_txt
    [regex]$byLine = $byLine_txt
    [regex]$newComment = $newComment_txt
    [regex]$reactionCount = "(?<=\d+ [\w\s!]+ reaction[s]?\.)[\r\n]+\d+"
    [regex]$lastRead = "Last read[\r|\n]+"
    [regex]$contextMenu = "[Hh]as context menu\.?"
    [regex]$date = "(Monday|Tuesday|Wednesday|Thursday|Friday|Saturday|Sunday), (January|February|March|April|May|June|July|August|September|October|November|December) \d{1,2}"
    [regex]$condenseBlankLines = "(\r*\n){3,}"
    [regex]$newLine = "(?<!\r)\n"

    # Ensure carriage returns precede each line feed character for proper formatting in C$ text box.
    $content = $content -replace $newLine, "`r`n"

    # Indicate message is in response to another message.
    $content = [regex]::Replace($content, $reply, {
        param($match)
        $ref = $match.Groups['ref']
        $replacement = "responding to:`r`n$ref"
        return $replacement
    })

    # Format the chat transcript by editing content that matches the regular expressions defined above.
    $null = Switch -regex ($content)
    {
        # Remove all preview lines containinng "by <userLast>, [<userfirst>] (<Department>)"
        $byLine { $content = $content -replace $byLine, "" }

        # Remove redundant digit following reaction indicators.
        $reactionCount { $content = $content -replace $reactionCount, "" }

        # Remove line that says "last read" if it exists. It simply indicates the last line of the chat that was read by the signed-in Teams user.
        $lastRead { $content = $content -replace $lastRead, "" }

        # Remove "Has context menu" line.
        $contextMenu { $content = $content -replace $contextMenu, "" }

        default { $_ }
    }

    # Separate new messages by adding a new line above the commenter's name.
    # With Teams in Comfy mode, date and time appear before the name of the signed in user, but after the name of all other users.
    # With Teams in Compact mode, only the time appears, but always after the name of the user. It doesn't appear at all if it is a sequential post by the same user.
    $content = $content -replace $newComment, $("`r`n" + '${header}' + "`r`n")  

    # Add newline above new date announcement.
    $content = $content -replace $date, "`r`n$&"

    # Replace multiple newlines with a single newline character.
    # Also remove blank lines at the beginning and end of the transcript.
    $content = ($content -replace $condenseBlankLines, "`r`n`r`n").Trim()

    return $content
}
Export-ModuleMember -Function Format-TeamsChat