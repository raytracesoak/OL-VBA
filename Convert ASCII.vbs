Sub Convert_ASCII()

    HTML_str = Cells(1, 1)
    found_flag = True
    
    Set regex = CreateObject("vbscript.regexp")
    regex.Pattern = "&#\d+;"
    regex.Global = True
    
    If InStr(HTML_str, "&#") > 0 Then
        Set matches = regex.Execute(HTML_str)
        For i = 0 To matches.Count - 1
            HTML_str = Replace(HTML_str, matches(i), ChrW(Mid(matches(i), 3, Len(matches(i)) - 3)))
        Next
    End If

    Cells(1, 2) = HTML_str

End Sub