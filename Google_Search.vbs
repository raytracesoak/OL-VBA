Sub Google_Search()

    For Each objWin In CreateObject("Shell.Application").Windows
        Do While objWin.ReadyState <> 4
            DoEvents
        Loop
        If LCase(TypeName(objWin.Document)) = "htmldocument" Then
            If objWin.locationurl Like "*google.com*" Then
                Set FindWin = objWin
                Exit For
            End If
        End If
    Next
    Set objWin = Nothing
    
    With FindWin.Document
        For Each q In .all.tags("INPUT")
            If q.Name = "q" Then
                q.Value = "bear"
            End If
        Next
        For Each b In .all.tags("INPUT")
            If b.Type = "submit" Then
                b.Click
                Exit For
            End If
        Next
    End With
    
    Do While FindWin.ReadyState = 4
        DoEvents
    Loop
    Do While FindWin.ReadyState <> 4
        DoEvents
    Loop
    
    Debug.Print "Done"

End Sub