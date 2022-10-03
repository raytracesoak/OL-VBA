Sub Change_Mail_Attributes()

    Set oInspector = Application.ActiveInspector
    If oInspector Is Nothing Then
        Set Item = Application.ActiveExplorer.Selection.Item(1)
    Else
        Set Item = oInspector.CurrentItem
    End If

    strSubject = Item.Subject
    Item.Subject = "asd"
    Item.To = "asd@asd.com"
    Item.CC = "asd@asd.com"

End Sub