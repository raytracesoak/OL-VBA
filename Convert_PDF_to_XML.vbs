Sub Convert_PDF_to_XML()

    Set AcroApp = CreateObject("AcroExch.App")
    Set avdoc = AcroApp.GetActiveDoc
    Set AcroXPDDoc = avdoc.GetPDDoc
    Set jsObj = AcroXPDDoc.GetJSObject

    jsObj.SaveAs ThisWorkbook.Path & "/test.xml", "com.adobe.acrobat.xml-1-00"

    Set AcroApp = Nothing
    Set avdoc = Nothing
    Set AcroXPDDoc = Nothing
    Set jsObj = Nothing

End Sub