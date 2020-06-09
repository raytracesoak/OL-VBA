Sub Sheet_to_PDF()

    Sheets(1).Paste
    Sheets(1).Pictures(1).ShapeRange.ScaleHeight 0.5, msoFalse, msoScaleFromTopLeft
    file_path = ThisWorkbook.Path & "/1.pdf"

    Sheets(1).ExportAsFixedFormat Type:=xlTypePDF, _
        Filename:=file_path, _
        Quality:=xlQualityStandard, _
        IncludeDocProperties:=True, _
        IgnorePrintAreas:=True, _
        OpenAfterPublish:=False
    
End Sub