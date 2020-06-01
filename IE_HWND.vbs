Sub IE_HWND()

    hWnd_1 = FindWindow("IEFrame", vbNullString)
    hWnd_2 = FindWindowEx(hWnd_1, 0, "Frame Tab", vbNullString)
    hWnd_3 = FindWindowEx(hWnd_2, 0, "TabWindowClass", vbNullString)
    hWnd_4 = FindWindowEx(hWnd_3, 0, "Shell DocObject View", vbNullString)
    hWnd_5 = FindWindowEx(hWnd_4, 0, "Internet Explorer_Server", vbNullString)
    ScreenDump (hWnd_5)
    DoEvents
    SavePicture ApiGetClipBmp, ThisWorkbook.Path & "\" & Format(Now, "yyyymmddhhnnss") & ".bmp"

End Sub