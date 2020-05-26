Attribute VB_Name = "Ä£¿é1"
Private Declare PtrSafe Sub keybd_event Lib "user32" (ByVal bVk As Byte, ByVal bScan As Byte, ByVal dwFlags As Long, ByVal dwExtraInfo As Long)

Sub sdhjsdh()

    For i = 1 To 3
        ThisWorkbook.FollowHyperlink "https://www.google.com.hk/search?q=" & i & "&btnG=Search&safe=active&gbv=1"
    Next




End Sub



Sub shelll()

    For i = 1 To 3
        uurrll = "explorer.exe ""https://www.google.com/search?q=" & i & "&btnG=Search&safe=active&gbv=1"""
        Shell uurrll, vbNormalFocus
    Next


End Sub



Sub sds()

AppActivate "1 - Google Search"

Call keybd_event(vbKeySnapshot, 1, 0, 0)



End Sub
