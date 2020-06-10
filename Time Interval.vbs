Private Declare PtrSafe Function timeGetTime Lib "winmm.dll" () As Long
Public Sub ddd(intv As Integer)

    Dim Savetime As Double
    
    Savetime = timeGetTime
    Do While timeGetTime < Savetime + intv
        DoEvents
    Loop

End Sub
