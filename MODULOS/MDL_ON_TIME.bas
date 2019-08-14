Attribute VB_Name = "MDL_ON_TIME"
Sub ScheduleTheDay()
'Application.OnTime EarliestTime:=TimeValue("10:25 PM"), Procedure:=captureData
Application.OnTime EarliestTime:=TimeValue("9:32 PM"), Procedure:="captureData"
End Sub

Sub captureData()
Dim WSQ As Worksheet
Dim NextRow As Long
Set WSQ = Worksheets("TESTES")
NextRow = WSQ.Cells(Rows.Count, 1).End(xlUp).Row
WSQ.Range("A2:B2").Copy WSQ.Cells(NextRow, 1)
Debug.Print "fIM"

End Sub


