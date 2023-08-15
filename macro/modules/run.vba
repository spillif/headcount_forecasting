Option Explicit
Dim lCalcSave As Long
Dim bScreenUpdate As Boolean
Sub SwitchOff(bSwitchOff As Boolean)
  Dim ws As Worksheet
    
  With Application
    If bSwitchOff Then

      ' OFF
      lCalcSave = .Calculation
    bScreenUpdate = .ScreenUpdating
      .Calculation = xlCalculationManual
      .ScreenUpdating = False
      .EnableAnimations = False
      
      '
      ' switch off display pagebreaks for all worksheets
      '
      For Each ws In ActiveWorkbook.Worksheets
        ws.DisplayPageBreaks = False
      Next ws
    Else
 
      ' ON
      If .Calculation <> lCalcSave And lCalcSave <> 0 Then .Calculation = lCalcSave
      .ScreenUpdating = bScreenUpdate
      .EnableAnimations = True
      
    End If
  End With
End Sub

Sub Historical()
Dim msg As Integer
SwitchOff (True) 'turn off these features

'Call updatehistorical
'Call loopingyear

msg = MsgBox("The Historical Data Has Been Update", vbOKOnly, "Workload Allocation")

SwitchOff (False) 'turn these features back on
End Sub

Sub Forecast()
Dim msg As Integer
SwitchOff (True) 'turn off these features

'Call updateforecast
'Call loopingyear
msg = MsgBox("The Forecast Data Has Been Update", vbOKOnly, "Workload Allocation")

SwitchOff (False) 'turn these features back on
End Sub

Sub BOH()
Dim msg As Integer
SwitchOff (True) 'turn off these features

'Call updateboh
'call loopingweek
msg = MsgBox("The Forecast Data Has Been Update", vbOKOnly, "Workload Allocation")

SwitchOff (False) 'turn these features back on
End Sub

Sub Actual()
Dim msg As Integer
SwitchOff (True) 'turn off these features

Call updateactual
Call loopingweek
msg = MsgBox("The Actual Data Has Been Update", vbOKOnly, "Workload Allocation")

SwitchOff (False) 'turn these features back on
End Sub

Sub ExecutionBOH()
Dim msg As Integer
SwitchOff (False) 'turn off these features
frBoxVolUpdate.Show

End Sub

Sub Continue()
SwitchOff (True) 'turn off these features

Call checkingvalue
MsgBox "Complete"

SwitchOff (False) 'turn these features back on
End Sub
