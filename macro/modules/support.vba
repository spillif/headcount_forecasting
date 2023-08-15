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

Sub step1()

Dim wsSource, wsDest, wsSupport As Worksheet
Dim j, LastRow1 As Integer

SwitchOff (True) 'turn off these features

'Set variables
Set wsSource = ActiveWorkbook.Worksheets("SummaryPerWeek") 'Sheet "Summary"
Set wsDest = ActiveWorkbook.Worksheets("raw_support") 'Sheet "raw_support"
Set wsSupport = ActiveWorkbook.Worksheets("Booking_SGN") 'Sheet "Booking_SGN"

'clear old data
wsDest.Activate
wsDest.Range("A2:J1000").ClearContents

'Identify last row and column
wsSource.Activate
LastRow1 = Cells(Rows.Count, 3).End(xlUp).Row

For j = 7 To LastRow1
    
    'get cs air data
    Range(Cells(j, 3), Cells(j, 6)).Copy 'team and cusomters
    wsDest.Range("A500").End(xlUp).Offset(1, 0).PasteSpecial xlPasteValues
    Cells(6, 7).Copy 'air text
    wsDest.Range("E500").End(xlUp).Offset(1, 0).PasteSpecial xlPasteValues
    Cells(5, 9).Copy 'cs text
    wsDest.Range("F500").End(xlUp).Offset(1, 0).PasteSpecial xlPasteValues
    Cells(j, 7).Copy 'air vol
    wsDest.Range("G500").End(xlUp).Offset(1, 0).PasteSpecial xlPasteValues
    Cells(j, 9).Copy 'air fte
    wsDest.Range("H500").End(xlUp).Offset(1, 0).PasteSpecial xlPasteValues
    wsSupport.Range("D3").Copy
    wsDest.Range("I500").End(xlUp).Offset(1, 0).PasteSpecial xlPasteValues
    wsSupport.Range("F3").Copy
    wsDest.Range("J500").End(xlUp).Offset(1, 0).PasteSpecial xlPasteValues
    
Next j

wsDest.Activate

SwitchOff (False) 'turn these features back on

End Sub

Sub step2()

Dim wsSource, wsDest, wsSupport  As Worksheet
Dim j, LastRow1 As Integer

SwitchOff (True) 'turn off these features

'Set variables
Set wsSource = ActiveWorkbook.Worksheets("SummaryPerWeek") 'Sheet "Summary"
Set wsDest = ActiveWorkbook.Worksheets("raw_support") 'Sheet "raw_support"
Set wsSupport = ActiveWorkbook.Worksheets("Booking_SGN") 'Sheet "Booking_SGN"

'Identify last row and column
wsSource.Activate
LastRow1 = Cells(Rows.Count, 3).End(xlUp).Row

For j = 7 To LastRow1

    'get cs sea data
    Range(Cells(j, 3), Cells(j, 6)).Copy 'team and cusomters
    wsDest.Range("A500").End(xlUp).Offset(1, 0).PasteSpecial xlPasteValues
    Cells(6, 8).Copy 'sea text
    wsDest.Range("E500").End(xlUp).Offset(1, 0).PasteSpecial xlPasteValues
    Cells(5, 9).Copy 'cs text
    wsDest.Range("F500").End(xlUp).Offset(1, 0).PasteSpecial xlPasteValues
    Cells(j, 8).Copy 'sea vol
    wsDest.Range("G500").End(xlUp).Offset(1, 0).PasteSpecial xlPasteValues
    Cells(j, 10).Copy 'sea fte
    wsDest.Range("H500").End(xlUp).Offset(1, 0).PasteSpecial xlPasteValues
    wsSupport.Range("D3").Copy
    wsDest.Range("I500").End(xlUp).Offset(1, 0).PasteSpecial xlPasteValues
    wsSupport.Range("F3").Copy
    wsDest.Range("J500").End(xlUp).Offset(1, 0).PasteSpecial xlPasteValues
Next j

wsDest.Activate

SwitchOff (False) 'turn these features back on

End Sub

Sub step3()

Dim wsSource, wsDest, wsSupport  As Worksheet
Dim j, LastRow1 As Integer
Set wsSupport = ActiveWorkbook.Worksheets("Booking_SGN") 'Sheet "Booking_SGN"

SwitchOff (True) 'turn off these features

'Set variables
Set wsSource = ActiveWorkbook.Worksheets("SummaryPerWeek") 'Sheet "Summary"
Set wsDest = ActiveWorkbook.Worksheets("raw_support") 'Sheet "raw_support"

'Identify last row and column
wsSource.Activate
LastRow1 = Cells(Rows.Count, 3).End(xlUp).Row

For j = 7 To LastRow1
    
    'get doc air data
    Range(Cells(j, 3), Cells(j, 6)).Copy 'team and cusomters
    wsDest.Range("A500").End(xlUp).Offset(1, 0).PasteSpecial xlPasteValues
    Cells(6, 7).Copy 'air text
    wsDest.Range("E500").End(xlUp).Offset(1, 0).PasteSpecial xlPasteValues
    Cells(5, 12).Copy 'doc text
    wsDest.Range("F500").End(xlUp).Offset(1, 0).PasteSpecial xlPasteValues
    Cells(j, 7).Copy 'air vol
    wsDest.Range("G500").End(xlUp).Offset(1, 0).PasteSpecial xlPasteValues
    Cells(j, 12).Copy 'air fte
    wsDest.Range("H500").End(xlUp).Offset(1, 0).PasteSpecial xlPasteValues
    wsSupport.Range("D3").Copy
    wsDest.Range("I500").End(xlUp).Offset(1, 0).PasteSpecial xlPasteValues
    wsSupport.Range("F3").Copy
    wsDest.Range("J500").End(xlUp).Offset(1, 0).PasteSpecial xlPasteValues
    
Next j

wsDest.Activate

SwitchOff (False) 'turn these features back on

End Sub

Sub step4()

Dim wsSource, wsDest, wsSupport  As Worksheet
Dim j, LastRow1 As Integer
Set wsSupport = ActiveWorkbook.Worksheets("Booking_SGN") 'Sheet "Booking_SGN"

SwitchOff (True) 'turn off these features

'Set variables
Set wsSource = ActiveWorkbook.Worksheets("SummaryPerWeek") 'Sheet "Summary"
Set wsDest = ActiveWorkbook.Worksheets("raw_support") 'Sheet "raw_support"

'Identify last row and column
wsSource.Activate
LastRow1 = Cells(Rows.Count, 3).End(xlUp).Row

For j = 7 To LastRow1

    'get doc sea data
    Range(Cells(j, 3), Cells(j, 6)).Copy 'team and cusomters
    wsDest.Range("A500").End(xlUp).Offset(1, 0).PasteSpecial xlPasteValues
    Cells(6, 8).Copy 'sea text
    wsDest.Range("E500").End(xlUp).Offset(1, 0).PasteSpecial xlPasteValues
    Cells(5, 12).Copy 'doc text
    wsDest.Range("F500").End(xlUp).Offset(1, 0).PasteSpecial xlPasteValues
    Cells(j, 8).Copy 'sea vol
    wsDest.Range("G500").End(xlUp).Offset(1, 0).PasteSpecial xlPasteValues
    Cells(j, 13).Copy 'sea fte
    wsDest.Range("H500").End(xlUp).Offset(1, 0).PasteSpecial xlPasteValues
    wsSupport.Range("D3").Copy
    wsDest.Range("I500").End(xlUp).Offset(1, 0).PasteSpecial xlPasteValues
    wsSupport.Range("F3").Copy
    wsDest.Range("J500").End(xlUp).Offset(1, 0).PasteSpecial xlPasteValues
    
Next j

wsDest.Activate

SwitchOff (False) 'turn these features back on

End Sub

Sub process()

SwitchOff (True) 'turn off these features

Call step1
Call step2
Call step3
Call step4

SwitchOff (False) 'turn these features back on

End Sub

Sub ftecounting()

Dim wsSource, wsDest As Worksheet

SwitchOff (True) 'turn off these features

'set variables
Set wsSource = ActiveWorkbook.Worksheets("raw_support")
Set wsDest = ActiveWorkbook.Worksheets("fte_data")
wsSource.Activate

'copy and paste to dest
wsSource.Range("A2", Range("A2").End(xlToRight).End(xlDown)).Copy
wsDest.Range("A" & Rows.Count).End(xlUp).Offset(1).PasteSpecial xlPasteValues

SwitchOff (False) 'turn these features back on

End Sub

Sub openwb()

SwitchOff (True) 'turn off these features

Call setpulicproperty

SwitchOff (False) 'turn these features back on

End Sub

Sub closewb()

SwitchOff (True) 'turn off these features

wbg1.Close
wbg2.Close
wbg3.Close
wbg4.Close
wbg5.Close

SwitchOff (False) 'turn these features back on

End Sub

Sub loopingyear()

Dim wsSource, wsDest As Worksheet
Dim sWeek As Range
Dim j As Integer
Dim StartTime As Single
Dim EndTime As Single

SwitchOff (True) 'turn off these features
'Get the start time
StartTime = Timer

'Call openwb

'set variable
Set wsSource = Workbooks("Test Auto.xlsb").Worksheets("Booking_SGN")
wsSource.Activate
Set wsDest = ActiveWorkbook.Worksheets("fte_data")
Set sWeek = wsSource.Range("C5", Range("C5").End(xlToRight))

For j = 51 To 53 'sWeek.Cells.Count
    
        wsSource.Range("C3") = j
        ActiveWorkbook.Save
        Call process
        Call ftecounting
        ActiveWorkbook.Save
    Next j
    
'Call closewb
'wsSource.Range("C3") = "=WEEKNUM(TODAY(),21)"
'wsDest.Activate
    
'Get the end time
EndTime = Timer
'Print the processing time
Debug.Print "Processing time: " & (EndTime - StartTime) & " seconds"

SwitchOff (False) 'turn these features back on

End Sub

Sub loopingweek()

Dim wsSource, wsDest As Worksheet
Dim StartTime As Single
Dim EndTime As Single

SwitchOff (True) 'turn off these features
'Get the start time
StartTime = Timer

'Call openwb

'set variable
Set wsSource = Workbooks("Test Auto.xlsb").Worksheets("Booking_SGN")
wsSource.Activate
Set wsDest = ActiveWorkbook.Worksheets("fte_data")
    
Call process
Call ftecounting
    
'Call closewb
wsDest.Activate
    
'Get the end time
EndTime = Timer
'Print the processing time
Debug.Print "Processing time: " & (EndTime - StartTime) & " seconds"

SwitchOff (False) 'turn these features back on

End Sub

Sub updatehistorical()
Dim ws As Worksheet

SwitchOff (True) 'turn off these features

'set variables
Set ws = Workbooks("Test Auto.xlsb").Worksheets("Booking_SGN")

    ws.Range("F3") = "Historical"
    ws.Range("D3") = "=YEAR(TODAY())-1"
    ActiveWorkbook.Save

SwitchOff (False) 'turn these features back on

End Sub

Sub updateforecast()
Dim ws As Worksheet

SwitchOff (True) 'turn off these features

'set variables
Set ws = Workbooks("Test Auto.xlsb").Worksheets("Booking_SGN")

    ws.Range("F3") = "Forecast"
     ActiveWorkbook.Save

SwitchOff (False) 'turn these features back on

End Sub

Sub updateboh()
Dim ws As Worksheet

SwitchOff (True) 'turn off these features

'set variables
Set ws = Workbooks("Test Auto.xlsb").Worksheets("Booking_SGN")

    ws.Range("F3") = "BOH"
    ActiveWorkbook.Save

SwitchOff (False) 'turn these features back on

End Sub

Sub updateactual()
Dim ws As Worksheet

SwitchOff (True) 'turn off these features

'set variables
Set ws = Workbooks("Test Auto.xlsb").Worksheets("Booking_SGN")

    ws.Range("F3") = "Actual"
    'ws.Range("D3") = "=YEAR(TODAY())"
    'ws.Range("C3") = "=WEEKNUM(TODAY(),21)-1"
    ActiveWorkbook.Save

SwitchOff (False) 'turn these features back on

End Sub

Sub continuelooping()
Dim wsSource, wsDest As Worksheet
Dim sWeek, rWeek, rSource As Range
Dim j As Integer
Dim StartTime As Single
Dim EndTime As Single

SwitchOff (True) 'turn off these features

'Get the start time
StartTime = Timer

'Call openwb

'set variable
Set wsSource = Workbooks("Test Auto.xlsb").Worksheets("Booking_SGN")
wsSource.Activate
Set wsDest = ActiveWorkbook.Worksheets("fte_data")
Set sWeek = wsSource.Range("C5", Range("C5").End(xlToRight))
Set rWeek = wsDest.Cells(Rows.Count, "C").End(xlUp)
Set rSource = wsDest.Cells(Rows.Count, "J").End(xlUp)

'wsSource.Range("C3") = rWeek.Value
'wsSource.Range("F3") = rSource.Value
'ActiveWorkbook.Save

For j = wsSource.Range("C3").Value + 1 To sWeek.Cells.Count
    
        wsSource.Range("C3") = j
        ActiveWorkbook.Save
        'Call process
        'Call ftecounting
        Debug.Print (MsgBox("j"))
    Next j
    
'Call closewb
'wsSource.Range("C3") = "=WEEKNUM(TODAY(),21)"
'wsDest.Activate
    
'Get the end time
EndTime = Timer
'Print the processing time
Debug.Print "Processing time: " & (EndTime - StartTime) & " seconds"

SwitchOff (False) 'turn these features back on

End Sub

Sub checkingvalue()
Dim wsSource, wsDest As Worksheet
Dim sWeek, rWeek, rSource As Range
Dim StartTime, EndTime As Single
Dim answer As Integer
Dim str As String

SwitchOff (True) 'turn off these features

'Get the start time
StartTime = Timer
    
'Call openwb

'set variable
Set wsSource = Workbooks("Test Auto.xlsb").Worksheets("Booking_SGN")
wsSource.Activate
Set wsDest = ActiveWorkbook.Worksheets("fte_data")
Set sWeek = wsSource.Range("C5", Range("C5").End(xlToRight))
Set rWeek = wsDest.Cells(Rows.Count, "C").End(xlUp)
Set rSource = wsDest.Cells(Rows.Count, "J").End(xlUp)
str = "You will continue the macro for the Forecast at " & Chr(13) & Chr(10) & Chr(148) & "Week: " & wsSource.Range("C3").Value & Chr(148) & " with " & Chr(148) & "Forecast Data" & Chr(148)

If rWeek.Value = wsSource.Range("C3").Value Then
    answer = MsgBox(str, vbQuestion + vbYesNo, "Continue to update Forecast data")
        If answer = vbYes Then
            MsgBox "Yes"
        Else
            Exit Sub
        End If
Else
    MsgBox "Insufficient Data." & vbNewLine & "Please recheck your 'fte_data' and 'Booking_SGN'.", vbCritical + vbOKOnly, "Continue to update Forecast data"
End If

'Get the end time
EndTime = Timer
'Print the processing time
Debug.Print "Processing time: " & (EndTime - StartTime) & " seconds"

SwitchOff (False) 'turn these features back on

End Sub
