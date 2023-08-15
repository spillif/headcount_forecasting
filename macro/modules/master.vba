Option Explicit

Public wbg1, wbg2, wbg3, wbg4, wbg5 As Workbook

'Set up Master Data
Global Const TITLE_POSITION = 5
Public DataRange() As DataRangeVisibility
Public bInit As Boolean
Public oBookingSgn As New cWorksheetEx
Public ows1 As New cWorksheetEx


Sub Init()
Dim i As Long, l As Long
oBookingSgn.Init ThisWorkbook.Sheets("Booking_SGN"), TITLE_POSITION
On Error GoTo ErrHandle
Excel.Application.ScreenUpdating = False
'Load data
ows1.Init ThisWorkbook.Sheets("Range Visibility"), 1
If ows1.Worksheet.AutoFilterMode = True Then ows1.Worksheet.AutoFilterMode = False
l = ows1.FindLastRecord()
ReDim DataRange(l - ows1.TitlePosition)
For i = 1 To l - ows1.TitlePosition
    DataRange(i - 1).Range = ows1.Cells(i, "Range").Value
    DataRange(i - 1).Description = ows1.Cells(i, "Description").Value
    DataRange(i - 1).Warning = ows1.Cells(i, "Note").Value
    DataRange(i - 1).ExecutionString = ows1.Cells(i, "Excecution String").Value
    'Excecution String
    'Debug.Print i & " " & DataRange(i - 1).Description
Next i
bInit = True
On Error GoTo 0
Exit Sub
ErrHandle:
Excel.Application.ScreenUpdating = True
On Error GoTo 0
Resume

End Sub


Sub setpulicproperty()
Application.ScreenUpdating = False

    Set wbg1 = Workbooks.Open("https://apll.sharepoint.com/sites/vietnam/VNM_SQ/Projects%20(Migrate%20from%20Local%20Drive)/2023/Workload%20Allocation/SGN/G1/Time%20Motion%20-%20G1.xlsx?web=1", 3)
    Set wbg2 = Workbooks.Open("https://apll.sharepoint.com/sites/vietnam/VNM_SQ/Projects%20(Migrate%20from%20Local%20Drive)/2023/Workload%20Allocation/SGN/G2/Time%20Motion%20-%20G2.xlsx?web=1", 3)
    Set wbg3 = Workbooks.Open("https://apll.sharepoint.com/sites/vietnam/VNM_SQ/Projects%20(Migrate%20from%20Local%20Drive)/2023/Workload%20Allocation/SGN/G3/Time%20Motion%20-%20G3.xlsx?web=1", 3)
    Set wbg4 = Workbooks.Open("https://apll.sharepoint.com/sites/vietnam/VNM_SQ/Projects%20(Migrate%20from%20Local%20Drive)/2023/Workload%20Allocation/SGN/G4/Time%20Motion%20-%20G4.xlsx?web=1", 3)
    Set wbg5 = Workbooks.Open("https://apll.sharepoint.com/sites/vietnam/VNM_SQ/Projects%20(Migrate%20from%20Local%20Drive)/2023/Workload%20Allocation/SGN/G5/Time%20Motion%20-%20G5.xlsx?web=1", 3)

Application.ScreenUpdating = True
End Sub

