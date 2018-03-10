Sub ConvertToPDFFromWS(ByVal Item As Variant, ByVal openYN As Boolean)
Dim ShArray, SR_sheet, r As String
Dim sh As Worksheet
Dim CurrOrOld As Long
Dim nameOfFile As String
On Error Resume Next

'ThisWorkbook.Sheets("1stPage").Select
ThisWorkbook.Sheets(Item).Select 'False
Call Page_Print_Setup(Item)

On Error GoTo 0

nameOfFile = "C:\temp\External_Law_Provider_" & Item & "_" & VBA.Format(VBA.Now(), "DD_MM_YYYY") & ".pdf"

ActiveSheet.ExportAsFixedFormat Type:=xlTypePDF, _
            Filename:=nameOfFile, _
            Quality:=xlQualityStandard, _
            openafterpublish:=openYN
'ThisWorkbook.Sheets(Item).Select

Exit Sub
ErrorHandler:

Call errHandler
Exit Sub

End Sub

Sub ConvertToPDF(ByVal Item As Variant)
Dim ShArray, SR_sheet, r As String
Dim sh As Worksheet
Dim CurrOrOld As Long

On Error Resume Next
'ThisWorkbook.Sheets("1stPage").Select
ThisWorkbook.Sheets(Item).Select ' False
Call Page_Print_Setup(Item)

On Error GoTo 0

ActiveSheet.ExportAsFixedFormat Type:=xlTypePDF, _
            Filename:="C:\temp\External_Law_Provider_" & Item & "_" & VBA.Format(VBA.Now(), "DD_MM_YYYY") & ".pdf", _
            Quality:=xlQualityStandard, _
            openafterpublish:=True
'ThisWorkbook.Sheets("Control Panel").Select

Exit Sub
ErrorHandler:

Call errHandler
Exit Sub

End Sub


Sub AllToPDF()
Dim ShArray, SR_sheet, r As String
Dim sh As Worksheet
Dim CurrOrOld As Long
Dim ProvColl As Collection
Dim rng As Range
Dim Item As Variant

Set ProvColl = New Collection
Set ws2 = ThisWorkbook.Sheets("ControlPanel")

LastRow = ws2.Cells(Rows.Count, 5).End(xlUp).Row
Set rng = ws2.Range("E3:E" & LastRow)
'Arr = Array("Control Panel", "Template", "Template_v2", "ControlPanel", "External_law_providers")

On Error Resume Next
ThisWorkbook.Sheets("1stPage").Select

For Each Item In Worksheets
    If IsError(Application.Match(Item.Name, rng, 0)) Then
        ThisWorkbook.Sheets(Item.Name).Select False
        
    End If
Next
On Error GoTo 0



ActiveSheet.ExportAsFixedFormat Type:=xlTypePDF, _
            Filename:="C:\temp\External_Law_Provider_" & Item & "_" & VBA.Format(VBA.Now(), "DD_MM_YYYY") & ".pdf", _
            Quality:=xlQualityStandard, _
            openafterpublish:=True
ThisWorkbook.Sheets("Control Panel").Select

Exit Sub
ErrorHandler:

Call errHandler
Exit Sub

End Sub

Sub ClassTest()
Dim shNames As clsShNames

Set shNames = New clsShNames

shNames.class_initialise

Debug.Print shNames.ControlPanel.Name


shNames.ControlPanel.Visible = xlSheetVisible

shNames.ControlPanel.Visible = xlVeryHidden

Set shNames = Nothing

End Sub

Sub Page_Print_Setup(ByVal SheetName As Variant)
Dim Page2Break, Page3Break As Variant
Dim num As Variant
Dim wks As Worksheet

Set wks = ThisWorkbook.Sheets(SheetName)
LastRow = wks.Cells(Rows.Count, 5).End(xlUp).Row + 1



For num = 1 To 2000
    If wks.Cells(num, 18) = "AllEnd" Then
        Page2Break = num
    ElseIf wks.Cells(num, 18) = "End" Then
        Page3Break = num
        Exit For
    End If
Next num

If LastRow < 135 Then
   wks.PageSetup.PrintArea = "$B$2:$Q$" & LastRow
       
    'Setting page breaks
    With wks
        .Rows(104).PageBreak = xlPageBreakManual
        .Rows(Page3Break).PageBreak = xlPageBreakManual
    End With
        
        
Else
    With wks.PageSetup
        '.PrintArea = False
        .PrintArea = "$B$2:$Q$" & LastRow
    '    .Orientation = xlPortrait
    '    .RightHeaderPicture.Filename = "X:\logos\AEI_Logo2.jpg"
    '    .RightHeader = "&G"
    '    .LeftHeaderPicture.Filename = "X:\logos\CS_logo.jpg"
    '    .LeftHeader = "&G"
    '    .CenterHeader = "&20" & wSheet.Cells(4, 13) & Chr(10) & "&20 Status Report" '
    '    .RightFooter = VBA.Format("&Dmmmm d, yyyy", "mmmm d, yyyy") 'VBA.Format(VBA.Now(), "mmmm d, yyyy")
    '    .CenterFooter = "&p/&N"
    '    .LeftFooter = "AEI PMO" 'Application.UserName 'to show address
    '    .PrintTitleRows = "$2:$2" 'repeat at top
    '    .Zoom = False
    '    .FitToPagesWide = 1 'to print in 01 page
    '    .FitToPagesTall = False 'to print in 01 page
    
    End With
    
    'Setting page breaks
    With wks
        .Rows(104).PageBreak = xlPageBreakManual
        .Rows(Page2Break).PageBreak = xlPageBreakManual
        .Rows(Page3Break).PageBreak = xlPageBreakManual
    End With
    
    
End If
    
End Sub


