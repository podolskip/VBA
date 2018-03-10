Option Explicit

Sub Create_Dashboard_PPT()
Dim x, y, z, w, a, b, c, k As Long
Dim TEMPLATE_NAME, TempString As String
Dim slide_number As Integer
Dim clipboard As MSForms.DataObject
Dim PPTFile As PowerPoint.Presentation
Dim PPApp As Object
Dim ReportDate As Date
Dim DivisionName As String
Dim Header As Range
Dim ws As Worksheet
Dim wb As Workbook
Dim SlideNr As Scripting.Dictionary
Dim TopH
Dim TopAll
Dim LefTall
Dim WidthAll
Dim Arr1() As Variant
Dim Coll1 As New Collection
Dim FilePath, LastRow As Variant
Dim lRow As Variant
Dim lCol As Variant
Dim UpR, DownR, TopHCurrent, ShapeRowsCount, RowsHeinghtCount As Variant


Call disableAlerts

On Error GoTo ErrorHandler:
Set clipboard = New MSForms.DataObject
Set wb = ThisWorkbook
MsgBox "Please choose Dashboards_Template.pptx"
'TEMPLATE_NAME = Application.GetOpenFilename("PPT Files (*.pptx), *.pptx", , "Please choose Dashboards_Template.pptx", , False)  ' ActiveWorkbook.Path & "\Dashboards_Template.pptx" 'Application.GetOpenFilename("PPT Files (*.pptx), *.pptx")

With Application.FileDialog(msoFileDialogFilePicker)
    .title = "Select Test File"
    .Filters.Clear 'Add "PPT Files (*.pptx), *.pptx"
    .AllowMultiSelect = False
    .InitialFileName = "*Dashboards_Template*.*"
     TEMPLATE_NAME = .Show
     
    If TEMPLATE_NAME <> 0 Then
         TEMPLATE_NAME = .SelectedItems.Item(1)
    Else
        Exit Sub
    End If
End With

ReportDate = VBA.Format(Now(), "DD/MM/YYYY")

If TEMPLATE_NAME = False Then  ' Len(Dir(TEMPLATE_NAME)) = 0 Then
    Application.ScreenUpdating = True
    MsgBox "Macro did not found template PPT file in: " & vbNewLine & ActiveWorkbook.Path & vbNewLine & vbNewLine & "Please make sure template PPT file is in same folder as this file, and then run macro again. Thank you!", vbCritical, "Macro stopped!"
    Exit Sub
End If


Set PPApp = CreateObject("PowerPoint.Application")
PPApp.WindowState = 2
Set PPTFile = PPApp.Presentations.Open(TEMPLATE_NAME)
'Slide master set
PPApp.ActivePresentation.Designs(1).SlideMaster.Shapes("Footer2") _
    .TextFrame.TextRange.Text = Application.Text(ReportDate, "[$-409]MMMM") & ", " & Application.Text(ReportDate, "[$-409]YYYY")
'PPApp.ActivePresentation.Designs(1).SlideMaster.Shapes("Footer") _
    .TextFrame.TextRange.Text = "AEI PMO | " & Application.UserName

'#1
slide_number = 1
PPApp.ActiveWindow.View.GotoSlide slide_number
'PPTFile.Slides(slide_number).Shapes("Subtitle 2").TextFrame.TextRange.text = "Program Milestones Progress status report" & chr(10) & VBA.Format(VBA.Now(), "DD.MM.YYYY")

'#other slides
Application.Wait Now() + TimeValue("00:00:03")

Set ws = ThisWorkbook.Sheets("Dashboard")
LastRow = ws.Cells(Rows.Count, 9).End(xlUp).Row

On Error Resume Next
For c = 5 To LastRow
    If ws.Cells(c, 9) <> "" Then
        Coll1.Add ws.Cells(c, 9), ws.Cells(c, 9)
    End If
Next c
On Error GoTo 0

For z = 1 To Coll1.Count 'Go through all Strategic Levels
    DivisionName = Coll1(z)
    
    Set Header = ws.Range("A4:H5")
    TopH = 85
    LefTall = 19.40654
    TopAll = 115  '103.5
    WidthAll = 680.3751
    Set SlideNr = CropTable2(ws, DivisionName)
    
    Dim SlideToStart, SlideCount As Long
    SlideToStart = 1
    Application.Wait Now() + TimeValue("00:00:01")
    For k = 0 To SlideNr.Count - 1 'Go through all slides in regions
        
        SlideCount = SlideToStart + k + 1
        PPTFile.Slides(SlideToStart).Copy
        PPTFile.Slides.Paste SlideCount
        

''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'copy header
        Header.Copy
        WaitAfterCopying

        PPTFile.Application.ActiveWindow.View.GotoSlide SlideCount
        PPTFile.Slides(SlideCount).Application.Activate
        PPTFile.Application.ActiveWindow.View.PasteSpecial
        Application.CutCopyMode = False

        PPTFile.Slides(SlideCount).Shapes(PPTFile.Slides(SlideCount).Shapes.Count).Top = TopH
        PPTFile.Slides(SlideCount).Shapes(PPTFile.Slides(SlideCount).Shapes.Count).Width = WidthAll
        PPTFile.Slides(SlideCount).Shapes(PPTFile.Slides(SlideCount).Shapes.Count).Left = LefTall

        With PPTFile.Slides(SlideCount).Shapes(PPTFile.Slides(SlideCount).Shapes.Count).Table
            For lRow = 1 To .Rows.Count
                For lCol = 1 To .Columns.Count
                    If .Cell(lRow, lCol).Shape.TextFrame.TextRange.Text <> "" Then
                        .Cell(lRow, lCol).Shape.TextFrame.TextRange.Font.Size = 8
                        .Cell(lRow, lCol).Shape.TextFrame.MarginLeft = 2
                    End If
                Next lCol
            Next lRow
        End With
        
        PPTFile.Slides(SlideCount).Shapes(PPTFile.Slides(SlideCount).Shapes.Count).Table.Rows(1).Height = 10.96685
        PPTFile.Slides(SlideCount).Shapes(PPTFile.Slides(SlideCount).Shapes.Count).Table.Rows(2).Height = 35.48102
        

        TopHCurrent = TopH + PPTFile.Slides(SlideCount).Shapes(PPTFile.Slides(SlideCount).Shapes.Count).Height
        
'copy each project seperately
        For UpR = WorksheetFunction.Min(SlideNr.Items()(k), SlideNr.Keys()(k)) To WorksheetFunction.Max(SlideNr.Items()(k), SlideNr.Keys()(k)) Step 5
            If UpR <> WorksheetFunction.Max(SlideNr.Items()(k), SlideNr.Keys()(k)) Then
                ws.Range("A" & UpR & ":H" & (UpR + 4)).Copy
                WaitAfterCopying
'Stop
                PPTFile.Application.ActiveWindow.View.GotoSlide SlideCount
                PPTFile.Slides(SlideCount).Application.Activate
                PPTFile.Application.ActiveWindow.View.PasteSpecial
                Application.CutCopyMode = False
                
                PPTFile.Slides(SlideCount).Shapes(PPTFile.Slides(SlideCount).Shapes.Count).Top = TopHCurrent
                PPTFile.Slides(SlideCount).Shapes(PPTFile.Slides(SlideCount).Shapes.Count).Width = WidthAll
                PPTFile.Slides(SlideCount).Shapes(PPTFile.Slides(SlideCount).Shapes.Count).Left = LefTall

                With PPTFile.Slides(SlideCount).Shapes(PPTFile.Slides(SlideCount).Shapes.Count).Table
                    For lRow = 1 To .Rows.Count
                        For lCol = 1 To .Columns.Count
                            If .Cell(lRow, lCol).Shape.TextFrame.TextRange.Text <> "" Then
                                .Cell(lRow, lCol).Shape.TextFrame.TextRange.Font.Size = 8
                                .Cell(lRow, lCol).Shape.TextFrame.MarginLeft = 2
                            End If
                        Next lCol
                    Next lRow
                End With

                TopHCurrent = TopHCurrent + PPTFile.Slides(SlideCount).Shapes(PPTFile.Slides(SlideCount).Shapes.Count).Height
                            
            Else
                          
            End If
            
            'Adjust height of RAG rows. works like Distribute Rows function in PPT
            With PPTFile.Slides(SlideCount).Shapes(PPTFile.Slides(SlideCount).Shapes.Count).Table
                
                If .Cell(1, 1).Shape.TextFrame.TextRange.Text <> "" Then
                    ShapeRowsCount = .Rows.Count '- 1
                Else
                    ShapeRowsCount = .Rows.Count
                End If
                
                RowsHeinghtCount = .Rows(ShapeRowsCount - 1).Height + .Rows(ShapeRowsCount - 2).Height + .Rows(ShapeRowsCount - 3).Height
                .Rows(ShapeRowsCount - 1).Height = RowsHeinghtCount / 3
                .Rows(ShapeRowsCount - 2).Height = RowsHeinghtCount / 3
                .Rows(ShapeRowsCount - 3).Height = RowsHeinghtCount / 3
            End With

        Next UpR
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

        PPTFile.Slides(SlideCount).Shapes("Text Placeholder 3").TextFrame.TextRange.Text = "A4"
        PPTFile.Slides(SlideCount).Shapes("Title 2").TextFrame.TextRange.Text = DivisionName & " (" & (SlideCount - 1) & "/" & SlideNr.Count & ")"
       
'Stop
        
        ws.Rows(SlideNr.Keys()(k) & ":" & SlideNr.Items()(k)).EntireRow.Hidden = True
    Next k
    
    
Next z
ws.Rows("6:500").EntireRow.Hidden = False
PPTFile.Slides(1).Delete
AppActivate (ThisWorkbook.Name)
'AppActivate Application ' "Microsoft Excel"

ChDir (ThisWorkbook.Path) ' Change directory
again:
FilePath = Application.GetSaveAsFilename("EAM_Project_Portfolio_Dashboard_" & VBA.Format(VBA.Now(), "YYYY_MM_DD") & ".pptx")



If FilePath = False Then
    MsgBox "Select file"
    GoTo again
End If

PPTFile.SaveAs FilePath

MsgBox "Done :)"
'PPTFile.Close
Set PPTFile = Nothing
'PPApp.Quit
On Error GoTo 0
Call enableAlerts
Exit Sub

ErrorHandler:
MsgBox ("Please try again." & Chr(10) & "Folowing obstacles have been met:" & Chr(10) _
        & Err.Number & " | " & Err.Description)

If Not PPApp Is Nothing Then
    PPApp.Quit
    If Not PPTFile Is Nothing Then
        PPTFile.Close
        Set PPApp = Nothing
        Set PPTFile = Nothing
    Else
        Set PPApp = Nothing
        Set PPTFile = Nothing
    End If
Else
    Set PPApp = Nothing
    Set PPTFile = Nothing
End If
Exit Sub

End Sub

Public Function WaitAfterCopying()
    Dim time_now As Variant
    time_now = Now
    While time_now + TimeValue("00:00:01") > Now
        DoEvents
    Wend
End Function


Sub aaa()

CropTable2 ThisWorkbook.Sheets("Dashboard"), "FLDS"

End Sub




Function CropTable(Sht As Worksheet, ByVal Item As Variant) As Scripting.Dictionary ' Microsoft scripting runtime
Dim LastRow, RangeTop, RangeBottom As Long
Dim c As Long
Dim c1 As Long
Dim c2 As Long
Dim H As Long
Dim defaultH As Long
Dim d As New Scripting.Dictionary
Dim CurrentRow As Variant
Dim x, y, z As Variant

LastRow = Sht.Cells(Rows.Count, 9).End(xlUp).Row

'Top Range lookup
For z = 5 To LastRow
    If Sht.Cells(z, 9) = Item Then
    RangeTop = z
    Exit For
    Else
    End If
    
Next z

'Bottom Range lookup
For z = 5 To LastRow
    If Sht.Cells(z, 9) = Item Then
    RangeBottom = z + 4
    
    Else
    End If
    
Next z



defaultH = 300 '383

LastRow = RangeBottom ' ThisWorkbook.Sheets(Sht.Name).Cells(Rows.Count, 4).End(xlUp).Row
c1 = RangeTop
H = 0
    For c = RangeTop To (LastRow)
        H = H + Sht.Rows(c).Height
            If H > defaultH Then
                c2 = c - 1
                d.Add c1, c2
                c1 = c
                H = Sht.Rows(c).Height
            End If
            If c = LastRow Then
                d.Add c, c1
            End If
    Next
    
   Set CropTable = d

End Function



Function CropTable2(Sht As Worksheet, ByVal Item As Variant) As Scripting.Dictionary ' Microsoft scripting runtime
Dim LastRow, RangeTop, RangeBottom As Long
Dim c As Long
Dim c1 As Long
Dim c2 As Long
Dim H, HCurrentRow As Long
Dim defaultH As Long
Dim d As New Scripting.Dictionary
Dim CurrentRow As Variant
Dim x, y, z, q, t, n, KeyTop, ItemBottom As Variant

LastRow = Sht.Cells(Rows.Count, 9).End(xlUp).Row

'Top Range lookup
For z = 5 To LastRow
    If Sht.Cells(z, 9) = Item Then
    RangeTop = z
    Exit For
    Else
    End If
    
Next z

'Bottom Range lookup
For z = 5 To LastRow
    If Sht.Cells(z, 9) = Item Then
    RangeBottom = z + 4
    
    Else
    End If
    
Next z

defaultH = 383  '290
LastRow = RangeBottom ' ThisWorkbook.Sheets(Sht.Name).Cells(Rows.Count, 4).End(xlUp).Row
KeyTop = RangeTop
H = 0
HCurrentRow = 34 '0 34 is the height of two first header rows
    For c = RangeTop To (LastRow + 1) Step 5
        
        y = c
            For q = 1 To 5
                HCurrentRow = HCurrentRow + Sht.Rows(y + q).Height
            Next q
        
            If HCurrentRow > defaultH Then
                
                ItemBottom = c - 1
                d.Add KeyTop, ItemBottom
                KeyTop = c
                c = c - 5
                HCurrentRow = 34 '0
            End If
            
            If HCurrentRow < defaultH And c >= LastRow And c <> KeyTop Then
                d.Add c - 1, KeyTop '- 1
            End If
    Next
    
   Set CropTable2 = d

End Function

