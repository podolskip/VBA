Option Explicit
Public Const g_DB_Path As String = "\\xyz.net\folder_name"
Public Const g_DB_Name As String = "databaseName.accdb"
Public Const g_ssap As String = "mar"
Public Const g_Developer As String = "Patryk Podolski || "
Public g_lastRow, g_lastColumn As Integer
Public g_a, g_b, g_c, g_x, g_y, g_z, g_i, g_w As Variant
Public g_ws, g_ws2, g_ws3, g_ws4, g_ws5, g_ws6, g_ws7, g_ws8 As Worksheet
Public g_wb, g_wb2, g_wb3, g_wb4, g_wb5, g_wb6, g_wb7, g_wb8 As Workbook
Public g_DBFullName, g_connect, g_SQL, g_SQLsource As Variant
Public g_CN As ADODB.Connection
Public g_RS As ADODB.Recordset
Public g_AppWFun As Object
Public g_Rng, g_Rng2, g_Rng3 As Range

Public Function SafeVlookup(lookupValue, tableArray, colIndex, rangeLookup, errorValue) As Variant
    Dim returnValue As Variant
    Set g_AppWFun = Application.WorksheetFunction
    
    
    On Error Resume Next
    Err.Clear
    returnValue = g_AppWFun.VLookup(lookupValue, tableArray, colIndex, rangeLookup)
    
    If Err <> 0 Then
        returnValue = errorValue
    End If
        
    SafeVlookup = returnValue
    On Error GoTo 0
End Function

Sub saveWithPassword()
    ActiveWorkbook.SaveAs Filename:=ActiveWorkbook.Path & "\fileName.xlsx", Password:=g_ssap

End Sub

Public Function cropUsername()
    
    cropUsername = Left(Application.userName, InStr(1, Application.userName, " ("))

End Function


Sub showUfUpdate()

    With ufUpdate
        .Caption = "Retriving data..."
        .Show vbModeless
        DoEvents
    End With

End Sub

Public Function WaitAfterCopying()
    Dim timeNow As Variant
    timeNow = Now
    While timeNow + TimeValue("00:00:01") > Now
        DoEvents
    Wend
    
End Function

Function isEmptyTF()
    Set ws = ThisWorkbook.Sheets("controlPanel")
    
    If ws2.Cell(3, 9).Value = "" And ws2.Cells(3, 10).Value = "" Then
        isEmptyTF = True
    Else
        isEmptyTF = False
    End If
        
End Function

Sub msgBoxFinish()
    MsgBox prompt:="Finished! :)", Buttons:=vbInformation
End Sub

Sub protectWorkbook()
    ThisWorkbook.Protect Password:=g_ssap
End Sub

Sub unprotectWorkbook()
    ThisWorkbook.Unprotect Password:=g_ssap
End Sub

Sub controlPanelVeryHidden()
    Call unprotectWorkbook
    ThisWorkbook.Sheets("").Visible = xlVeryHidden
    Call protectWorkbook
End Sub

Sub controlPanelVisible()
    Call unprotectWorkbook
    ThisWorkbook.Sheets("").Visible = True
    Call protectWorkbook
End Sub

Function disableAlerts()
    With Application
        .DisplayAlerts = False
        .ScreenUpdating = False
    End With
End Function

Function enableAlerts()
    With Application
        .DisplayAlerts = True
        .ScreenUpdating = True
    End With
End Function

Sub allTabsVeryHidden()
    
        For Each g_ws In Worksheets
            If g_ws.Name <> "Info" Or sh.Name <> "ControlPanel" Then
                g_ws.Visible = xlVeryHidden
            End If
        Next g_ws
End Sub

Sub allTabsVisible()
    
        For Each g_ws In Worksheets
            g_ws.Visible = True
        Next g_ws
End Sub

Sub errHandler()
    If Not Connection Is Nothing Then
        Connection.Close
        If Not RS Is Nothing Then
            RS.Close
            Call setConnRSNothing
        Else
            Call setConnRSNothing
        End If
    Else
        Call setConnRSNothing
    End If
    
    MsgBox ("Following obstacle have been met:" & Chr(10) _
        & Err.Number & " | " & Err.Description)

    On Error GoTo 0
End Sub

Function setConnRSNothing()
    Set Connection = Nothing
    Set RS = Nothing
End Function

Function protectSheet(ByVal shName As Variant)
    ThisWorkbook.Sheets(shName).Protect DrawingObjects:=True, _
                                        Contents:=True, _
                                        Scenarios:=True, _
                                        AllowFormattingColumns:=True, _
                                        AllowFormattingRows:=True, _
                                        AllowFiltering:=True, _
                                        AllowUsingPivotTables:=True, _
                                        Passwird:=g_ssap
                                        
End Function

Function unprotectSheet(ByVal shName As Variant)
    ThisWorkbook.Sheets(shName).Unprotect Password:=g_ssap
End Function

Function protectSheet2()
    ThisWorkbook.Protect Password:=g_ssap, Structure:=True, Windows:=False
End Function

Sub removeFolder(ByVal folderPath As String)
    Dim fso As FileSystemObject
    Set fso = New FileSystemObject
    
    fso.DeleteFile folderPath & "\*.*"
    fso.DeleteFolder folderPath
    
    Set fso = Nothing
End Sub
