Option Explicit

Sub Prepare_tsbleSrc()
    Call disableAlerts
    Call showUfUpdate
    
    Call Create_main_table
    
    Calculate
    
    uf_update.Hide
    MsgBox ("success!")
    Call enableAlerts
End Sub

'// ***************
Sub RefreshConn()
Dim link As String
Dim Arr1()
Dim rng As Range
Dim ClassColumn As Variant
Dim headerRow As Integer

    link = ""
    ThisWorkbook.Connections(1).Refresh
    ThisWorkbook.Connections(2).Refresh
    ThisWorkbook.Connections(3).Refresh
    ThisWorkbook.Connections(4).Refresh
    Application.Wait (Now + TimeValue("00:00:07"))

End Sub
