Sub recordsetSql()
    Dim SQL, BEATWEEN, SQLsource, DBFullName, connect As String

    DBFullName = ThisWorkbook.FullName
    Set g_CN = CreateObject("ADODB.Connection")
    connect = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" & DBFullName _
    & ";Extended Properties=""Excel 12.0;HDR=Yes;IMEX=1"";"
    g_CN.Open connect
    
    Set g_RS = CreateObject("ADODB.recordset")

    SQL = "SELECT * FROM [worksheetName$A1Z1000]"

    .Open SQL, g_CN, adOpenStatic
        
        With Data_sht
            .Range("A2:DH300000").Value = ""
            .Cells(2, 1).Offset(0, 0).CopyFromRecordset g_RS
        End With
        .Close
    End With
  
    
    
    Set g_RS = Nothing
    g_CN.Close
    Set g_CN = Nothing


End Sub




