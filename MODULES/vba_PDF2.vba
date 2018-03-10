Option Explicit


Sub Send_Email_With_PDF()
Dim nameOfFile As String
Dim Item As String
    
    Item = ThisWorkbook.ActiveSheet.Name
    'define pdf file that is going to me attached to email
    nameOfFile = "C:\temp\External_Law_Provider_" & Item & "_" & VBA.Format(VBA.Now(), "DD_MM_YYYY") & ".pdf"
    
    Call Send_Email_From_Excel(Subcject_Create, CSS_Create, Body_Create, nameOfFile)
    
    Kill nameOfFile
    
End Sub

'1)
Public Function Subcject_Create()
Dim subject As String
    
    subject = "External Law Providers Survey Result for - " & ThisWorkbook.ActiveSheet.Name
    
    Subcject_Create = subject

End Function

'2)
Public Function CSS_Create()
Dim CSS As String
    
    CSS = "<head><style>" & _
            ".email {font-size: 20px;} " & _
            "</style></head>"
    
    CSS_Create = CSS

End Function

'3)
Public Function Body_Create()
Dim body As String
    
    body = "<div class=""email"" >" & _
            "<br>" & _
            "Dear all, <br> " & _
            "Please find attached the report containing the External Law Providers survey results for: <br>" & _
            "<h2>" & ThisWorkbook.ActiveSheet.Name & "</h2>" & _
            "<br>" & _
            "Kind Regards," & _
            "</div>"
    
    Body_Create = body

End Function

'4)
Sub Send_Email_From_Excel(ByVal ToSubj As String, ByVal CSS As String, ByVal ToMsg As String, ByVal nameOfFile As String)
Dim OutApp As Object
Dim oExcelEmailApp As Object

    'Click Tools -> References -> Microsoft Outlook nn.n Object Library 14.0
    Set OutApp = CreateObject("Outlook.Application")
    Set oExcelEmailApp = OutApp.CreateItem(0)

    'VBA Create email
    With oExcelEmailApp
        .Display
        .To = ""
        .CC = Environ("username")
        .BCC = ""
        .subject = ToSubj
        .htmlbody = ("<html>" & CSS & ToMsg & "</html>") & .htmlbody 'ToMsg & .htmlbody
        .Attachments.Add (nameOfFile)
         'to send automatic mail from excel instead of .display use .send
        '.send
    End With
    
    Set OutApp = Nothing
    Set oExcelEmailApp = Nothing
End Sub



