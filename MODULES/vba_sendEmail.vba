Option Explicit



Private Declare Function URLDownloadToFile Lib "urlmon" _
Alias "URLDownloadToFileA" (ByVal pCaller As Long, _
ByVal szURL As String, ByVal szFileName As String, _
ByVal dwReserved As Long, ByVal lpfnCB As Long) As Long

Public Const sht_1 As Variant = "Control"
Public Const sht_2 As Variant = "ControlPanel"
Public Const sht_3 As Variant = "Current NEWSLETTER"
Public Const sht_4 As Variant = "Newsletter_Shp_List"


Sub send_Email_To_Authors(ByRef authorColl As Collection, ByVal recordCount As Variant, ByVal newsletterWeek As String)
Dim toTO As String
    'create authors-senders To list
    toTO = authorColl.Item(1)
    For g_z = 2 To authorColl.Count
        toTO = toTO & "; " & authorColl.Item(g_z)
    Next g_z
    
    Call Send_Email_From_Excel(toTO, Subcject_Create(newsletterWeek), CSS_Create, Body_Create, "", "")

End Sub

Sub send_Newsletter_Preview(ByRef sectionColl As Collection, ByVal recordCount As Variant, ByVal newsletterWeek As String)
Dim toTO As String


    Call Send_Email_From_Excel(toTO, Subcject_Create(newsletterWeek), CSS_Create, Preview_Body_Create(sectionColl), "", "")

End Sub

Sub Send_Email_With_IMGs(ByRef imgArr(), ByVal recordCount As Variant, ByVal newsletterWeek As String)
Dim nameOfFile, strDate, tempPath, dateTday, exlNewsName, exlNewsPath As String
    
    dateTday = Int(Rnd() * 100000) 'Int(CDbl(Date))
    tempPath = "C:\temp\news" & dateTday '& "\"
    strDate = Format(Now, " dd-mmm-yy h-mm-ss")
    exlNewsName = "Newsletter" & "_" & newsletterWeek
    
    '1
    Call check_if_folder_exists(tempPath)
    
    '2
    Call save_imgs(imgArr, recordCount, tempPath)
    
    '3
    Call Zip_All_Files_in_Folder_Browse(tempPath, strDate)
    
    '4
    exlNewsPath = sb_Copy_Save_Worksheet_As_Workbook(tempPath, exlNewsName)
    nameOfFile = tempPath & "\IMG " & strDate & ".zip"
    
    Call Send_Email_From_Excel("web.services2@credit-suisse.com", Subcject_Create(newsletterWeek), CSS_Create, Body_Create, nameOfFile, exlNewsPath)
    
    removeFolder (tempPath & "\img") 'removeFolder in aAPI_GLOBALS
    removeFolder (tempPath) 'removeFolder in aAPI_GLOBALS
    
    
End Sub


Sub check_if_folder_exists(ByVal fPath As String)

If Len(Dir(fPath, vbDirectory)) = 0 Then
    MkDir fPath
    MkDir fPath & "\img"
Else
    removeFolder (fPath & "\img") 'removeFolder in aAPI_GLOBALS
    removeFolder (fPath) 'removeFolder in aAPI_GLOBALS
    
    MkDir fPath
    MkDir fPath & "\img"
End If

End Sub

'(3)
Sub save_imgs(ByRef imgArr(), ByVal recordCount As Variant, ByVal tempPath As String)
Dim res As Variant


For g_z = 1 To recordCount
    res = URLDownloadToFile(0, imgArr(g_z, 2), tempPath & "\img\" & imgArr(g_z, 1) & ".jpg", 0, 0)
    
Next g_z

End Sub

'(4)
Function sb_Copy_Save_Worksheet_As_Workbook(ByVal tempPath As String, ByVal exlNewsName As String)
Dim wb As Workbook
Dim finalPath As String
    
    Set wb = Workbooks.Add
    ThisWorkbook.Sheets(sht_3).Copy Before:=wb.Sheets(1)
    finalPath = tempPath & "\" & VBA.Replace(exlNewsName, "/", "") & ".xlsx" '"C:\temp\test1.xlsx"
    wb.SaveAs finalPath
    Workbooks(VBA.Replace(exlNewsName, "/", "") & ".xlsx").Close SaveChanges:=False
    
    sb_Copy_Save_Worksheet_As_Workbook = finalPath
    
End Function

'1)
Public Function Subcject_Create(ByVal eTitle As String)
Dim subject As String
    
    subject = "CS Poland Newsletter - " & eTitle
    Subcject_Create = subject

End Function

'2)
Public Function CSS_Create()
Dim CSS As String
    
    CSS = "<head><style>" & _
            ".email {font-size: 20px;} " & _
            ".imgResize {height: 155px !important; width: 310px !important;}" & _
            "</style></head>"
    
    CSS_Create = CSS

End Function

'3.1)
Public Function Body_Create()
Dim body, title As String
Dim Body2 As String
Dim currentOption As Variant
    
    currentOption = ThisWorkbook.Sheets("ControlPanel").Cells(4, 2)
    
    If currentOption = 2 Then
        body = "<div class=""email"" >" & _
                "Dear all, <br>" & _
                "<br>" & _
                "please find attached the input for the next issue of the CS Poland Newsletter.<br>" & _
                "<br>" & _
                "Kind regards, <br>" & _
                "Newsletter Team" & _
                "</div>" & _
                "<br>"
    Else
        body = "<div class=""email"" >" & _
                "Dear all, <br>" & _
                "<br>" & _
                "please find attached the preview of next issue of the CS Poland Newsletter.<br>" & _
                "<br>" & _
                "Please provide your feedback by: <br>" & _
                "<br>" & _
                "Kind regards, <br>" & _
                "Newsletter Team" & _
                "</div>" & _
                "<br>"
    End If
    
    Body_Create = body

End Function

'3.2)
Public Function Preview_Body_Create(ByVal sectionColl As Collection)
Dim body, title, startDIV, endDIV As String
Dim Body2 As String
Dim titleTmp, bodyTmp As String
    
    startDIV = "<div align=center>" & _
                "<table class=MsoNormalTable border=0 cellspacing=0 cellpadding=0 width=640 style='width:480.0pt;mso-cellspacing:0in;background:white;mso-yfti-tbllook: 1184;mso-padding-alt:0in 0in 0in 0in'>" & _
                "<tr style='mso-yfti-irow:0;mso-yfti-firstrow:yes;mso-yfti-lastrow:yes'><td width=""100%"" style='width:100.0%;padding:0in 0in 0in 0in'>" & _
                "<table class=MsoNormalTable border=0 cellspacing=0 cellpadding=0 width=""100%"" style='width:100.0%;mso-cellspacing:0in;mso-yfti-tbllook: 1184;mso-padding-alt:0in 0in 0in 0in'>"
                
    endDIV = "</table></td></tr></table></div>"
    
    '{{SECTION}}
    title = "<br><br><table class=MsoNormalTable border=0 cellspacing=0 cellpadding=0 width=""100%"" style='width:100.0%;mso-cellspacing:0in;mso-yfti-tbllook: 1184;mso-padding-alt:0in 0in 0in 0in'>" & _
            "<tr style='mso-yfti-irow:0;mso-yfti-firstrow:yes;mso-yfti-lastrow: yes'><td style='padding:0in 0in 11.25pt 0in'><p class=MsoNormal style='line-height:18.0pt'><span style='font-size: 13.5pt;font-family:""Arial"",""sans-serif"";mso-fareast-font-family:""Times New Roman""; color:#6D6E71'>{{SECTION}}<o:p></o:p></span></p>" & _
            "</td></tr></table>" & _
            "<br>"
    
    
    '{{TITLE}}
    '{{CONTENT}}
    '{{LINK1NAME}}
    '{{LINK1}}
    '{{LINK2NAME}}
    '{{LINK2}}
    '{{LINK3NAME}}
    '{{LINK3}}
    '{{IMG}}
    
    Body2 = "<table class=MsoNormalTable border=0 cellspacing=0 cellpadding=0 width=""100%"" style='width:100.0%;mso-cellspacing:0in;mso-yfti-tbllook: 1184;mso-padding-alt:0in 0in 0in 0in'><tr style='mso-yfti-irow:0;mso-yfti-firstrow:yes;mso-yfti-lastrow: yes '>" & _
        "<td width=310 valign=top style='width:232.5pt;padding:0in 0in 0in 0in'><table class=MsoNormalTable border=0 cellspacing=0 cellpadding=0 width=""100%"" style='width:100.0%;mso-cellspacing:0in;mso-yfti-tbllook: 1184;mso-padding-alt:0in 0in 0in 0in'>" & _
                "<tr style='mso-yfti-irow:0;mso-yfti-firstrow:yes'><td style='border:none;border-top:solid #004C97 2.25pt;padding: 0in 0in 0in 0in'><p class=MsoNormal style='line-height:11.25pt'><span style='font-size:1.0pt;font-family:""Arial"",""sans-serif"";mso-fareast-font-family: ""Times New Roman""'>&nbsp;<o:p></o:p></span></p>" & _
                    "</td><td width=1 style='width:.75pt;padding:0in 0in 0in 0in;max-width: 0px; max-height: 0px'><p class=MsoNormal style='mso-line-height-alt:0pt'><span style='mso-fareast-font-family:""Times New Roman"";display:none; mso -Hide: all '>" & _
                    "<img border=0 width=1 id=""_x0000_i1033"" src=""{{IMG}}"" style='max-width: 0px;max-height: 0px;display:none; width:310px !important; height:155px !important;' alt="" Spotlight"" class=""mobile-only content-s imgResize""><o:p></o:p></span></p>" & _
                    "</td></tr><tr style='mso-yfti-irow:1'><td style='padding:0in 0in 6.0pt 0in'><p class=MsoNormal style='line-height:22.5pt'><b><span style='font-size:18.0pt;font-family:""Arial"",""sans-serif""; mso-fareast-font-family:""Times New Roman"";color:black'>{{TITLE}}<o:p></o:p></span></b></p>" & _
                    "</td><td style='padding:0in 0in 0in 0in'></td></tr><tr style='mso-yfti-irow:2'><td style='padding:0in 0in 7.5pt 0in'><p class=MsoNormal style='line-height:15.75pt'><span style='font-size:11.5pt;font-family:""Arial"",""sans-serif""; mso-fareast-font-family:""Times New Roman"";color:black'>{{CONTENT}}<o:p></o:p></span></p>" & _
                    "</td><td style='padding:0in 0in 0in 0in'></td>" & _
                    "</tr><tr style='mso-yfti-irow:3;mso-yfti-lastrow:yes'><td style='padding:0in 0in 7.5pt 0in'><table class=MsoNormalTable border=0 cellspacing=0 cellpadding=0 style='mso-cellspacing:0in;mso-yfti-tbllook:1184;mso-padding-alt: 0in 0in 0in 0in'>" & _
                            "<tr style='mso-yfti-irow:0;mso-yfti-firstrow:yes;mso-yfti-lastrow: yes'><td style='padding:0in 0in 0in 0in'><p class=MsoNormal style='line-height:15.75pt'><span style='font-size:11.5pt;font-family:""Arial"",""sans-serif""; mso-fareast-font-family:""Times New Roman"";color:#004C97'>" & _
                            "<a href=""{{LINK1}}"" target=""_blank"">{{LINK1NAME}}</a><o:p></o:p></span></p>" & _
                                "</td></tr></table></td><td style='padding:0in 0in 0in 0in'></td></tr>" & _
                                "<tr style='mso-yfti-irow:4'><td style='padding:0in 0in 7.5pt 0in'><table class=MsoNormalTable border=0 cellspacing=0 cellpadding=0 style='mso-cellspacing:0in;mso-yfti-tbllook:1184;mso-padding-alt: 0in 0in 0in 0in'>" & _
                                    "<tr style='mso-yfti-irow:0;mso-yfti-firstrow:yes;mso-yfti-lastrow: yes '><td style='padding:0in 0in 0in 0in'><p class=MsoNormal style='line-height:15.75pt'><span style='font-size:11.5pt;font-family:""Arial"",""sans-serif""; mso-fareast-font-family:""Times New Roman"";color:#004C97'>" & _
                                    "<a href = ""{{LINK2}}"">{{LINK2NAME}}</a><o:p></o:p></span></p></td><td valign=top style='padding:0in 0in 0in 0in'><p class=MsoNormal style='line-height:13.5pt'>" & _
                                    "<span style='font-size:13.5pt;font-family:""Arial"",""sans-serif""; mso-fareast-font-family:""Times New Roman""'><o:p></o:p></span></p>" & _
                                    "</td></tr></table></td><td style='padding:0in 0in 0in 0in'></td></tr>" & _
                            "<tr style='mso-yfti-irow:5;mso-yfti-lastrow:yes'><td style='padding:0in 0in 7.5pt 0in'><table class=MsoNormalTable border=0 cellspacing=0 cellpadding=0 style='mso-cellspacing:0in;mso-yfti-tbllook:1184;mso-padding-alt: 0in 0in 0in 0in'>" & _
                                    "<tr style='mso-yfti-irow:0;mso-yfti-firstrow:yes;mso-yfti-lastrow: yes '><td style='padding:0in 0in 0in 0in'><p class=MsoNormal style='line-height:15.75pt'><span style='font-size:11.5pt;font-family:""Arial"",""sans-serif""; mso-fareast-font-family:""Times New Roman"";color:#004C97'>" & _
                                    "<a href = ""{{LINK3}}"" target=""_blank"">{{LINK3NAME}}</a><o:p></o:p></span></p></td><td valign=top style='padding:0in 0in 0in 0in'><p class=MsoNormal style='line-height:13.5pt'>" & _
                                    "<span style='font-size:13.5pt;font-family:""Arial"",""sans-serif""; mso-fareast-font-family:""Times New Roman""'><o:p></o:p></span></p>" & _
                                    "</td></tr></table></td><td style='padding:0in 0in 0in 0in'></td></tr>" & _
                                "</table></td><td width=20 style='width:15.0pt;padding:0in 0in 0in 0in'></td><td width=310 valign=top style='width:232.5pt;padding:0in 0in 0in 0in'>" & _
            "<p class=MsoNormal><span style='font-size:1.0pt;font-family:""Arial"",""sans-serif""; mso-fareast-font-family:""Times New Roman""'><img border=0 width=310 height=155 id=""_x0000_i1035"" src=""{{IMG}}"" style='display:block;' alt="" Spotlight"" class=""imgResize""><img border=0 height=15 id=""_x0000_i1036"" src=""https://cdn.credit-suisse.com/media/applications/inxmail/products-services/spacer.gif"" style='display:block'><o:p></o:p></span></p>" & _
        "</td></tr></table>"
            
    Dim SQL, BEATWEEN, SQLsource, DBFullName, connect, SQLColumns As String
    Dim Data_sht As Worksheet
    
    ThisWorkbook.SaveAs
    DBFullName = ThisWorkbook.FullName
    DBFullName = Parse_Resource(ThisWorkbook.FullName)
    Set g_CN = CreateObject("ADODB.Connection")
    connect = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" & DBFullName _
    & ";Extended Properties=""Excel 12.0;HDR=Yes;IMEX=1"";"
    g_CN.Open connect
    
    For g_z = 1 To sectionColl.Count
        Set g_RS = CreateObject("ADODB.recordset")
        SQLsource = "[Current NEWSLETTER$A1:Q2000]"
        SQL = "SELECT * FROM " & SQLsource & "WHERE [Section] = '" & sectionColl.Item(g_z) & "' ORDER BY [Section] ASC"
        
        With g_RS
            .Open SQL, g_CN, adOpenStatic
            .MoveFirst
            
            titleTmp = title
            titleTmp = VBA.Replace(titleTmp, "{{SECTION}}", VBA.Replace(sectionColl.Item(g_z), Left(sectionColl.Item(g_z), 2), ""))
            body = body + titleTmp
            For g_x = 0 To (g_RS.recordCount - 1)
                bodyTmp = Body2
                bodyTmp = VBA.Replace(bodyTmp, "{{TITLE}}", g_RS.Fields(6).Value)
                bodyTmp = VBA.Replace(bodyTmp, "{{CONTENT}}", g_RS.Fields(8).Value)
                
                If IsNull(g_RS.Fields(11).Value) Then
                    bodyTmp = VBA.Replace(bodyTmp, "{{LINK1NAME}}", "")
                    bodyTmp = VBA.Replace(bodyTmp, "{{LINK1}}", "")
                Else
                    bodyTmp = VBA.Replace(bodyTmp, "{{LINK1NAME}}", g_RS.Fields(10).Value)
                    bodyTmp = VBA.Replace(bodyTmp, "{{LINK1}}", g_RS.Fields(11).Value)
                End If
                
                If IsNull(g_RS.Fields(13).Value) Then
                    bodyTmp = VBA.Replace(bodyTmp, "{{LINK2NAME}}", "")
                    bodyTmp = VBA.Replace(bodyTmp, "{{LINK2}}", "")
                Else
                    bodyTmp = VBA.Replace(bodyTmp, "{{LINK2NAME}}", g_RS.Fields(12).Value)
                    bodyTmp = VBA.Replace(bodyTmp, "{{LINK2}}", g_RS.Fields(13).Value)
                End If
                
                If IsNull(g_RS.Fields(15).Value) Then
                    bodyTmp = VBA.Replace(bodyTmp, "{{LINK3NAME}}", "")
                    bodyTmp = VBA.Replace(bodyTmp, "{{LINK3}}", "")
                Else
                    bodyTmp = VBA.Replace(bodyTmp, "{{LINK3NAME}}", g_RS.Fields(14).Value)
                    bodyTmp = VBA.Replace(bodyTmp, "{{LINK3}}", g_RS.Fields(15).Value)
                End If
                
                
                
                'bodyTmp = VBA.Replace(bodyTmp, "{{LINK2NAME}}", g_RS.Fields(12).Value)
                'bodyTmp = VBA.Replace(bodyTmp, "{{LINK2}}", g_RS.Fields(13).Value)
                
                
                'bodyTmp = VBA.Replace(bodyTmp, "{{LINK3NAME}}", g_RS.Fields(14).Value)
                'bodyTmp = VBA.Replace(bodyTmp, "{{LINK3}}", g_RS.Fields(15).Value)
                'FINISH!!!!!!!!!!!!!!!
                
                bodyTmp = VBA.Replace(bodyTmp, "{{IMG}}", g_RS.Fields(16).Value)
                body = body + bodyTmp
                .MoveNext
            Next g_x
            
            
        End With
        g_RS.Close
        Set g_RS = Nothing
    Next g_z
    
    g_CN.Close
    Set g_CN = Nothing
    
    Preview_Body_Create = startDIV + body + endDIV

End Function

'4)
Sub Send_Email_From_Excel(ByVal toTO As String, ByVal ToSubj As String, ByVal CSS As String, ByVal ToMsg As String, ByVal nameOfFile1 As String, ByVal nameOfFile2 As String)
Dim OutApp As Object
Dim oExcelEmailApp As Object

    'Click Tools -> References -> Microsoft Outlook nn.n Object Library 14.0
    Set OutApp = CreateObject("Outlook.Application")
    Set oExcelEmailApp = OutApp.CreateItem(0)

    'VBA Create email
    With oExcelEmailApp
        .Display
        .To = toTO
        .CC = Environ("username")
        .BCC = ""
        .subject = ToSubj
        .htmlbody = ("<html>" & CSS & ToMsg & "</html>") & .htmlbody 'ToMsg & .htmlbody
        If nameOfFile1 <> "" Then
            .Attachments.Add (nameOfFile2)
        End If
        If nameOfFile2 <> "" Then
            .Attachments.Add (nameOfFile1)
        End If
         'to send automatic mail from excel instead of .display use .send
        '.send
    End With
    
    Set OutApp = Nothing
    Set oExcelEmailApp = Nothing
End Sub

