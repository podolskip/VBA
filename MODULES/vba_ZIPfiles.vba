Option Explicit

Sub NewZip(sPath)
'Create empty Zip File
'Changed by keepITcool Dec-12-2005
    If Len(Dir(sPath)) > 0 Then Kill sPath
    Open sPath For Output As #1
    Print #1, Chr$(80) & Chr$(75) & Chr$(5) & Chr$(6) & String(18, 0)
    Close #1
End Sub

Sub Zip_All_Files_in_Folder_Browse(ByVal DefPath As String, ByVal strDate As String)
    Dim FileNameZip, FolderName, oFolder
    
    Dim oApp As Object

    'DefPath = Application.DefaultFilePath
    If Right(DefPath, 1) <> "\" Then
        DefPath = DefPath & "\"
    End If

    'strDate = Format(Now, " dd-mmm-yy h-mm-ss")
    FileNameZip = DefPath & "IMG " & strDate & ".zip"

    Set oApp = CreateObject("Shell.Application")

    'Browse to the folder
    'Set oFolder = oApp.BrowseForFolder(0, "Select folder to Zip", 512)
    If Not FolderName <> "" Then 'oFolder Is Nothing Then
        'Create empty Zip File
        NewZip (FileNameZip)

        FolderName = DefPath & "img\" ' oFolder.Self.Path
        If Right(FolderName, 1) <> "\" Then
            FolderName = FolderName & "\"
        End If

        'Copy the files to the compressed folder
        oApp.Namespace(FileNameZip).CopyHere oApp.Namespace(FolderName).Items

        'Keep script waiting until Compressing is done
        On Error Resume Next
        Do Until oApp.Namespace(FileNameZip).Items.Count = _
        oApp.Namespace(FolderName).Items.Count
            Application.Wait (Now + TimeValue("0:00:01"))
        Loop
        On Error GoTo 0

        'MsgBox "You find the zipfile here: " & FileNameZip

    End If
End Sub

