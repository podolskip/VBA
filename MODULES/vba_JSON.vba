Option Explicit

Sub convertJsonString()
Dim jsonArr, scriptControl, Item As Object

Set scriptControl = CreateObject("MSScriptControl.ScriptControl")
scriptControl.Language = "JScript"

Set jsonArr = scriptControl.Eval("(" + ThisWorkbook.Sheets("sheetName").Cells(1, 1).Value + ")")


End Sub

