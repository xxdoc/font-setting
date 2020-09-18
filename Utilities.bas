Attribute VB_Name = "Utilities"
Public Function DecToHex(number As Long) As String
    DecToHex = Hex(number)
    DecToHex = Format(DecToHex, "@@@@@@")
    DecToHex = Replace(DecToHex, " ", "0")
End Function

Public Function JSONParse(ByVal JSONString As String) As Object
    Dim Script As Object
    Set Script = CreateObject("MSScriptControl.ScriptControl")
    Script.Language = "JavaScript"
    Set JSONParse = Script.eval("JSON=" & JSONString & ";JSON.from_title=JSON.from;JSON")
    Set Script = Nothing
End Function
