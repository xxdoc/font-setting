Attribute VB_Name = "Utilities"
Public Function DecToHex(number As Long) As String
    DecToHex = Hex(number)
    DecToHex = Format(DecToHex, "@@@@@@")
    DecToHex = Replace(DecToHex, " ", "0")
End Function
