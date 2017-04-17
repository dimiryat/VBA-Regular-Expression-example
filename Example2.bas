Attribute VB_Name = "Module1"
Option Explicit

Sub Macro_entry()

    Dim regEx As New RegExp
    Dim strPattern As String
    Dim tempStr As String
    
    strPattern = "  ([0-9]{8}) ([ \-\.a-zA-Z0-9_]{25}) ([ 0-9]{1,10})([ 0-9]{1,10})[ 0-9]{1,10}[ 0-9]{1,10}"
    regEx.Pattern = strPattern
    regEx.Global = True

    Do While Not IsEmpty(ActiveCell.Value)
        tempStr = ActiveCell.Value
        If regEx.test(tempStr) Then
            ActiveCell.Offset(0, 1).Select
            ActiveCell.Value = CLng(regEx.Replace(tempStr, "$1"))
            ActiveCell.Offset(0, 1).Select
            ActiveCell.Value = regEx.Replace(tempStr, "$2")
            ActiveCell.Offset(0, 1).Select
            ActiveCell.Value = CInt(regEx.Replace(tempStr, "$3"))
            ActiveCell.Offset(0, 1).Select
            ActiveCell.Value = CInt(regEx.Replace(tempStr, "$4"))
        End If
        ActiveCell.Offset(1, -4).Select
    Loop

End Sub
