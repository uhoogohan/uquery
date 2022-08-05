VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "UStringlist"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private str$(), idx%, delim$

Public Sub add(s$)
    ReDim Preserve str(idx + 1)
    str(idx) = s
    idx = idx + 1
End Sub

Public Function GetStr() As String
    Dim i%, s

    For i = 0 To idx - 1
        s = s & IIf(i > 0, delim, "") & str(i)
    Next

    GetStr = s
End Function

Public Property Let delimiter(ByVal val$)
    delim = val
End Property

Public Function items$(i%)
    items = str(i)
End Function

Public Function length%()
    length = idx
End Function

Public Sub clear()
    idx = 0
    Erase str()
End Sub

Private Sub Class_Initialize()
    clear
    delim = " "
End Sub
