VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "UConnection"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private con As ADODB.Connection

Public Sub OpenConnection()
    Set con = New ADODB.Connection

    With con
        .Provider = "Microsoft.ACE.OLEDB.12.0;"
        .Properties("Extended Properties") = "Excel 12.0 xml;"
        .Open ThisWorkbook.FullName
    End With
End Sub

Public Sub CloseConnection()
    If con.State <> adStateClosed Then
        con.Close
    End If

    Set con = Nothing
End Sub

Public Function GetCon() As ADODB.Connection
    Set GetCon = con
End Function

Private Sub Class_Initialize()
    OpenConnection
End Sub

Private Sub Class_Terminate()
    CloseConnection
End Sub
