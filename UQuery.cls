VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "UQuery"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private con As UConnection
Private rs As New ADODB.Recordset
Private cmd As New ADODB.Command
Public sql As New UStringlist
Private flds As New UStringlist

Public Sub SetCon(con_ As UConnection)
    Set con = con_
End Sub

' パラメータの追加
Public Sub AddParam(key$, val$) ' 文字列
    Dim prm As ADODB.Parameter
    Set prm = cmd.CreateParameter(key, adVarChar, adParamInput, 255)
    cmd.Parameters.Append prm
    cmd.Parameters(key).Value = val
End Sub

Public Sub AddParamByInt(key$, val%) ' 整数値
    Dim prm As ADODB.Parameter
    Set prm = cmd.CreateParameter(key, adInteger, adParamInput)
    cmd.Parameters.Append prm
    cmd.Parameters(key).Value = val
End Sub

Public Sub AddParamByDouble(key$, val As Double) ' 倍精度浮動小数点
    Dim prm As ADODB.Parameter
    Set prm = cmd.CreateParameter(key, adDouble, adParamInput)
    cmd.Parameters.Append prm
    cmd.Parameters(key).Value = val
End Sub

Public Sub OpenRecordset()
    Set cmd.ActiveConnection = con.GetCon()
    cmd.CommandText = sql.GetStr()
    cmd.ActiveConnection.CursorLocation = adUseClient
    Set rs = cmd.execute
End Sub

Public Sub MoveNext()
    rs.MoveNext
End Sub

Public Function hasNext() As Boolean ' 次レコードの有無
    hasNext = rs.EOF
End Function

Public Function RecordCount%() ' レコード数
    RecordCount = rs.RecordCount
End Function

Public Function fields(name$) As Variant ' フィールドの取り出し
    fields = rs(name)
End Function

Function get_rs() As ADODB.Recordset ' レコードセットオブジェクト
    Set get_rs = rs
End Function

Public Sub CloseRecordset()
    If rs.State <> adStateClosed Then rs.Close
    Set rs = Nothing
    sql.clear
    clear_param
End Sub

Private Sub clear_param() ' パラメータをクリア
    Dim i%
    
    For i = 0 To cmd.Parameters.Count - 1
        cmd.Parameters.Delete 0
    Next i
End Sub

Public Function get_fields() As UStringlist ' フィールドの一覧
    Dim fld As ADODB.Field

    flds.clear

    For Each fld In rs.fields
        flds.add fld.name
    Next

    Set get_fields = flds
End Function

Public Sub execute()
    OpenRecordset
End Sub

Private Sub Class_Terminate()
    CloseRecordset
End Sub
