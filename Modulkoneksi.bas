Attribute VB_Name = "Modulkoneksi"
Public conn As New ADODB.Connection
Public RS As New ADODB.Recordset

Sub Koneksi()
On Error GoTo gagal

Set conn = New ADODB.Connection
Set RS = New ADODB.Recordset
conn.Open "Provider=Microsoft.Jet.OLEDB.4.0;Data Source= " & App.Path & "\Database.mdb;persist security info=false"
Exit Sub

gagal:
MsgBox "Gagal menghubungkan ke Database! Kesalahan pada:" & Err.Description, vbCritical, "Peringatan"
End Sub
