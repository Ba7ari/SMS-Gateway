Attribute VB_Name = "modMain"
Option Explicit

Public Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)

Public LoggedIn As Boolean

Public AForm As ALib.FormUtil

Public Type TypKategori
    Kode As Long
    Teks As String * 20
    Pesan As String * 160
    Aktif As Boolean
End Type
Public KategoriList() As TypKategori
Public jKategori As Integer
Public KategoriIdx As Integer

Dim MyMsgDll As msgdll.clsMessage

Public Const smsBaru = 0
Public Const smsTerbaca = 1
Public Const smsDiproses = 2
Public Const smsDitandai = 3
Public Const strBaru = "Baru"
Public Const strTerbaca = "Terbaca"
Public Const strDiproses = "Diproses"
Public Const strDitandai = "Ditandai"

Public OC As ADODB.Connection

Sub InitApp()
Set OC = New ADODB.Connection

Set AForm = New ALib.FormUtil
LoggedIn = False
Set MyMsgDll = New msgdll.clsMessage

On Error GoTo KoneksiErr

OC.Open "DSN=SMSAR"

WriteLog "Memulai aplikasi"

Exit Sub
   
KoneksiErr:
    TampilkanPesan "Koneksi ke Database Server tidak terlaksana" & vbCrLf & _
                    "Pesan: " & Err.Description & vbCrLf & _
                    "Kode: " & Err.Number
    End
End Sub

Public Sub DeInitApp()
Set AForm = Nothing

WriteLog "Menutup aplikasi"
OC.Close
If OC.State = adStateOpen Then Set OC = Nothing
Set MyMsgDll = Nothing
End Sub

Function Decrypt(ByVal Teks As String) As String
Dim i As Integer, j As Integer
Dim s As String

j = 197
s = ""

For i = 1 To Len(Teks)
    s = s & Chr(Asc(Mid(Teks, i, 1)) Xor j)
    j = j + 1
    If j > 255 Then j = 1
Next

Decrypt = s
End Function

Function IsNumberValid(ByVal Nomer As String) As Boolean
If Len(Nomer) < 1 Then
    IsNumberValid = False
ElseIf Not IsNumeric(Nomer) Then
    IsNumberValid = False
ElseIf (Left(Nomer, 1) = "+") And (Mid(Nomer, 2, 1) = "0") Then
    IsNumberValid = False
ElseIf (Left(Nomer, 1) = "+") And IsNumeric(Mid(Nomer, 2)) Then
    IsNumberValid = True
ElseIf IsNumeric(Nomer) Then
    IsNumberValid = True
Else
    IsNumberValid = False
End If
End Function

Function VldTeks(ByVal Teks As String) As String
VldTeks = AddQuote(DoubleQuote(Teks))
End Function

Function AddQuote(ByVal Teks As String) As String
AddQuote = "'" & Trim(Teks) & "'"
End Function

Function DoubleQuote(ByVal Teks As String) As String
DoubleQuote = Replace(Teks, "'", "''")
End Function

Sub TampilkanPesan(ByVal Pesan As String)
With MyMsgDll
    .AddButton 0, "Tutup", True
    .MessageBoxEx Pesan
    ', , Msg_Information, , , "Square721 BT", , 10, 12, , , False, True, Msg_Center, Msg_Center, True, , , , , True
End With
End Sub

Function TampilkanTanya(ByVal Tanya As String) As Long
With MyMsgDll
    .AddButton 0, "Ya", True, False
    .AddButton 1, "Tidak", False, True
    TampilkanTanya = .MessageBoxEx(Tanya) ', , Msg_Information, , , "Square721 BT", , 10, 12, , , False, True, Msg_Center, Msg_Center, True, , , , , True)
End With
End Function

Public Sub WriteLog(ByVal Teks As String)
OC.Execute "insert into logtable(waktu,writer,logtext) values('" & Format(Now, "yyyy/MM/dd hh:mm:ss") & "','EnterpriseManager'," & VldTeks(Teks) & ")"
End Sub

Public Function GetMainNumber(ByVal Nomor As String) As String
If Left(Nomor, 1) = "+" Then
    GetMainNumber = Mid(Nomor, 4)
Else
    GetMainNumber = Mid(Nomor, 2)
End If
End Function

Public Sub KirimSMS(ByVal Tujuan As String, ByVal Pesan As String)
On Error GoTo KirimErr

OC.Execute "insert into outbox(tujuan,teks) " & _
            "values(" & VldTeks(Tujuan) & "," & VldTeks(Pesan) & ")"
            
Exit Sub
KirimErr:
    TampilkanPesan "Gagal menyimpan ke Outbox" & vbCrLf & _
                    "Pesan: " & Err.Description & vbCrLf & _
                    "Kode: " & Err.Number
End Sub

Public Sub SimpanFileText(ByVal NF As String)
Dim l As Integer
Dim RS As ADODB.Recordset

l = FreeFile
Open NF For Output As #l

On Error GoTo KirimErr

Set RS = New ADODB.Recordset
RS.Open "select * from temppesan", OC
While Not RS.EOF
    Print #l, RS("pengirim") & vbTab; RS("waktu") & vbTab; RS("teks")
    RS.MoveNext
Wend
RS.Close
Close #l
Set RS = Nothing
Exit Sub
KirimErr:
    TampilkanPesan "Gagal melakukan export" & vbCrLf & _
                    "Pesan: " & Err.Description & vbCrLf & _
                    "Kode: " & Err.Number
End Sub

Public Sub SimpanFileXML(ByVal NF As String)
Dim l As Integer
Dim RS As ADODB.Recordset

l = FreeFile
Open NF For Output As #l

On Error GoTo KirimErr

Set RS = New ADODB.Recordset

Print #l, "<?xml version=""1.0""?>"
Print #l, "<Workbook xmlns=""urn:schemas-microsoft-com:office:spreadsheet"""
Print #l, "xmlns:o=""urn:schemas-microsoft-com:office:office"""
Print #l, "xmlns:x=""urn:schemas-microsoft-com:office:excel"""
Print #l, "xmlns:ss=""urn:schemas-microsoft-com:office:spreadsheet"""
Print #l, "xmlns:html=""http://www.w3.org/TR/REC-html40"">"
Print #l, "<Worksheet ss:Name=""Sheet1"">"
Print #l, "<Table ss:ExpandedColumnCount=""3"">"
Print #l, "<Row>"
Print #l, "<Cell><Data ss:Type=""String"">Pengirim</Data></Cell>"
Print #l, "<Cell><Data ss:Type=""String"">Waktu</Data></Cell>"
Print #l, "<Cell><Data ss:Type=""String"">Pesan</Data></Cell>"
Print #l, "</Row>"
RS.Open "select * from temppesan", OC
While Not RS.EOF
    Print #l, "<Row>"
    Print #l, "<Cell><Data ss:Type=""String"">" & RS("pengirim") & "</Data></Cell>"
    Print #l, "<Cell><Data ss:Type=""String"">" & RS("waktu") & "</Data></Cell>"
    Print #l, "<Cell><Data ss:Type=""String"">" & RS("teks") & "</Data></Cell>"
    Print #l, "</Row>"
    RS.MoveNext
Wend
RS.Close
Print #l, "</Table>"
Print #l, "</Worksheet>"
Print #l, "</Workbook>"
Close #l
Set RS = Nothing
Exit Sub
KirimErr:
    TampilkanPesan "Gagal melakukan export" & vbCrLf & _
                    "Pesan: " & Err.Description & vbCrLf & _
                    "Kode: " & Err.Number
End Sub

Public Function AcakKarakter(ByVal PanjangMin As Integer, ByVal PanjangMax As Integer, ByVal ChrMin As Integer, ByVal ChrMax As Integer) As String
Dim i As Integer, j As Integer
Dim c  As Integer
Dim s As String

i = PanjangMin + Int(Rnd() * (PanjangMax - PanjangMin))
If i < PanjangMin Then i = PanjangMin
If i > PanjangMax Then i = PanjangMax
s = ""
For j = 1 To i
    c = ChrMin + Int(Rnd() * (ChrMax - ChrMin))
    If c < ChrMin Then c = ChrMin
    If c > ChrMax Then c = ChrMax
    s = s & Chr(c)
Next
AcakKarakter = s
End Function
