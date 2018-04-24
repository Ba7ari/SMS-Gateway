VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ClassAmbil"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Public Function ambil_tiket(Optional ByVal Kdpesan As String) As String()
Dim i As Integer
Dim sData() As String

Set getData.rsMaskapai = New ADODB.Recordset
If Len(Kdpesan) <> 0 Then
    getData.rsMaskapai.Open "SELECT * FROM reservasi WHERE No_Pesan LIKE '" & Kdpesan & "'", cn, 1, 3
Else
    getData.rsMaskapai.Open "SELECT * FROM reservasi", cn, 1, 3
End If
With getData.rsMaskapai
            If .RecordCount = 0 Then
                MsgBox "Data tidak ditemukan", vbInformation + vbSystemModal, "Informasi"
            Else
                Do While Not .EOF
                    ReDim Preserve sData(i)
                        sData(i) = !hp & "#" & !TglPesan & "#" & !StatusPesan & "#" & !BlmByr & "#" & !jmlDewasa & "#" & !totalBayar & "#" & !No_Rek & "#" & !nmBank
                        i = i + 1
                        .MoveNext
                Loop
                
            End If
End With
getData.rsMaskapai.Close
cn.Close
ambil_tiket = sData
End Function

Public Function Add_reservasi() As String()
Dim i As Integer
Dim sData() As String
Set getData.rsReservasi = New ADODB.Recordset
getData.rsReservasi.Open "SELECT No_Pesan FROM reservasi", cn, 1, 3
With getData.rsReservasi
        Do While Not .EOF
                ReDim Preserve sData(i)
                sData(i) = !No_Pesan
                i = i + 1
                .MoveNext
        Loop
End With
getData.rsReservasi.Close
cn.Close
 Add_reservasi = sData
End Function
Public Function cari__pemesan(Optional ByVal nopesan As String) As String
Dim sData As String
Set getData.rsReservasi = New ADODB.Recordset
getData.rsReservasi.Open "SELECT Pelanggan.Hp,Nama FROM Pelanggan,reservasi WHERE reservasi.No_Pesan LIKE '" & nopesan & "'", cn, 1, 3
With getData.rsReservasi
    If .RecordCount = 0 Then
            MsgBox "Identitas pemesan tidak ditemukan", vbInformation + vbSystemModal
    Else
        sData = !hp & " " & !nama
    End If
End With
getData.rsReservasi.Close
cn.Close
cari__pemesan = sData
End Function

Public Function hapus_reservasi()
Set getData.rsReservasi = New ADODB.Recordset
getData.rsReservasi.Open "SELECT * FROM reservasi", cn, 1, 3
With getData.rsReservasi
    Do While Not .EOF
            .Delete
            .MoveNext
    Loop
End With

Set getData.rsAda = New ADODB.Recordset
getData.rsAda.Open "SELECT * FROM dapat", cn, 1, 3
With getData.rsAda
    Do While Not .EOF
            .Delete
            .MoveNext
    Loop
End With
            
getData.rsReservasi.Close
getData.rsAda.Close
cn.Close
End Function
