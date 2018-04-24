VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ClassPelanggan"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Dim kosongPelanggan As Boolean

Public Function simpan_pelanggan(ByVal hp As String, ByVal nm As String, ByVal alamat As String)
hp_konek = False: err_konek = False
Set getData.rsPelanggan = New ADODB.Recordset
getData.rsPelanggan.Open "SELECT * FROM Pelanggan WHERE HP LIKE '" & hp & "'", cn, 1, 3
With getData.rsPelanggan
    If .RecordCount <> 0 Then
            balasanPesan = "Nomor Handphone Anda Sudah Terdaftar"
            
    Else
           
                    .AddNew
                            !hp = hp
                            !nama = nm
                            !alamat = alamat
                    .Update
                    balasanPesan = "Terima kasih tlh bergbng di komunitas Mandala Airlines,Ktk: Booking<spasi>(No.FLT)<spasi>(Jml.Dws)<spasi>(jml.Ank)<spasi>(Tgl.brkt),utk bkg tkt"
            
    End If
End With
getData.rsPelanggan.Close
cn.Close
SMSDibalas strNoHP, balasanPesan
strNoHP = ""
balasanPesan = ""
End Function
Public Function AddDataPelanggan() As String()
Dim i As Integer
Dim sData() As String
Set getData.rsPelanggan = New ADODB.Recordset
getData.rsPelanggan.Open "SELECT HP,Nama,Alamat FROM Pelanggan", cn, 1, 3
With getData.rsPelanggan
    Do While Not .EOF
        ReDim Preserve sData(i)
        sData(i) = !hp & "#" & !nama & "#" & !alamat
        i = i + 1
        .MoveNext
    Loop
End With
getData.rsPelanggan.Close
cn.Close
AddDataPelanggan = sData
End Function

Public Function cari_status_pelanggan(ByVal nama As String) As String()
Dim i As Integer, sNama() As String
Set getData.rsPelanggan = New ADODB.Recordset
getData.rsPelanggan.Open "SELECT Nama FROM Pelanggan WHERE Nama LIKE '" & nama & "%'", cn, 1, 3
With getData.rsPelanggan
        If .RecordCount = 0 Then
                status_kosong = True
                Exit Function
        Else
        Do While Not .EOF
                ReDim Preserve sNama(i)
                sNama(i) = !nama
                i = 1 + 1
                .MoveNext
        Loop
        End If
End With
getData.rsPelanggan.Close
cn.Close
cari_status_pelanggan = sNama
End Function


Public Property Get status_kosong() As Boolean
status_kosong = kosongPelanggan
End Property

Public Property Let status_kosong(ByVal vNewValue As Boolean)
kosongPelanggan = vNewValue
End Property

Public Function hapus_pelanggan()
Set getData.rsPelanggan = New ADODB.Recordset
getData.rsPelanggan.Open "SELECT * FROM Pelanggan", cn, 1, 3
With getData.rsPelanggan
    Do While Not .EOF
            .Delete
            .MoveNext
    Loop
End With
getData.rsPelanggan.Close
cn.Close
End Function
Public Function simpan_pelanggan_manual(ByVal nohp As String, ByVal nmplg As String, ByVal almt As String)
Set getData.rsPelanggan = New ADODB.Recordset
getData.rsPelanggan.Open "SELECT * FROM Pelanggan WHERE Hp like '" & nohp & "'", cn, 1, 3
With getData.rsPelanggan
    If .RecordCount <> 0 Then
            MsgBox "No.Hp '" & nohp & "' sudah pernah di-Input", vbInformation + vbSystemModal, "Informasi"
    Else
        .AddNew
                !hp = nohp
                !nama = nmplg
                !alamat = almt
        .Update
    End If
End With
getData.rsPelanggan.Close
cn.Close
End Function

Public Function hapus_nomor_pelanggan(ByVal nohp As String, Optional nama As String, Optional alamat As String, Optional opsi As Integer)
Set getData.rsPelanggan = New ADODB.Recordset
getData.rsPelanggan.Open "SELECT * FROM Pelanggan Where Hp like '" & nohp & "'", cn, 1, 3
With getData.rsPelanggan
    If .RecordCount <> 0 Then
            Select Case opsi
                    Case 1
                            Do While Not .EOF
                                .Delete
                                .MoveNext
                            Loop
                    Case 2
                            !hp = nohp
                            !nama = nama
                            !alamat = alamat
                            .Update
            End Select
    Else
            MsgBox "Nomor Hp pelanggan tidak ditemukan", vbInformation + vbSystemModal, "Informasi"
    End If
End With
getData.rsPelanggan.Close
cn.Close
End Function
