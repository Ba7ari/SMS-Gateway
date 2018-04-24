VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Class_sms"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Public Function confirm_pesan(Optional ByVal nopesan As String, Optional noRek As String, Optional ByVal nmBank As String)
'cek no pesanan
Dim berangkat As String
Dim stkClass As Long
Dim getClassSetting As String
hp_konek = False: err_konek = False

            Set getData.rsReservasi = New ADODB.Recordset
            getData.rsReservasi.Open "SELECT statusPesan,BlmByr,No_Rek,NmBank,JmlDewasa,JmlAnak FROM reservasi WHERE StatusPesan = 'Hold' and No_Pesan LIKE '" & nopesan & "'", cn, 1, 3
            
            With getData.rsReservasi
                If .RecordCount = 0 Then
                        balasanPesan = "Anda sudah pernah melakukan Konfirmasi,pemesanan anda telah diproses"
                Else
                        !StatusPesan = "Confirm"
                        !BlmByr = 0
                        !No_Rek = noRek
                        !nmBank = nmBank
                        .Update
                        balasanPesan = "Terima Kasih,Konfirmasi anda berhasil,ketik status<spasi>(no.Booking),utk melihat status anda"
                        Tunggu
                End If
            End With
            getData.rsReservasi.Close
            cn.Close
            
            SMSDibalas strNoHP, balasanPesan
            balasanPesan = "": strNoHP = ""

Exit Function
pass:
    SMSDibalas strNoHP, balasanPesan
End Function

Function hitStokPnmpang(ByVal Kdpesan As String, ByVal nm_Class As String, ByVal jmlDewasa As Long, ByVal JmlBayi As Long)
Dim getSet As String
Dim rsClassBayi As New ADODB.Recordset
'hitClass
rsClassBayi.Open "SELECT StockClass,reservasi.JmlDewasa,JmlAnak FROM Class,Maskapai,dapat,reservasi " & _
                 "WHERE Class.No_flt = Maskapai.No_flt and dapat.No_flt = Maskapai.No_flt and " & _
                 "dapat.No_Pesan = reservasi.No_pesan and reservasi.No_pesan = '" & Kdpesan & "' and Class.Nm_Class LIKE '" & nm_Class & "'", cn, 1, 3
    
With rsClassBayi
If .RecordCount = 0 Then
        MsgBox "tidak ada Kode kelas"
Else
        If !jmlDewasa > 0 Then
                !StockClass = !StockClass - jmlDewasa
        End If
        If !jmlAnak > 0 Then
                !StockClass = !StockClass - JmlBayi
        End If
        
        .Update
End If
End With

rsClassBayi.Close
cn.Close
End Function

Public Function cek_status(ByVal nopesan As String)
hp_konek = False: err_konek = False
Set getData.rsReservasi = New ADODB.Recordset
getData.rsReservasi.Open "SELECT distinct reservasi.No_Pesan,StatusPesan,JmlDewasa,JmlAnak,dapat.TglBerangkat,Maskapai.No_flt,Jadwal,reservasi.TotalBayar  FROM Class,Maskapai,dapat,reservasi,Route WHERE Class.No_flt = Maskapai.No_flt and Maskapai.Kd_route = Route.Kd_route and dapat.No_flt = Maskapai.No_flt and " & _
"dapat.No_Pesan = reservasi.No_Pesan and reservasi.No_Pesan LIKE '" & nopesan & "'", cn, 1, 3
With getData.rsReservasi
    If .RecordCount = 0 Then
            balasanPesan = "Status anda salah"
    Else
        balasanPesan = "No.Booking :" & !No_Pesan & ", status :" & !StatusPesan & ", Jml.Dws: " & !jmlDewasa & ",Jml.Ank : " & !jmlAnak & ",Brgkt Tgl :" & !tglBerangkat & " pkl." & Format(!jadwal, "hh:mm") & ",No.Flt. " & !no_flt
        Tunggu
    End If
    
End With

    
    getData.rsReservasi.Close
    cn.Close
    SMSDibalas strNoHP, balasanPesan
    strNoHP = "": balasanPesan = ""
End Function
Public Function batal_pesan(ByVal nopesan As String)
Set getData.rsReservasi = New ADODB.Recordset
getData.rsReservasi.Open "SELECT * From reservasi,dapat WHERE reservasi.StatusPesan = 'Hold' and dapat.No_Pesan = reservasi.No_Pesan and reservasi.No_Pesan LIKE '" & nopesan & "' ", cn, 1, 3
With getData.rsReservasi
    If .RecordCount = 0 Then
            balasanPesan = "Anda sudah melakukan Konfirmasi,pembatalan tidak dpt dilakukan"
    Else
            Do While Not .EOF
                    .Delete
                    .MoveNext
            Loop
            balasanPesan = "Pembatalan Booking berhasil dilakukan"
    End If
End With
getData.rsReservasi.Close
cn.Close
SMSDibalas strNoHP, balasanPesan
strNoHP = ""
balasanPesan = ""
End Function
'Public Function kurangi_stok_bangku(ByVal noBooking As String) As Long
'Dim jmlBangku As Long
'Set getData.rsMaskapai = New ADODB.Recordset
'getData.rsMaskapai.Open "SELECT Class.StockClass From Class,Maskapai,dapat,reservasi WHERE Class.No_flt =Maskapai.No_Flt and Maskapai.No_Flt = dapat.No_Flt and dapat.No_Pesan = reservasi.No_Pesan and reservasi.No_Pesan LIKE '" & noBooking & "'", cn, 1, 3
'With getData.rsMaskapai
'        !stockClass = !stock - kurangi_stok_bangku
'        .Update
'        MsgBox !stock
'End With
'getData.rsMaskapai.Close
'cn.Close
'End Function
