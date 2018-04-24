VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ClassMaskapai"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Dim status_gagal_ As Boolean
Dim NoPsnBls As String
Dim GTBayar As Currency
Dim JmlChild As Integer


Public Function simpanMaskapai(ByVal noflt As String, ByVal kdRoute As String, ByVal jdwal As String)
Set getData.rsMaskapai = New ADODB.Recordset
getData.rsMaskapai.Open "SELECT * FROM Maskapai WHERE No_flt LIKE '" & noflt & "'", cn, 1, 3
With getData.rsMaskapai
If .RecordCount <> 0 Then
        MsgBox "Jadwal penerbangan ini sudah terpakai", vbInformation + vbSystemModal, "Informasi"
Else
    .AddNew
            !no_flt = noflt
            !Kd_route = kdRoute
            !jadwal = jdwal
    .Update
End If
End With
getData.rsMaskapai.Close
cn.Close
End Function
Public Function panggilMaskapai() As String()
Dim i As Integer, sMaskapai() As String
Set getData.rsMaskapai = New ADODB.Recordset
getData.rsMaskapai.Open "SELECT * FROM Maskapai", cn, 1, 3
With getData.rsMaskapai
    If .RecordCount <> 0 Then
           Do While Not .EOF
                ReDim Preserve sMaskapai(i)
                sMaskapai(i) = !no_flt & "#" & !Kd_route & "#" & Format(!jadwal, "HH:MM AMPM")
                i = i + 1
                .MoveNext
            Loop
    End If
End With
getData.rsMaskapai.Close
cn.Close
panggilMaskapai = sMaskapai
                
End Function
Public Function cek_jadwal(ByVal nmRoute As String)
Dim msgTemp As String

Set getData.rsMaskapai = New ADODB.Recordset
getData.rsMaskapai.Open "SELECT Maskapai.No_flt,Jadwal From Maskapai,Route WHERE Maskapai.Kd_Route = Route.Kd_Route and Route.Kd_Route LIKE '" & nmRoute & "'", cn, 1, 3
With getData.rsMaskapai
        If .RecordCount = 0 Then
                balasanPesan = "Jadwal Keberangkatan belum ada"
                
                
        Else
            Do While Not .EOF
                    msgTemp = msgTemp & !no_flt & " Jdwal terbng: " & Format(!jadwal, "HH:MM AMPM") & " "
                    .MoveNext
            Loop
                    balasanPesan = " No.Flt " & " " & msgTemp
                    
        End If
End With
getData.rsMaskapai.Close
cn.Close
SMSDibalas strNoHP, balasanPesan
strNoHP = "": balasanPesan = ""
End Function
Public Function pesan_tiket(ByVal nohp As String, ByVal noflt As String, ByVal nmClass, ByVal jmlDewasa As Long, ByVal tglBerangkat As String)
Dim nopesan As String, totalByr As Long, totalByrBayi As Long
Dim berangkat As String
Dim varNMClass As String
On Error GoTo pesan_error
'hp_konek = False: err_konek = False
status_gagal = False 'Untuk mengecek keanggotaan pelanggan

If Not status_anggota(nohp) = True Then 'jika nomor hp belum terdaftar
        balasanPesan = "Maaf,Permintaan ditolak, harap melakukan registrasi terlebih dahulu"
Else
            Set getData.rsAda = New ADODB.Recordset
            getData.rsAda.Open "SELECT StatusPesan FROM reservasi WHERE StatusPesan = 'Hold' and Hp LIKE '" & nohp & "'", cn, 1, 3
            With getData.rsAda
                  If .RecordCount <> 0 Then 'jika sudah melakukan pemesanan namun blm byr
                                balasanPesan = "Anda belum melakukan Konfirmasi pemesanan Tiket anda sebelumnya"
                  Else
                            nopesan = auto_kode_pesan 'lakukan proses pemesanan
                            Set getData.rsMaskapai = New ADODB.Recordset
                            getData.rsMaskapai.Open "SELECT * FROM reservasi WHERE No_Pesan LIKE '" & nopesan & "'", cn, 1, 3
                            With getData.rsMaskapai
                                If .RecordCount <> 0 Then 'jika nomor pemesanan/booking sudah ada
                                        balasanPesan = "Maaf,untuk sementara tidak dpt melakukan pemesanan"
                            
                                Else
                                    totalByr = hitung_totalPembayaran(noflt, jmlDewasa, conv_class(nmClass)) 'Hitung total pembayaran
                                       MsgBox totalByr
                                       If status_gagal = True Then
                                              balasanPesan = "Tarif No.flt belum ada"
                                        Else
                                                     
                                        .AddNew
                                                !No_Pesan = nopesan
                                                !hp = nohp
                                                !TglPesan = Format(Date, "dd/mm/yyyy")
                                                !JamPesan = Time
                                                !StatusPesan = "Hold"
                                                !BlmByr = totalByr
                                                !jmlDewasa = jmlDewasa
                                                !jmlAnak = 0
                                                
                                                !totalBayar = totalByr
                                        .Update
                                                 Set getData.rsReservasi = New ADODB.Recordset
                                                 getData.rsReservasi.Open "SELECT * FROM dapat", cn, 1, 3
                                                 With getData.rsReservasi
                                                       .AddNew
                                                               !No_Pesan = nopesan
                                                               !no_flt = noflt
                                                               berangkat = Mid(tglBerangkat, 1, 2) & "/" & Mid(tglBerangkat, 3, 2) & "/" & Mid(tglBerangkat, 5, 4)
                                                               !tglBerangkat = berangkat
                                                       .Update
                                                
                                                End With
                                                getData.rsReservasi.Close
                                                getData.rsMaskapai.Close
'                                                 NoPsnBls = nopesan
'                                                 GTBayar = Format(totalByr, "Rp ###,###,###")
                                                 balasanPesan = "Terima - kasih Pemesanan tiket berhasil No.Booking '" & nopesan & "',Total bayar   :" & Format(totalByr, "Rp #########") & ", Ketik Confirm<spasi>(No.Booking)<spasi>(no.rekening)<spasi>(nm.bank)"
                                                  Tunggu
                                                 
                                    End If
                            End If
                    End With
                End If
            End With
            getData.rsAda.Close
End If
cn.Close
'SMSDibalas strNoHP, balasanPesan
'balasanPesan = "": strNoHP = ""

Exit Function
pesan_error:
    MsgBox Err.Description
End Function
Public Function auto_kode_pesan() As String
Dim kode As String
Dim angka As Integer
Set getData.rsMaskapai = New ADODB.Recordset
getData.rsMaskapai.Open "SELECT No_Pesan FROM reservasi", cn, 1, 3
With getData.rsMaskapai
        If .RecordCount = 0 Then
            kode = "0001"
        Else
            .MoveLast
            angka = angka + Val(!No_Pesan + 1)
            Select Case angka
                    Case Is < 10
                            kode = "000" & angka
                    Case Is < 100
                            kode = "00" & angka
                    Case Is < 1000
                            kode = "0" & angka
                    Case Else
                        kode = angka
            End Select
        End If
        auto_kode_pesan = kode
End With
getData.rsMaskapai.Close
cn.Close
End Function
Public Function hitung_totalPembayaran(ByVal noflt As String, ByVal jmlPenmpang As Long, ByVal nmClass As String) As Long 'Dewasa
Dim totalByr As Long
Dim sellingTemp As Long
Dim getSetClass As String
status_gagal = False
'getSetClass = GetSetting("MandalaGateway", "Settings", "Class")

Set getData.rsClass = New ADODB.Recordset
getData.rsClass.Open "SELECT SellingPrice,StockClass FROM Class WHERE No_flt = '" & noflt & "' and Nm_Class LIKE '" & nmClass & "'", cn, 1, 3
     With getData.rsClass
         If .RecordCount <> 0 Then
                sellingTemp = !sellingprice
                totalByr = sellingTemp * jmlPenmpang
         Else
'                MsgBox "Gagal hitung total"
                status_gagal = True
         End If
      End With
getData.rsClass.Close
cn.Close
hitung_totalPembayaran = totalByr


End Function
Public Function hitung_totalPembayaranBayi(ByVal noflt As String, ByVal JmlBayi As Long) As Long
Dim totalBayar As Long
Dim sellingTemp As Long
        
        Set getData.rsClass = New ADODB.Recordset
        getData.rsClass.Open "SELECT SellingPrice FROM Class WHERE No_flt = '" & noflt & "' and Kd_Class LIKE 'INF'", cn, 1, 3
        With getData.rsClass
            If .RecordCount <> 0 Then
                sellingTemp = !sellingprice
                totalBayar = sellingTemp * JmlBayi
            Else
                status_gagal = True
                MsgBox "gagal hitung total"
                
            End If
        End With
        getData.rsClass.Close
        cn.Close
        hitung_totalPembayaranBayi = totalBayar

End Function
Public Function cari_maskapai(Optional ByVal noflt As String) As String()
Dim sData() As String, i As Integer
Set getData.rsMaskapai = New ADODB.Recordset
getData.rsMaskapai.Open "SELECT * FROM Maskapai WHERE No_Flt LIKE '" & noflt & "'", cn, 1, 3
With getData.rsMaskapai
        ReDim Preserve sData(i)
            sData(i) = !Kd_route & "#" & Format(!jadwal, "HH:MM AMPM")
            i = i + 1
End With
getData.rsMaskapai.Close
cn.Close
cari_maskapai = sData
End Function
Public Function ubah_maskapai(Optional ByVal noflt As String, Optional ByVal kdRoute, Optional ByVal jadwal As String)
Set getData.rsMaskapai = New ADODB.Recordset
getData.rsMaskapai.Open "SELECT * FROM Maskapai WHERE No_Flt LIKE '" & noflt & "'", cn, 1, 3
With getData.rsMaskapai
        If .RecordCount = 0 Then
                MsgBox "Data Maskapai tidak ditemukan", vbInformation + vbSystemModal, "Informasi"
        Else
            !Kd_route = kdRoute
            !jadwal = jadwal
            .Update
        End If
End With
getData.rsMaskapai.Close
cn.Close
End Function

Public Function status_pemesanan(ByVal nopesan As String)
Dim xData() As String, bagiData() As String, hasil As String
Dim i As Integer

Set getData.rsReservasi = New ADODB.Recordset
getData.rsReservasi.Open "SELECT StatusPesan,JmlDewasa,TglPesan,TotalBayar FROM reservasi WHERE No_Pesan LIKE '" & nopesan & "'", cn, 1, 3
With getData.rsReservasi
        If .RecordCount = 0 Then
                balasanPesan = "Nmr pesanan anda tidak dikenal"
        Else
                ReDim Preserve xData(i)
                xData(i) = !StatusPesan & " " & !jmlDewasa & " " & !TglPesan & " " & !totalBayar
                bagiData = Split(xData(i), " ")
                hasil = hasil & "Tgl.Pemesanan: " & bagiData(2) & ", " & "Status Pemesanan: " & bagiData(0) & ", " & "jml Penumpang: " & bagiData(1) & ", " & "Total pembayaran anda: " & Format(bagiData(3), "Rp ###,###,###")
                i = i + 1
                balasanPesan = hasil & "- Terima kasih"
        End If
End With
getData.rsReservasi.Close
cn.Close

End Function

Public Function hapus_maskapai()
Set getData.rsMaskapai = New ADODB.Recordset
getData.rsMaskapai.Open "SELECT * FROM Maskapai", cn, 1, 3
With getData.rsMaskapai
    Do While Not .EOF
            .Delete
            .MoveNext
    Loop
End With

getData.rsMaskapai.Close
cn.Close
End Function
Public Function status_anggota(ByVal nohp As String) As Boolean
Set getData.rsPelanggan = New ADODB.Recordset
getData.rsPelanggan.Open "SELECT Hp FROM Pelanggan WHERE Hp LIKE '" & nohp & "'", cn, 1, 3
With getData.rsPelanggan
    If .RecordCount = 0 Then
            status_anggota = False
    Else
            status_anggota = True
    End If
End With
getData.rsPelanggan.Close
cn.Close
End Function
Public Function edit_nopesan(ByVal nopesan As String, Optional ByVal jmlDewasa As Long, Optional ByVal jmlAnak As Long, Optional ByVal opsi As Integer)
Dim strTemp As String

Set getData.rsReservasi = New ADODB.Recordset
getData.rsReservasi.Open "SELECT JmlDewasa,JmlAnak FROM reservasi WHERE StatusPesan = 'Hold' and No_Pesan LIKE '" & nopesan & "'", cn, 1, 3
With getData.rsReservasi
        If .RecordCount = 0 Then
                balasanPesan = "No booking anda salah"
        Else
                Select Case opsi
                        Case 1 'penumpang dewasa
                            
                            !jmlDewasa = jmlDewasa
                            .Update
                             strTemp = hit_ubahPenmpang(nopesan, jmlDewasa, , 1)
                            
                         Case 2 'Anak
                            !jmlAnak = jmlAnak
                            .Update
                            strTemp = hit_ubahPenmpang(nopesan, , jmlAnak, 2)
                End Select
                balasanPesan = "Perubahan status booking Total Pembayaran" & strTemp
        End If
        
        
  End With
  getData.rsReservasi.Close
  cn.Close
  SMSDibalas strNoHP, balasanPesan
  balasanPesan = ""
  strNoHP = ""
  strTemp = ""
End Function
Public Function hit_ubahPenmpang(ByVal nopesan As String, Optional hitDewasa As Long, Optional hitAnak As Long, Optional ByVal pilih As Integer) As Long
Dim jlh As Long
Dim getClass As String
getClass = GetSetting("MandalaGateway", "Settings", "Class")
Set getData.rsHitung = New ADODB.Recordset
getData.rsHitung.Open "select reservasi.jmlDewasa,JmlAnak,TotalBayar,Class.SellingPrice from Class,Maskapai,dapat,reservasi WHERE Class.Kd_Class = '" & getClass & "' and Class.No_flt=Maskapai.No_flt and dapat.No_flt = Maskapai.No_flt and dapat.No_Pesan = reservasi.No_Pesan and reservasi.StatusPesan = 'Hold' and reservasi.No_Pesan LIKE '" & nopesan & "'", cn, 1, 3
        With getData.rsHitung
                Select Case pilih
                        Case 1
                                jlh = !sellingprice * hitDewasa
                                jlh = jlh + !totalBayar
                        Case 2
                                jlh = !sellingprice * hitAnak
'                                jlh = jlh + !totalBayar
                End Select
        End With
getData.rsHitung.Close
cn.Close
hit_ubahPenmpang = jlh
End Function
Public Function cariHps_maskapai(ByVal noflt As String, ByVal kdrute As String)
Set getData.rsMaskapai = New ADODB.Recordset
getData.rsMaskapai.Open "SELECT * FROM Maskapai WHERE Kd_route = '" & kdrute & "' and No_flt LIKE '" & noflt & "'", cn, 1, 3

With getData.rsMaskapai
        If .RecordCount = 0 Then
                MsgBox "No.Flt yang anda cari tidak ada", vbInformation + vbSystemModal, "Informasi"
        Else
            Do While Not .EOF
                    .Delete
                    .MoveNext
            Loop
        End If
        
End With
getData.rsMaskapai.Close
cn.Close
End Function
Public Property Get status_gagal() As Boolean
status_gagal = status_gagal_
End Property

Public Property Let status_gagal(ByVal vNewValue As Boolean)
status_gagal_ = vNewValue
End Property
Public Function Hapus_NoFlt(ByVal noflt As String, ByVal kdClass As String)
Set getData.rsMaskapai = New ADODB.Recordset
getData.rsMaskapai.Open "SELECT * FROM Class WHERE No_Flt = '" & noflt & "' and Kd_Class LIKE '" & kdClass & "'", cn, 1, 3
With getData.rsMaskapai
    If .RecordCount = 0 Then
            MsgBox "No.Flt and Class tidak dikenal", vbInformation + vbSystemModal, "Informasi"
    Else
        Do While Not .EOF
            .Delete
            .MoveNext
        Loop
    End If
End With
getData.rsMaskapai.Close
cn.Close
End Function
Public Function cari_Semua_maskapai(Optional ByVal noflt As String) As String()
Dim sData() As String, i As Integer
Dim tmpJadwal As Variant
Set getData.rsMaskapai = New ADODB.Recordset
getData.rsMaskapai.Open "SELECT distinct Maskapai.No_flt,Route.kd_route,Class.nm_Class,Maskapai.Jadwal,Class.SellingPrice,StockClass FROM Maskapai,Route,Class WHERE Maskapai.No_flt = Class.No_flt and Maskapai.Kd_route = Route.Kd_route and Class.No_flt like '" & noflt & "'", cn, 1, 3
With getData.rsMaskapai
    Do While Not .EOF
        ReDim Preserve sData(i)
'        tmpJadwal = FormatDateTime(!jadwal, vbShortTime)
            sData(i) = !no_flt & "#" & !Kd_route & "#" & !nm_Class & "#" & Format(!jadwal, "HH:MM AMPM") & "#" & !sellingprice & "#" & !StockClass
            i = i + 1
            .MoveNext
    Loop
End With
getData.rsMaskapai.Close
cn.Close
cari_Semua_maskapai = sData
End Function

Public Function cari_noflt_class(ByVal noflt As String, ByVal kdClass As String) As String()
Dim sData() As String, i As Integer
Set getData.rsMaskapai = New ADODB.Recordset
getData.rsMaskapai.Open "SELECT * FROM Class WHERE No_Flt = '" & noflt & "' and kd_class LIKE '" & kdClass & "'", cn, 1, 3
With getData.rsMaskapai
    
            If .RecordCount = 0 Then
                    MsgBox "No.Flt dan Kode kelas tidak ada", vbInformation + vbSystemModal, "Pencarian gagal"
            Else
                Do While Not .EOF
                    ReDim Preserve sData(i)
                    sData(i) = !nm_Class & "#" & !fareBasic & "#" & !fuelSurc & "#" & !ppn & "#" & !iwjr & "#" & !sellingprice & "#" & !StockClass
                    i = i + 1
                    .MoveNext
                Loop
            End If
        
   
End With
getData.rsMaskapai.Close
cn.Close
cari_noflt_class = sData
End Function

Public Function ubah_noflt_class(ByVal noflt As String, ByVal kdClass As String, ByVal fareBasic As Currency, ByVal fuelSurc As Currency, ByVal ppn As Currency, ByVal iwjr As Currency, ByVal sellingprice As Currency, ByVal stock As Long, ByVal nmClass As String)
On Error GoTo err_handler
Set getData.rsMaskapai = New ADODB.Recordset
getData.rsMaskapai.Open "SELECT * FROM Class WHERE No_Flt = '" & noflt & "' and Kd_Class LIKE '" & kdClass & "'", cn, 1, 3
With getData.rsMaskapai
    !no_flt = noflt
    !Kd_Class = kdClass
    !nm_Class = nmClass
    !fareBasic = fareBasic
    !fuelSurc = fuelSurc
    !ppn = ppn
    !sellingprice = sellingprice
    !StockClass = stock
    .Update
End With
getData.rsMaskapai.Close
Exit Function
err_handler:
    MsgBox Err.Description & Err.Number, vbCritical + vbSystemModal, "Kesalahan"
End Function
Public Function pesan_tiket_plus_anak(ByVal nohp As String, ByVal noflt As String, ByVal nmClass As String, ByVal jmlDewasa As Long, ByVal jmlAnak As Long, ByVal arrive As String)
Dim jumTotDewasa As Currency, jumTotBayi As Currency, jumGrandTotal As Currency
Dim tmpVarClass As String
Dim nopesan As String, berangkat As String
status_gagal = False
On Error GoTo err_handler
'If Not status_anggota(nohp) = True Then 'jika nomor hp belum terdaftar
'        balasanPesan = "Maaf,Permintaan ditolak, harap melakukan registrasi terlebih dahulu"
'Else
            Set getData.rsAda = New ADODB.Recordset
            getData.rsAda.Open "SELECT StatusPesan FROM reservasi WHERE StatusPesan = 'Hold' and Hp LIKE '" & nohp & "'", cn, 1, 3
            With getData.rsAda
                        If .RecordCount <> 0 Then 'jika sudah melakukan pemesanan namun blm byr
                                balasanPesan = "Anda belum melakukan Konfirmasi pemesanan Tiket anda sebelumnya"
                        Else
                            nopesan = auto_kode_pesan 'lakukan proses pemesanan
                            Set getData.rsMaskapai = New ADODB.Recordset
                            getData.rsMaskapai.Open "SELECT * FROM reservasi WHERE No_Pesan LIKE '" & nopesan & "'", cn, 1, 3
                            With getData.rsMaskapai
                                If .RecordCount <> 0 Then
                                        balasanPesan = "Maaf,tidak dapat melanjutkan pemesanan"
                                       
                                Else
                                   
                                   jumTotDewasa = hitung_totalPembayaran(noflt, jmlDewasa, conv_class(nmClass)) 'Hitung total pembayaran
                                   jumTotBayi = hitung_totalPembayaranBayi(noflt, jmlAnak)
                                   jumGrandTotal = jumTotDewasa + jumTotBayi

                                   .AddNew
                                        !No_Pesan = nopesan
                                        !hp = nohp
                                        !TglPesan = Format(Date, "dd/mm/yyyy")
                                         !JamPesan = Time
                                        !StatusPesan = "Hold"
                                        !BlmByr = jumGrandTotal
    
                                        !jmlDewasa = jmlDewasa
                                        !jmlAnak = jmlAnak '2 thn kebawah
                                        !totalBayar = jumGrandTotal
                                     .Update
                                     Set getData.rsReservasi = New ADODB.Recordset
                                     getData.rsReservasi.Open "SELECT * FROM dapat", cn, 1, 3
                                     With getData.rsReservasi
                                           .AddNew
                                                   !No_Pesan = nopesan
                                                   !no_flt = noflt
                                                   berangkat = Mid(arrive, 1, 2) & "/" & Mid(arrive, 3, 2) & "/" & Mid(arrive, 5, 4)
                                                   !tglBerangkat = berangkat
                                           .Update
                                    End With
                                    getData.rsReservasi.Close
                                 
                                 balasanPesan = "Terima-kasih Pemesanan tiket berhasil No.Booking '" & nopesan & "',jml bayi : '" & jmlAnak & "' Total bayar: '" & Format(jumGrandTotal, "Rp ###,###,###") & ",Ketik Confirm<spasi>(No.Booking)<spasi>(no.rekening)<spasi>(nm.bank)"
                                 MsgBox Len(balasanPesan)
                                 Tunggu
                                  

                               End If
                                
                            End With
                            getData.rsMaskapai.Close
                        End If
        End With

'End If
cn.Close

'SMSDibalas strNoHP, balasanPesan
'balasanPesan = "": strNoHP = ""
Exit Function
err_handler:
    MsgBox Err.Description & Err.Number, vbCritical + vbSystemModal, "Modul - PesanTiketPlusAnak"
    
End Function
Public Function edit_tiket(ByVal nohp As String, ByVal noBooking As String, ByVal nmClass As String, ByVal jmlDewasa As Long, ByVal jmlAnak As Long)
Dim jumTotDewasa As Currency, jumTotBayi As Currency, jumGrandTotal As Currency
Set getData.rsReservasi = New ADODB.Recordset
getData.rsReservasi.Open "SELECT * FROM reservasi,dapat WHERE dapat.No_Pesan = reservasi.No_Pesan and reservasi.No_pesan LIKE '" & noBooking & "'", cn, 1, 3
With getData.rsReservasi
 If !StatusPesan = "Hold" Then
    jumTotDewasa = hitung_totalPembayaran(!no_flt, jmlDewasa, conv_class(nmClass))
'    MsgBox jumTotDewasa
    jumTotBayi = hitung_totalPembayaranBayi(!no_flt, jmlAnak)
    jumGrandTotal = jumTotDewasa + jumTotBayi
    
    If .RecordCount <> 0 Then
        !jmlDewasa = jmlDewasa
        !jmlAnak = jmlAnak
        !totalBayar = jumGrandTotal
        .Update
        'ambil modul pengurangan stok
        balasanPesan = "Perubahan berhasil,jml.Penmpang Dewasa : " & jmlDewasa & " Class : " & nmClass & ", Jml.Penmpang Bayi " & jmlAnak & " Total :" & Format(jumGrandTotal, "Rp ###,###,###")
    Else
        balasanPesan = "Maaf,No.Booking anda tidak terdaftar"
    End If
  Else
        balasanPesan = "Maaf,anda sudah melakukan konfirmasi,perubahan data tidak dapat dilakukan"
  End If
End With


getData.rsReservasi.Close
cn.Close
SMSDibalas strNoHP, balasanPesan
balasanPesan = "": strNoHP = ""
End Function
Public Function conv_class(ByVal varClass As String) As String
Dim tmpClass As String
 Select Case UCase(varClass)
     Case Is = "FLEX"
         tmpClass = "Flexi"
     Case Is = "SAVER"
         tmpClass = "Saver"
     Case Is = "INF"
         tmpClass = "Infant"
End Select
conv_class = tmpClass
End Function
Public Function reply_Pemesanan_Adult_Child()
Dim sTrGTBayar As String

balasanPesan = "Terima - kasih Pemesanan tiket berhasil No.Booking '" & NoPsnBls & "',jml bayi : '" & JmlChild & "' Total bayar  :" & GTBayar & ", Ketik Confirm<spasi>(No.Booking)<spasi>(no.rekening)<spasi>(nm.bank)"
'SMSDibalas strNoHP, balasanPesan
'strNoHP = "": balasanPesan = ""
'NoPsnBls = "": sTrGTBayar = ""
Tunggu
End Function
