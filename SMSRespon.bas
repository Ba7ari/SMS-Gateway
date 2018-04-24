Attribute VB_Name = "SMSRespon"
Option Explicit
Private mvarIndexSend As String
Public strPesan As String
Public strNoHP As String
Public status_konek As Boolean
Public status_signal As Boolean

Public hp_konek As Boolean, err_konek As Boolean
Public jmlAkhir         As Integer
Public arrKirim()       As String
Public arrDel() As String
Public balasanPesan As Variant
Public mvarMsgReport As String
Public mvarMsgType As String
Public mvarMsgTime As String
Public nomHP() As String
Public intUrut As Integer
Public intTambah As Integer
Public jmlAwal As Integer
Public Function SMSIn(ByVal Msg As String) As String
    Dim FO As String, PID As String, DCS As String, SCTS As String
    Dim UDL As String, UD As String, SCTS_Tgl As String, SCTS_Jam As String
    Dim lnSCA As String, typeSCA As String, noSCA As String
    Dim newMsg As String, lnOA As String, typeOA As String, noOA As String
    Dim SCTS_a As String, SCTS_Tgl_a As String, SCTS_Jam_a As String
    On Error GoTo pesan_error
    newMsg = Msg
    lnSCA = HexToDec(Left(Msg, 2)) * 2  'length of SCA
    newMsg = Right(newMsg, Len(newMsg) - 2)
    
    typeSCA = Left(newMsg, 2) '91:int,81:local
    newMsg = Right(newMsg, Len(newMsg) - 2)
    
    noSCA = RevNum(Left(newMsg, lnSCA - 2)) 'service center
    If UCase(Right(noSCA, 1)) = "F" Then noSCA = Left(noSCA, Len(noSCA) - 1)
    newMsg = Right(newMsg, Len(newMsg) - lnSCA + 2)
    
    FO = Left(newMsg, 2)
    newMsg = Right(newMsg, Len(newMsg) - 2)
    If FO = "06" Then 'code of send report (Indonesia)
        mvarIndexSend = HexToDec(Left(newMsg, 2))
        newMsg = Right(newMsg, Len(newMsg) - 2)
    End If
    
    'Origine Address
    lnOA = HexToDec(Left(newMsg, 2))
    If lnOA Mod 2 <> 0 Then lnOA = lnOA + 1
    newMsg = Right(newMsg, Len(newMsg) - 2)
    typeOA = Left(newMsg, 2)
    newMsg = Right(newMsg, Len(newMsg) - 2)
    
    noOA = Left(newMsg, lnOA)
    If typeOA = "D0" Then
        noOA = CharHex(noOA, 7)
    Else
        noOA = RevNum(noOA)
        If UCase(Right(noOA, 1)) = "F" Then noOA = Left(noOA, Len(noOA) - 1)
    End If
    newMsg = Right(newMsg, Len(newMsg) - lnOA)
   
    If FO <> "06" Then 'if not report message
        PID = Left(newMsg, 2)
        newMsg = Right(newMsg, Len(newMsg) - 2)
    
        DCS = Left(newMsg, 2)
        newMsg = Right(newMsg, Len(newMsg) - 2)
    
        SCTS = RevNum(Left(newMsg, 14))
        SCTS_Tgl = Mid(SCTS, 3, 2) & "/" & Mid(SCTS, 5, 2) & "/20" & Mid(SCTS, 1, 2) 'mm/dd/yyyy,jj:mn
        SCTS_Jam = Mid(SCTS, 7, 2) & ":" & Mid(SCTS, 9, 2) & ":" & Mid(SCTS, 11, 2) 'hh:mm:dd
        newMsg = Right(newMsg, Len(newMsg) - 14)
    
        UDL = CInt(HexToDec(Left(newMsg, 2)))
        newMsg = Right(newMsg, Len(newMsg) - 2)
    
        UD = CharHex(newMsg, 7)
        UD = Left(UD, UDL)
    Else
        SCTS = RevNum(Left(newMsg, 14))
        SCTS_Tgl = Mid(SCTS, 3, 2) & "/" & Mid(SCTS, 5, 2) & "/20" & Mid(SCTS, 1, 2) 'mm/dd/yyyy,jj:mn
        SCTS_Jam = Mid(SCTS, 7, 2) & ":" & Mid(SCTS, 9, 2) & ":" & Mid(SCTS, 11, 2) 'hh:mm:dd
        newMsg = Right(newMsg, Len(newMsg) - 14)
        
        SCTS_a = RevNum(Left(newMsg, 14))
        SCTS_Tgl_a = Mid(SCTS_a, 3, 2) & "/" & Mid(SCTS_a, 5, 2) & "/20" & Mid(SCTS_a, 1, 2) 'mm/dd/yyyy,jj:mn
        SCTS_Jam_a = Mid(SCTS_a, 7, 2) & ":" & Mid(SCTS_a, 9, 2) & ":" & Mid(SCTS_a, 11, 2) 'hh:mm:dd
        
        
    End If
    
    strPesan = UD
    strNoHP = noOA
'    strTgl = SCTS_Tgl & " " & SCTS_Jam
'    mvarnoSCA = noSCA
'    mvarFO = FO
'    mvarDCS = DCS
'    mvarSCTS_Jam = SCTS_Jam
'    mvarSCTS_Tgl_A = SCTS_Tgl_a
'    mvarSCTS_Jam_A = SCTS_Jam_a
'    mvarUDL = UDL
Exit Function
pesan_error:
    MsgBox "Modul SMS_IN : " & Err.Number & Err.Description

End Function
Private Function Biner(Bilangan) As String
    Dim Basis As Integer
    Dim Hsltemp As Variant
    Dim sisa As Variant
    Dim HslBagi As Variant
    
    Hsltemp = ""
    sisa = ""
    Basis = 2
    
    Do
        Hsltemp = sisa & Hsltemp
        HslBagi = Bilangan \ Basis
        sisa = Bilangan Mod Basis
        Bilangan = HslBagi
    Loop Until HslBagi <= 1
    
    Biner = HslBagi & sisa & Hsltemp
    Biner = Right("0000000" & Biner, 7)
End Function

'Create 7 bit / 8 bit
'charhex(str,7)-->7 bit for receiving SMS
'charhex(str,8)-->8 bit for send
Public Function CharHex(ByVal Txt As String, ByVal bit As Integer)
    Dim i As Integer, bin As String, nbin As String, n As String
    Dim bil As Integer, sisa As Integer, lbin As Integer, nol As String
    bin = ""
    nbin = ""
 
    If bit = 7 Then
        For i = 1 To Len(Txt) Step 2
            n = Mid(Txt, i, 2)
            bin = HexToBin(n) & bin
        Next
        bil = Len(bin) \ bit
        sisa = Len(bin) Mod bit
        For i = 1 To (Len(bin) - sisa) Step bit
           ' MsgBox Chr$(HexToDec(BinToHex(Mid(bin, i + Sisa, bit))))
            nbin = Chr$(HexToDec(BinToHex(Mid(bin, i + sisa, bit)))) & nbin
        Next
    Else
        For i = 1 To Len(Txt)
            n = Mid(Txt, i, 1)
            bin = Biner(Asc(n)) & bin
        Next
        sisa = Len(bin) Mod bit
    
        If sisa > 0 Then
            For i = 1 To bit - sisa
                nol = nol & "0"
            Next
        End If
    
        bin = nol & bin
        bil = Len(bin) \ bit
        For i = 1 To bil
            nbin = nbin & BinToHex(Mid(bin, Len(bin) + 1 - bit * i, bit))
        Next
    End If
    CharHex = nbin
End Function

Public Function BinToHex(ByVal Biner As String) As String
    Dim bin As String, n As String, nil As String, i As Integer
    bin = ""
    Biner = Right("00000000" & Biner, 8)
    For i = 1 To 2
        bin = Mid(Biner, Len(Biner) + 1 - 4 * i, 4)
        Select Case bin
            Case "0000": n = "0"
            Case "0001": n = "1"
            Case "0010": n = "2"
            Case "0011": n = "3"
            Case "0100": n = "4"
            Case "0101": n = "5"
            Case "0110": n = "6"
            Case "0111": n = "7"
            Case "1000": n = "8"
            Case "1001": n = "9"
            Case "1010": n = "A"
            Case "1011": n = "B"
            Case "1100": n = "C"
            Case "1101": n = "D"
            Case "1110": n = "E"
            Case "1111": n = "F"
        End Select
        nil = n & nil
    Next
    BinToHex = nil
End Function

Public Function HexToBin(ByVal Biner As String) As String
    Dim bin As String, n As String, nil As String, i As Integer
    bin = ""
    For i = 1 To Len(Biner)
        bin = Mid(Biner, i, 1)
        Select Case bin
            Case "0": n = "0000"
            Case "1": n = "0001"
            Case "2": n = "0010"
            Case "3": n = "0011"
            Case "4": n = "0100"
            Case "5": n = "0101"
            Case "6": n = "0110"
            Case "7": n = "0111"
            Case "8": n = "1000"
            Case "9": n = "1001"
            Case "A": n = "1010"
            Case "B": n = "1011"
            Case "C": n = "1100"
            Case "D": n = "1101"
            Case "E": n = "1110"
            Case "F": n = "1111"
        End Select
        nil = nil & n
    Next
    HexToBin = nil
End Function

Public Function ConvToChar(ByVal hx As String) As String
    Dim i As Integer, tx As String
    For i = 1 To Len(hx) Step 2
        tx = tx & Chr(HexToDec(Mid(hx, i, 2)))
    Next
    ConvToChar = tx
End Function

Public Function HexToDec(ByVal x As String) As Integer
    Dim m As String, i As Byte, nil As Integer, n As Integer
    For i = 1 To 2
        m = Mid(x, i, 1)
        Select Case UCase(m)
            Case "A": n = 10
            Case "B": n = 11
            Case "C": n = 12
            Case "D": n = 13
            Case "E": n = 14
            Case "F": n = 15
            Case Else: n = CInt(m)
        End Select
        If i = 1 Then
            nil = n * 16
        Else
            nil = nil + n
        End If
    Next
    HexToDec = nil
End Function

Public Function DecToHex(ByVal x As Integer) As String
    Dim nil As String
    nil = Hex(x)
    If Len(nil) = 1 Then
        nil = "0" & nil
    End If
    DecToHex = nil
End Function

'Reverse number
Public Function RevNum(ByVal numb As String) As String
    Dim s As Integer, ma As String, b As String, a As String
    Dim ta As String
     s = 1
     ma = ""
     While (s <= Len(numb))
       ta = Mid(numb, s, 2)
       a = Mid(ta, 1, 1)
       b = Mid(ta, 2, 1)
       If b = "" Then b = "F"
       ma = ma & b & a
       s = s + 2
     Wend
     RevNum = ma
End Function

Public Function kirimPesan(ByVal DestinationNo As String, ByVal Message As String)
On Error Resume Next
    Dim SCA As String, PDU As String, MR As String
    Dim DA As String, PID As String, DCS As String
    Dim VP As String, UDL As String, UD As String
    
    
    SCA = "00"
    PDU = mvarMsgReport 'unreceived:"11"/received:"31")
    If PDU = "" Then PDU = "11" 'default:unreceived
    MR = "00"
    
    'DA: Destination Address
    
    DA = DecToHex(Len(DestinationNo)) 'Panjang DestinationNo dlm Hex
    If Left(DestinationNo, 2) = "62" Then
        DA = DA & "91" '"91":Int. Number(62...),"81":Loc. Number(081..)
    Else
        DA = DA & "81"
    End If
    DA = DA & RevNum(DestinationNo)
    
    PID = "00"
    DCS = mvarMsgType 'Normal:"00",Flash:"F0"
    If DCS = "" Then DCS = "00" 'default normal
    
    VP = mvarMsgTime 'Limit Period of delivery
    If VP = "" Then VP = "A7" ' default:1 days
    
    UDL = DecToHex(Len(Message)) ' length of message in Hex
    UD = CharHex(Message, 8) 'Message in Hex 8bit /octet
    
    'Format of SMS Submit PDU
    kirimPesan = SCA & PDU & MR & DA & PID & DCS & VP & UDL & UD

End Function
Public Function SMSDibalas(ByVal nomorhp As String, ByVal stdata As String)
On Error GoTo err_handler
hp_konek = False: err_konek = False
With utama
    .com.Output = "AT+CMGS=" & Len(kirimPesan(nomorhp, stdata)) / 2 - 1 & vbCrLf
    Tunda 0.1
    .com.Output = kirimPesan(nomorhp, stdata) & Chr(26)
End With
Exit Function
err_handler:
    MsgBox Err.Number & Err.Description
End Function


Public Function ambilpesan()
Dim x As Integer
On Error GoTo errHandler


If status_konek = True Then
    hp_konek = False: err_konek = False
    ReDim arrDel(1)
    arrDel(1) = 0
    utama.com.Output = "AT+CMGL=4" + vbCrLf
    Do While Not hp_konek
        DoEvents
        Tunggu
        If err_konek Then Exit Do
    Loop
x = 1
 Do While x <= UBound(arrDel)
    hp_konek = False: err_konek = False
    If Val(arrDel(x)) > 0 Then
            utama.com.Output = "AT+CMGD=" & Val(arrDel(x)) & vbCrLf
            Do While Not hp_konek
                DoEvents
                Tunggu
                If err_konek Then Exit Do
            Loop
    End If
    x = x + 1
Loop
Else
    status_konek = False
    MsgBox "Tidak ada Koneksi", vbInformation + vbSystemModal, "Gagal koneksi"
End If
If frmCommand.Listcommand.ListCount > 100 Then
        frmCommand.Listcommand.Clear
End If
'Timer1.Enabled = False
Exit Function
errHandler:
    MsgBox Err.Number & Err.Description, vbInformation, "KESALAHAN"
End Function

Public Function SMSReply()
Dim clsSMS As New Class_sms
Dim clsPlg As New ClassPelanggan
Dim clsrute As New ClassRoute
Dim clsKelas As New ClassKelas
Dim clsmaskapai As New ClassMaskapai
Dim listDat() As String
Dim xListDat() As String, bagiXDat() As String
Dim msgTemp As String
Dim i As Integer, x As Integer
'=======Var tarif
Dim tarifFuel() As String
On Error GoTo pesan_error
hp_konek = False: err_konek = False
    Do While Not hp_konek
        DoEvents
        If err_konek Then Exit Do
        Tunggu
    Loop
        listDat = Split(strPesan, " ")
        
        
        
 If UBound(listDat) >= 0 Then
 
      Select Case UCase(listDat(0))
           Case Is = "REG" 'Reg<spasi>[Nama]<spasi>[Alamat]
                    If UBound(listDat) = 2 Then
                       clsPlg.simpan_pelanggan strNoHP, listDat(1), listDat(2)
                    Else
                        GoTo Err_Msg
                    End If
           Case Is = "INFO" 'Info
                    
                    Select Case UCase(listDat(1))
                            Case Is = "RUTE" 'Info rute======================1
                                    If UBound(listDat) = 1 Then
                                        xListDat = clsrute.ambilRoute
                                        For i = 0 To UBound(xListDat)
                                                bagiXDat = Split(xListDat(i))
                                                msgTemp = msgTemp & bagiXDat(0) & ","
                                        Next
                                        balasanPesan = "Tujuan yg ada : " & msgTemp
                                    Else
                                        GoTo smsFormatSalah
                                    End If

                                    
                            Case Is = "JADWAL" 'Info<spasi>Jadwal<spasi>[Nama Tujuan]====================3
                                    If UBound(listDat) = 2 Then
                                        clsmaskapai.cek_jadwal listDat(2)
                                    Else
                                        GoTo smsFormatSalah
                                    End If
                            Case Is = "BOOKING" '================================4
                                    If UBound(listDat) = 1 Then
                                        balasanPesan = "Ketik Reg<spasi>(nama anda)<spasi>(Alamat)"
                                    Else
                                        GoTo smsFormatSalah
                                    End If

                            Case Is = "FLT" 'Info flt [tujuan]=======
                                    If UBound(listDat) = 2 Then
                                        clsrute.cari_noFlt listDat(2) 'Info flt [nmroute]
                                    Else
                                        GoTo smsFormatSalah
                                    End If
                            Case Else
                                GoTo smsFormatSalah
                    End Select
            Case Is = "TARIF" 'Tarif<spasi>[No.Flt]
                    If UBound(listDat) = 1 Then
                        clsKelas.cek_tarif listDat(1)
                    Else
                        GoTo smsFormatSalah
                    End If
            
            Case Is = "BOOKING"
                         
                        If UBound(listDat) = 4 Then
                            clsmaskapai.pesan_tiket strNoHP, listDat(1), listDat(2), listDat(3), listDat(4)
'
                        ElseIf UBound(listDat) = 5 Then  'Booking<spasi>[No.Flt]<spasi>[jml dewasa]<spasi>[Clas]<spasi>[tglBerangkat]

                             clsmaskapai.pesan_tiket_plus_anak strNoHP, listDat(1), listDat(2), listDat(3), listDat(4), listDat(5)
                                
'                        Else
'                            GoTo smsFormatSalah
                        End If
            Case Is = "CONFIRM" 'Confirm<spasi>[Kd pesan]<spasi>[no.rekening]<spasi>[Nm Bank]
                        If UBound(listDat) = 3 Then
                                clsSMS.confirm_pesan listDat(1), listDat(2), listDat(3)
                        Else
                            GoTo smsFormatSalah
                        End If
            Case Is = "UBAH" 'Ubah<spasi>[No.Booking]<spasi>[Dewasa]<spasi>[Bayi]
                        If UBound(listDat) = 4 Then
                            clsmaskapai.edit_tiket strNoHP, listDat(1), listDat(2), listDat(3), listDat(4)
                        Else
                            GoTo smsFormatSalah
                        End If
                        
            Case Is = "STATUS" 'Status<spasi>[no.pesan]
                        If UBound(listDat) = 1 Then
                                clsSMS.cek_status listDat(1)
                        Else
                            GoTo smsFormatSalah
                        End If
            Case Is = "BATAL" 'Batal<spasi>[no.pesan]
                        clsSMS.batal_pesan listDat(1)
            Case Else
                    balasanPesan = "Ketik : Info<spasi>(pesan,tarif,jadwal,rute,fasilitas,flt) - utk informasi pemesanan tiket"
            End Select
      

SMSDibalas strNoHP, balasanPesan
strNoHP = "": balasanPesan = ""

End If
Exit Function
Err_Msg:
    balasanPesan = "Format SMS terlalu panjang"
    SMSDibalas strNoHP, balasanPesan
    strNoHP = "": balasanPesan = ""
Exit Function
smsFormatSalah:
        balasanPesan = "Format SMS anda Salah"
        SMSDibalas strNoHP, balasanPesan
        strNoHP = "": balasanPesan = ""
Exit Function
pesan_error:
    MsgBox Err.Description & Err.Number
End Function
Function translate()
frmCommand.ListTrans.AddItem (strPesan & "           " & "(Dari: '" & strNoHP & "')")
End Function
Public Function Tunggu()
    Dim Start
    Start = Timer
    Do While Timer < Start + 2
       DoEvents
       If hp_konek Or err_konek Then Exit Do
    Loop
End Function
Sub Tunda(ByVal dtk As Single)
    Dim awal As Variant
    awal = Timer
    Do While Timer < awal + dtk
      DoEvents
    Loop
End Sub
Public Function bacaMerkHp()
hp_konek = False: err_konek = False
With utama.com
    .Output = "AT+CGMI" + vbCrLf
        Do While Not hp_konek
                DoEvents
                If err_konek Then Exit Do
                Tunggu
        Loop
        utama.sbar.Panels(1).Text = "Merk :" & "     " & frmCommand.Listcommand.List(frmCommand.Listcommand.ListCount - 1)
End With
End Function
Public Function bacatipe()
Dim tipe As String
hp_konek = False: err_konek = False
With utama
    .com.Output = "AT+CGMM" + vbCrLf
    .sbar.Panels(2).Text = .sbar.Panels(2).Text & "     " & tipe
    Do While Not hp_konek
        If err_konek Then Exit Do
           DoEvents
    Loop
     tipe = frmCommand.Listcommand.List(frmCommand.Listcommand.ListCount - 1)
End With
End Function
Public Function provid()
With utama
    hp_konek = False: err_konek = False
    .com.Output = "AT+COPS?" + vbCrLf
    Do While Not hp_konek
        If err_konek Then Exit Do
           DoEvents
    Loop
    .sbar.Panels(3).Text = Mid(frmCommand.Listcommand.List(frmCommand.Listcommand.ListCount - 1), 12)
End With
    
End Function
Public Function cek_string(isi As String)
cek_string = Replace(isi, "'", "`")
End Function
Sub bagiKirim(strHasil As String)
    
    
    
    ' membagi jumlah kalimat yang akan dikirim per 160 karakter
    jmlAwal = 0
    jmlAkhir = 0
    intUrut = 1
    If Len(strHasil) > 160 Then
        jmlAwal = 1
        jmlAkhir = fcCariJmlAkhir(strHasil, 160)
        arrKirim(intUrut) = Trim(Mid(strHasil, jmlAwal, jmlAkhir))
        
        jmlAwal = jmlAkhir + 1
        If Len(strHasil) - jmlAwal > 160 Then
            jmlAkhir = jmlAwal + 160
        Else: jmlAkhir = Len(strHasil) + 1
        End If
            
        While Len(strHasil) > jmlAwal
            intUrut = intUrut + 1
            If Len(strHasil) - jmlAwal > 160 Then jmlAkhir = fcCariJmlAkhir(strHasil, jmlAkhir)
            
            intTambah = UBound(arrKirim) + 1
            ReDim Preserve arrKirim(intTambah)
            arrKirim(intUrut) = Trim(Mid(strHasil, jmlAwal, jmlAkhir - jmlAwal))
            
            jmlAwal = jmlAkhir + 1
            If Len(strHasil) - jmlAwal > 160 Then
                jmlAkhir = jmlAwal + 160
            Else: jmlAkhir = Len(strHasil) + 1
            End If
        Wend
    Else
        arrKirim(intUrut) = strHasil
    End If
End Sub

Function fcCariJmlAkhir(strHasilTmp As String, nilAkhir As Integer) As Integer
    Dim strSpasi As String
    Dim intX As Integer
    
    strSpasi = Mid(strHasilTmp, jmlAwal, nilAkhir + 1)
    
    intX = 1
    If Mid(strHasilTmp, jmlAwal, nilAkhir + intX) <> "" Then
        Do While strSpasi <> ""
            If Mid(strHasilTmp, nilAkhir - intX, 1) = " " Then
                fcCariJmlAkhir = nilAkhir - intX
                Exit Do
            End If
            strSpasi = Mid(strHasilTmp, nilAkhir - intX, 1)
            intX = intX + 1
        Loop
    Else
        fcCariJmlAkhir = nilAkhir
    End If
End Function
Public Function cekpesan(strEvent As String)
 
    Dim bag() As String
    Dim bg2() As String
    Dim msgTmp() As String
    Dim ceks As String
    
'    On Error Resume Next
    If strEvent <> "AT+CMGL=4" And strEvent <> "OK" Then frmCommand.Listcommand.AddItem (strEvent)
    If Mid$(strEvent, 1, 5) = "+CMGL" Then ' List to Message
        msgTmp = Split(strEvent, ",", , vbTextCompare)
        arrDel(UBound(arrDel)) = Val(Trim(Right(msgTmp(0), 2)))
        ReDim Preserve arrDel(UBound(arrDel) + 1)
    End If
    
    If Mid$(frmCommand.Listcommand.List(frmCommand.Listcommand.ListCount - 2), 1, 5) = "+CMGL" And strEvent <> "OK" Then
        strPesan = "": strNoHP = ""
        status_signal = False
        utama.Timer2 = False
        SMSIn CStr(frmCommand.Listcommand.List(frmCommand.Listcommand.ListCount - 1))
        ceks = cek_string(strPesan)
        translate
        Call SMSReply
        utama.Timer2 = True
        status_signal = True
        
    Else
        If Mid$(frmCommand.Listcommand.List(frmCommand.Listcommand.ListCount - 2), 1, 5) = "+CSQ:" Then
            status_konek = False
            utama.Timer1 = False
            bag() = Split(Mid(frmCommand.Listcommand.List(frmCommand.Listcommand.ListCount - 2), 1, 11), " ")
            bg2() = Split(bag(1), ",")
'
            utama.pbsignal.Value = bg2(0)
            If frmCommand.Listcommand.ListCount > 50 Then frmCommand.Listcommand.Clear
        Else
            status_konek = True
            utama.Timer1 = True
        End If
    End If
     Select Case strEvent
        Case "OK": hp_konek = True
        Case "ERROR": err_konek = True
        Case Else
            If InStr(1, strEvent, "ERROR") Then: err_konek = True
    End Select
End Function

