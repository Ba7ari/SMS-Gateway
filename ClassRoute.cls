VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ClassRoute"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Public Function simpan_route(ByVal kdRoute As String, ByVal nmRoute As String)
Set getData.rsRoute = New ADODB.Recordset
getData.rsRoute.Open "SELECT * FROM Route where kd_route Like '" & kdRoute & "'", cn, 1, 3
With getData.rsRoute
    If .RecordCount <> 0 Then
            MsgBox "Kode route '" & kdRoute & "'sudah pernah diinput", vbInformation + vbSystemModal, "Informasi"
    Else
        .AddNew
                !Kd_route = kdRoute
                !Nm_route = nmRoute
        .Update
    End If
End With
getData.rsRoute.Close
cn.Close
End Function
Public Function hapus_Route(Optional ByVal kdRoute As String)
Set getData.rsRoute = New ADODB.Recordset
getData.rsRoute.Open "SELECT * FROM Route WHERE kd_route LIKE '" & kdRoute & "'", cn, 1, 3
With getData.rsRoute
    If .RecordCount <> 0 Then
    Do While Not .EOF
            .Delete
            .MoveNext
    Loop
    End If
End With
getData.rsRoute.Close
cn.Close
End Function
Public Function ambilRoute() As String()
Dim i As Integer
Dim sRoute() As String
Set getData.rsRoute = New ADODB.Recordset
getData.rsRoute.Open "SELECT distinct * FROM Route", cn, 1, 3
With getData.rsRoute
If .RecordCount > 0 Then
   Do While Not .EOF
        ReDim Preserve sRoute(i)
        sRoute(i) = !Kd_route & "#" & !Nm_route
        i = i + 1
        .MoveNext
    Loop
End If
End With
getData.rsRoute.Close
cn.Close
ambilRoute = sRoute
End Function
Public Function EditRoute(ByVal kdRoute As String, ByVal nmRoute As String)
Set getData.rsRoute = New ADODB.Recordset
getData.rsRoute.Open "SELECT * FROM Route WHERE Kd_Route LIKE '" & kdRoute & "'", cn, 1, 3
        With getData.rsRoute
           If .RecordCount = 0 Then
                MsgBox "Tidak dapat melakukan perubahan"
            Else
            !Kd_route = kdRoute
            !Nm_route = nmRoute
            .Update
           End If
        End With
getData.rsRoute.Close
cn.Close
End Function
Public Function CariRoute(ByVal kdRoute As String) As String()
Dim i As Integer
Dim sRoute() As String
Set getData.rsRoute = New ADODB.Recordset
getData.rsRoute.Open "SELECT * FROM Route WHERE Kd_Route LIKE '" & kdRoute & "'", cn, 1, 3
With getData.rsRoute
If .RecordCount > 0 Then
   Do While Not .EOF
        ReDim Preserve sRoute(i)
        sRoute(i) = !Kd_route & " " & !Nm_route
        i = i + 1
        .MoveNext
    Loop
End If
End With
getData.rsRoute.Close
cn.Close
CariRoute = sRoute
End Function

Public Function cari_noFlt(ByVal rute As String)
Dim i As Integer
Dim xData() As String
Dim Bagi() As String
Dim hasil As String
hp_konek = False: err_konek = False
Set getData.rsRoute = New ADODB.Recordset
getData.rsRoute.Open "SELECT Maskapai.No_flt,jadwal,Route.Kd_route,Nm_route FROM Maskapai,Route WHERE Maskapai.Kd_Route = Route.Kd_Route and Route.Kd_Route LIKE '" & rute & "'", cn, 1, 3
With getData.rsRoute
   If .RecordCount = 0 Then
        balasanPesan = "Maaf,Rute yang anda cari belum ada pada maskapai kami"
   Else
     Do While Not .EOF
            hasil = hasil & " No.flt " & !no_flt & " Brkt Pkl :" & Format(!jadwal, "hh:MM AMPM") & " "
    .MoveNext
     Loop
     balasanPesan = hasil
     Tunggu
   End If
End With
getData.rsRoute.Close
cn.Close
'SMSDibalas strNoHP, balasanPesan
'strNoHP = "": balasanPesan = ""
End Function

Public Function hapus_semuaRute()
Set getData.rsRoute = New ADODB.Recordset
getData.rsRoute.Open "SELECT * FROM Route", cn, 1, 3
With getData.rsRoute
    Do While Not .EOF
            .Delete
            .MoveNext
    Loop
End With
getData.rsRoute.Close
cn.Close
End Function
Public Function cari_Flight(ByVal no_flt As String)
Dim nama_route_ As String
Set getData.rsMaskapai = New ADODB.Recordset
getData.rsMaskapai.Open "SELECT Nm_Route FROM Maskapai,Route WHERE Maskapai.Kd_route = Route.Kd_route and Maskapai.No_flt LIKE '" & no_flt & "'", cn, 1, 3
With getData.rsMaskapai
        If .RecordCount = 0 Then
                MsgBox "No.flt tidak terdaftar", vbInformation + vbSystemModal, "Informasi"
        Else
            nama_route_ = !Nm_route
        End If
        
End With
getData.rsMaskapai.Close
cn.Close
cari_Flight = nama_route_
End Function
Public Function cari_jadwal(ByVal kdTujuan As String) As String()
Dim i As Integer
Dim xData() As String

Set getData.rsMaskapai = New ADODB.Recordset
getData.rsMaskapai.Open "SELECT distinct Maskapai.No_flt,Route.kd_route,Class.nm_Class,Maskapai.Jadwal,Class.SellingPrice,StockClass FROM Maskapai,Route,Class WHERE Maskapai.No_flt = Class.No_flt and Maskapai.Kd_route = Route.Kd_route and Route.Kd_route like '" & kdTujuan & "'", cn, 1, 3
With getData.rsMaskapai
    If .RecordCount > 0 Then
        Do While Not .EOF
                ReDim Preserve xData(i)
                xData(i) = !no_flt & "#" & !Kd_route & "#" & !nm_Class & "#" & Format(!jadwal, "HH:MM AMPM") & "#" & !sellingprice & "#" & !StockClass
                i = i + 1
            .MoveNext
       Loop
    End If
End With
getData.rsMaskapai.Close
cn.Close
cari_jadwal = xData
End Function
