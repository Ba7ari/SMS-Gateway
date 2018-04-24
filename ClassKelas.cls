VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ClassKelas"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Dim kdClass_ As String
Dim noFlt_ As String
Public Function simpan_kelas(ByVal nmKelas As String, ByVal fareBasic As Long, ByVal fuel As Long, ByVal ppn As Long, ByVal iwjr As Long, ByVal Sell As Long, ByVal stok As Long)
Set getData.rsClass = New ADODB.Recordset
getData.rsClass.Open "SELECT * FROM Class WHERE Kd_Class = '" & kdClass & "' and No_Flt LIKE '" & noflt & "'", cn, 1, 3
With getData.rsClass
    If .RecordCount <> 0 Then
            MsgBox "Kode Class '" & kdClass & "' No.Flt '" & noflt & "' sudah pernah diinput", vbInformation + vbSystemModal, "Informasi"
    Else
    .AddNew
            !Kd_Class = kdClass
            !nm_Class = nmKelas
            !no_flt = noflt
            !fareBasic = fareBasic
            !fuelSurc = fuel
            !iwjr = iwjr
            !ppn = ppn
            !sellingprice = Sell
            !StockClass = stok
    .Update
    End If
End With
getData.rsClass.Close
cn.Close
End Function
Public Function ClassDetail() As String()
Dim i As Integer
Dim sClass() As String
Dim tmpJadwal As Variant

Set getData.rsClass = New ADODB.Recordset
getData.rsClass.Open "SELECT distinct Maskapai.No_flt,Route.kd_route,Class.Nm_Class,Maskapai.Jadwal,Class.SellingPrice,StockClass FROM Maskapai,Route,Class WHERE Maskapai.No_flt = Class.No_flt and Maskapai.Kd_route = Route.Kd_route order by Route.Kd_route asc", cn, 1, 3
With getData.rsClass
        If .RecordCount > 0 Then
            Do While Not .EOF
                    ReDim Preserve sClass(i)
'                    tmpJadwal = FormatDateTime(!jadwal, vbShortTime)
                    sClass(i) = !no_flt & "#" & !Kd_route & "#" & !nm_Class & "#" & Format(!jadwal, "HH:MM AMPM") & "#" & !sellingprice & "#" & !StockClass
                    i = i + 1
                    .MoveNext
            Loop
        End If
End With
getData.rsClass.Close
cn.Close
ClassDetail = sClass
End Function
Public Function cek_tarif(ByVal noflt As String)
Dim msgTemp As String, nmClass As String
Dim clsTemp As String
hp_konek = False: err_konek = False
Set getData.rsClass = New ADODB.Recordset
getData.rsClass.Open "SELECT Class.Nm_Class,SellingPrice,Route.Nm_Route,Maskapai.No_flt FROM Class,Maskapai,Route WHERE Class.No_flt = Maskapai.No_flt and Maskapai.Kd_Route = Route.Kd_Route and Maskapai.No_Flt LIKE '" & noflt & "'", cn, 1, 3
With getData.rsClass
    If .RecordCount = 0 Then
            balasanPesan = "No.Flight tidak ada"
    Else
         msgTemp = "No.Flt: " & !no_flt
        Do While Not .EOF
            Select Case !nm_Class
            Case Is = "Flexi"
                clsTemp = "Flex"
            Case Is = "Saver"
                clsTemp = "Saver"
            Case Is = "Infant"
                clsTemp = "INF"
            End Select
            
            msgTemp = msgTemp & " Cls " & clsTemp & " " & "Hrg:" & Format(!sellingprice, "Rp###,###,###") & ","
            
            .MoveNext
        Loop
        balasanPesan = msgTemp & "INF(Bayi umr 2 thn kbwh)"
'        Do While hp_konek
'                DoEvents
                If err_konek = True Then MsgBox Err.Description
                Tunggu
'        Loop
                    
    End If
End With
getData.rsClass.Close
cn.Close
'MsgBox Len(msgTemp)

SMSDibalas strNoHP, balasanPesan
msgTemp = ""
balasanPesan = "": strNoHP = ""
    
End Function
Public Property Get kdClass() As String
kdClass = kdClass_
End Property

Public Property Let kdClass(ByVal vNewValue As String)
kdClass_ = vNewValue
End Property

Public Property Get noflt() As String
noflt = noFlt_
End Property

Public Property Let noflt(ByVal vNewValue As String)
noFlt_ = vNewValue
End Property

Public Function hapus_kelas()
Set getData.rsClass = New ADODB.Recordset
getData.rsClass.Open "SELECT * FROM Class", cn, 1, 3
With getData.rsClass
    Do While Not .EOF
            .Delete
            .MoveNext
    Loop
End With
getData.rsClass.Close
cn.Close
End Function
Public Function cek_tarif_flt(ByVal noflt As String)
Set getData.rsClass = New ADODB.Recordset
getData.rsClass.Open "SELECT SellingPrice From Class Where Kd_Class = 'N' and No_flt like '" & noflt & "'", cn, 1, 3
With getData.rsClass
    If .RecordCount = 0 Then
            balasanPesan = "No.Flt tidak ada"
    Else
        Do While Not .EOF
            balasanPesan = !sellingprice
            .MoveNext
        Loop
    End If
End With
getData.rsClass.Close
cn.Close
End Function
