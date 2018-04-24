Attribute VB_Name = "ckonek"
Option Explicit
Public Const HWND_TOPMOST = -1
Public Const HWND_NOTOPMOST = -2
Public Const SWP_NOMOVE = &H2
Public Const SWP_NOSIZE = &H1
Public Const SWP_NOACTIVATE = &H10
Public Const SWP_SHOWWINDOW = &H40
Public Const TOPMOST_FLAGS = SWP_NOMOVE Or SWP_NOSIZE

Public Declare Function SetWindowPos Lib "user32" _
(ByVal hwnd As Long, ByVal hWndInsertAfter As Long, _
    ByVal x As Long, ByVal y As Long, _
    ByVal cx As Long, ByVal cy As Long, _
    ByVal wFlags As Long) As Long
Public m_mode As Boolean
Public Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)
Public mulai As Boolean
Public Type dt 'ambil rekord
    rsRoute As ADODB.Recordset
    rsMaskapai As ADODB.Recordset
    rsClass As ADODB.Recordset
    rsPelanggan As ADODB.Recordset
    rsReservasi As ADODB.Recordset
    rsAda As ADODB.Recordset
    rsHitung As ADODB.Recordset
End Type
Public getData As dt
Public Function cn() As ADODB.Connection
Set cn = New ADODB.Connection
cn.Open "PROVIDER=MICROSOFT.Jet.oledb.4.0;data source = " & App.Path & "\Database\data.mdb"
End Function
Public Function centerscreen(ByVal frm As Form)
frm.Top = (utama.Height - frm.Height) / 2
frm.Left = (utama.Width - frm.Width) / 2
End Function

Public Function cekdigit(ByVal KeyAscii As Integer) As Boolean
If Not IsNumeric(Chr(KeyAscii)) = True Then
    If KeyAscii <> vbKeyBack Then
        If KeyAscii <> 13 Then
            cekdigit = True
        End If
    End If
End If
End Function
Function splashCenter(ByVal frm As Form)
frm.Top = (Screen.Height - frm.Height) / 2
frm.Left = (Screen.Width - frm.Width) / 2
End Function
Public Function cnAdmin() As ADODB.Connection
Set cnAdmin = New ADODB.Connection
cnAdmin.Open "PROVIDER=MICROSOFT.Jet.oledb.4.0;data source = " & App.Path & "\Database\datMin.mdb;Jet oledb:Database Password ='root'"
End Function
Public Function clearscreen(ByVal frm As Form)
Dim ctrl As Control
For Each ctrl In frm.Controls
    If TypeOf ctrl Is TextBox Then
            ctrl.Text = ""
    End If
Next
End Function
Sub main()
Dim cekSett As String
With frmMulai
    .Showsplash = True
    .Show
     
    .lblpsnloading.Caption = "Loading mohon tunggu..."
    .Refresh
.proses.Value = 10
    Call Sleep(1000)
.proses.Value = 25
    Call Sleep(1000)
.proses.Value = 35
    Call Sleep(1000)
.proses.Value = 45
    cekSett = GetSetting("MandalaGateway", "Pass", "Settings")
Call Sleep(1000)
   .lblpsnloading.Caption = "get setting mandala getway..."
    .Refresh
.proses.Value = 60
        Call Sleep(2000)
        .proses.Value = 85
        .lblpsnloading.Caption = "cek status setting login..."
         .Refresh
        If Len(cekSett) <> 0 Then
            
            .lblpsnloading.Caption = "Proses pengecekan Status login..."
            .Refresh
            
                If cekSett = "1" Then
                        Call Sleep(1000)
                        .proses.Value = 100
                        .lblpsnloading.Caption = "loading password..."
                        .Refresh
                        Call Sleep(1000)
                        Unload frmMulai
                        frmLOGIN.Show 1
                        
                Else
                    
                    .proses.Value = 100
                    .lblpsnloading.Caption = "Pengecekan sukses"
                    .Refresh
                    Call Sleep(1000)
                    
                    Unload frmMulai
                    utama.Show
                End If
        Else
                SaveSetting "MandalaGateway", "Pass", "Settings", "0" 'beri nilai default
        End If
.proses.Value = 100
End With
End Sub
Public Function tbox_kosong(ByVal frm As Form) As Boolean
Dim ctrl As Control
For Each ctrl In frm.Controls
If TypeOf ctrl Is TextBox Then
        If Len(ctrl) = 0 Then
            tbox_kosong = True
        End If
End If
Next
End Function
