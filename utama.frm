VERSION 5.00
Object = "{38911DA0-E448-11D0-84A3-00DD01104159}#1.1#0"; "COMCT332.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{648A5603-2C6E-101B-82B6-000000000014}#1.1#0"; "MSCOMM32.OCX"
Object = "{8E37F849-94CB-11D5-B563-000021FCAE1A}#1.0#0"; "mdpBar.ocx"
Begin VB.MDIForm utama 
   AutoShowChildren=   0   'False
   BackColor       =   &H00FF8080&
   Caption         =   "Mandala Airlines ( SMS Explorer )"
   ClientHeight    =   7305
   ClientLeft      =   165
   ClientTop       =   735
   ClientWidth     =   11145
   Icon            =   "utama.frx":0000
   LinkTopic       =   "MDIForm1"
   LockControls    =   -1  'True
   Picture         =   "utama.frx":1CBF2
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   Begin VB.Timer Timer2 
      Interval        =   15000
      Left            =   2040
      Top             =   6240
   End
   Begin MSComctlLib.ImageList ImageList3 
      Left            =   360
      Top             =   5400
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   10
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "utama.frx":2AF98
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "utama.frx":47B9A
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "utama.frx":65134
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "utama.frx":651DB
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "utama.frx":65282
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "utama.frx":65591
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "utama.frx":65A6D
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "utama.frx":67E4F
            Key             =   ""
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "utama.frx":73411
            Key             =   ""
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "utama.frx":7372B
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin ComCtl3.CoolBar CoolBar1 
      Align           =   1  'Align Top
      Height          =   1215
      Left            =   0
      TabIndex        =   1
      Top             =   0
      Width           =   11145
      _ExtentX        =   19659
      _ExtentY        =   2143
      _CBWidth        =   11145
      _CBHeight       =   1215
      _Version        =   "6.0.8169"
      Caption1        =   "Icon Menus :"
      Child1          =   "tbar2"
      MinHeight1      =   390
      Width1          =   1935
      NewRow1         =   0   'False
      Caption2        =   "Port :"
      Child2          =   "Frame1"
      MinHeight2      =   735
      Width2          =   6690
      NewRow2         =   -1  'True
      Child3          =   "Frame2"
      MinHeight3      =   615
      Width3          =   9390
      NewRow3         =   0   'False
      BandTag3        =   "0"
      Begin VB.Frame Frame2 
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   0  'None
         Height          =   615
         Left            =   6885
         TabIndex        =   6
         Top             =   510
         Width           =   4170
         Begin VB.Label Label1 
            BackColor       =   &H00E0E0E0&
            Caption         =   "Mandala Airlines"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   18
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FFC0C0&
            Height          =   375
            Left            =   120
            TabIndex        =   7
            Top             =   120
            Width           =   3855
         End
      End
      Begin MSComctlLib.Toolbar tbar2 
         Height          =   390
         Left            =   1155
         TabIndex        =   5
         Top             =   30
         Width           =   9900
         _ExtentX        =   17463
         _ExtentY        =   688
         ButtonWidth     =   609
         ButtonHeight    =   582
         Appearance      =   1
         ImageList       =   "ImageList3"
         DisabledImageList=   "ImageList3"
         HotImageList    =   "ImageList3"
         _Version        =   393216
         BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
            NumButtons      =   7
            BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Description     =   "File Master"
               Object.ToolTipText     =   "File Master"
               ImageIndex      =   3
            EndProperty
            BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Description     =   "Check Pemesanan"
               Object.ToolTipText     =   "Check Pemesanan"
               ImageIndex      =   4
            EndProperty
            BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Description     =   "Broadcast"
               Object.ToolTipText     =   "Sms Broadcast"
               ImageIndex      =   1
            EndProperty
            BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Description     =   "Membuat password"
               Object.ToolTipText     =   "Membuat password"
               ImageIndex      =   7
            EndProperty
            BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Description     =   "Ganti password"
               Object.ToolTipText     =   "Ganti password"
               ImageIndex      =   5
            EndProperty
            BeginProperty Button6 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Description     =   "Database utility"
               Object.ToolTipText     =   "Database utility"
               ImageIndex      =   9
            EndProperty
            BeginProperty Button7 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Object.ToolTipText     =   "Command Report"
               ImageIndex      =   10
            EndProperty
         EndProperty
         Begin mdBar.mdpBar pbsignal 
            Height          =   255
            Left            =   8520
            TabIndex        =   8
            Top             =   120
            Width           =   1095
            _ExtentX        =   1931
            _ExtentY        =   450
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Max             =   30
            Value           =   1
            CaptionAlingment=   2
            Frame           =   -1  'True
            Animate         =   2
         End
         Begin VB.Frame Frame3 
            BorderStyle     =   0  'None
            Height          =   495
            Left            =   7440
            TabIndex        =   9
            Top             =   0
            Width           =   1455
            Begin VB.Label Label2 
               Caption         =   "Status Sinyal :"
               Height          =   255
               Left            =   0
               TabIndex        =   10
               Top             =   120
               Width           =   1095
            End
         End
      End
      Begin VB.Frame Frame1 
         BorderStyle     =   0  'None
         Height          =   735
         Left            =   630
         TabIndex        =   2
         Top             =   450
         Width           =   6030
         Begin MSComctlLib.Toolbar Tbar1 
            Height          =   600
            Left            =   1200
            TabIndex        =   4
            Top             =   0
            Width           =   3810
            _ExtentX        =   6720
            _ExtentY        =   1058
            ButtonWidth     =   1244
            ButtonHeight    =   1058
            Wrappable       =   0   'False
            Style           =   1
            ImageList       =   "ImageList1"
            DisabledImageList=   "ImageList1"
            HotImageList    =   "ImageList2"
            _Version        =   393216
            BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
               NumButtons      =   3
               BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
                  Caption         =   "Connect"
                  Description     =   "Connect"
                  Object.ToolTipText     =   "Connect Hp"
                  Object.Tag             =   "Connect"
                  ImageIndex      =   8
                  BeginProperty ButtonMenus {66833FEC-8583-11D1-B16A-00C0F0283628} 
                     NumButtonMenus  =   1
                     BeginProperty ButtonMenu1 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                        Text            =   "test"
                     EndProperty
                  EndProperty
               EndProperty
               BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
                  Caption         =   "Stop"
                  Description     =   "Stop"
                  Object.ToolTipText     =   "Disconnect"
                  ImageIndex      =   3
               EndProperty
               BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
                  Caption         =   "Refresh"
                  Description     =   "Refresh"
                  Object.ToolTipText     =   "Reconnect"
                  ImageIndex      =   4
               EndProperty
            EndProperty
         End
         Begin VB.ComboBox cmbport 
            Height          =   315
            ItemData        =   "utama.frx":7EB65
            Left            =   50
            List            =   "utama.frx":7EB67
            TabIndex        =   3
            Text            =   "COM8"
            Top             =   120
            Width           =   1095
         End
      End
   End
   Begin VB.Timer Timer1 
      Interval        =   1000
      Left            =   1560
      Top             =   6240
   End
   Begin MSComctlLib.StatusBar sbar 
      Align           =   2  'Align Bottom
      Height          =   255
      Left            =   0
      TabIndex        =   0
      Top             =   7050
      Width           =   11145
      _ExtentX        =   19659
      _ExtentY        =   450
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   4
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
         EndProperty
         BeginProperty Panel3 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Object.Width           =   6174
            MinWidth        =   6174
         EndProperty
         BeginProperty Panel4 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
         EndProperty
      EndProperty
   End
   Begin MSCommLib.MSComm com 
      Left            =   960
      Top             =   6120
      _ExtentX        =   1005
      _ExtentY        =   1005
      _Version        =   393216
      CommPort        =   8
      DTREnable       =   -1  'True
      BaudRate        =   19200
   End
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   960
      Top             =   5400
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   20
      ImageHeight     =   20
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   8
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "utama.frx":7EB69
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "utama.frx":7F0C5
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "utama.frx":7F621
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "utama.frx":7FB7D
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "utama.frx":800D9
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "utama.frx":80635
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "utama.frx":80B91
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "utama.frx":810ED
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ImageList ImageList2 
      Left            =   360
      Top             =   6120
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   20
      ImageHeight     =   20
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   8
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "utama.frx":81407
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "utama.frx":81963
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "utama.frx":81EBF
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "utama.frx":8241B
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "utama.frx":82977
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "utama.frx":82ED3
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "utama.frx":8342F
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "utama.frx":8398B
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin VB.Menu mnFile 
      Caption         =   "&File"
      Begin VB.Menu mnMaster 
         Caption         =   "Input File Master"
      End
      Begin VB.Menu mnKeluar 
         Caption         =   "Keluar"
      End
   End
   Begin VB.Menu mnProses 
      Caption         =   "&Proses"
      Begin VB.Menu mnambil 
         Caption         =   "Check Pemesanan"
      End
   End
   Begin VB.Menu mnBroadcast 
      Caption         =   "&Broadcast"
   End
   Begin VB.Menu mnSet 
      Caption         =   "&Settings"
      Begin VB.Menu mnPs 
         Caption         =   "Password"
         Begin VB.Menu mnPsbaru 
            Caption         =   "Input Password"
         End
         Begin VB.Menu mnGpas 
            Caption         =   "Ganti Password"
         End
      End
      Begin VB.Menu mnDbSet 
         Caption         =   "Database Setup"
      End
   End
   Begin VB.Menu mnHelp 
      Caption         =   "&Help"
      Begin VB.Menu mnManual 
         Caption         =   "Manual"
      End
      Begin VB.Menu mnAbout 
         Caption         =   "About"
      End
   End
End
Attribute VB_Name = "utama"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Sub com_OnComm()
Static stEvent As String
Dim stComChar As String * 1
    Select Case com.CommEvent
        Case comEvReceive
            Do
                stComChar = com.Input
                Select Case stComChar
                    Case ">"
                        frmCommand.Listcommand.AddItem (stComChar)
                    Case vbLf
                    Case vbCr
                        If Len(stEvent) > 0 Then
                          cekpesan stEvent
                          stEvent = ""
                        End If
                    Case Else
                        stEvent = stEvent + stComChar
                End Select
            Loop While com.InBufferCount
     End Select
End Sub


Private Sub MDIForm_Load()
Dim i As Integer

cmbport.Clear
For i = 1 To 20
    cmbport.AddItem ("COM" & i)
Next
'pbsignal.Max = pbsignal.Value
End Sub

Private Sub MDIForm_QueryUnload(Cancel As Integer, UnloadMode As Integer)
Timer1 = False
Timer2 = False
status_signal = False
status_konek = False

End Sub

Private Sub MDIForm_Resize()
pbsignal.Left = (Me.Width - 2680)
Frame3.Left = Me.Width - 3800
End Sub

Private Sub mnAbout_Click()
frmAbout.Show
End Sub

Private Sub mnambil_Click()
frmDelay.Show 1
frmambil.Show
End Sub

Private Sub mnBroadcast_Click()
frmDelay.Show 1
frmbroadcast.Show
End Sub

Private Sub mnDbSet_Click()
frmDbase.Show
End Sub

Private Sub mnGpas_Click()
frmUbahPass.Show
End Sub

Private Sub mnKeluar_Click()
frmLogout.Show
End Sub

Private Sub mnManual_Click()
App.HelpFile = App.Path & "\Manual\manual2.chm"
Call HtmlHelp(0, App.HelpFile, HH_DISPLAY_TOPIC, ByVal 0&)
End Sub
Function WaitForProcess(taskId As Long, Optional msecs As Long = -1) _
    As Boolean
    Dim procHandle As Long
    ' Get the process handle.
    procHandle = OpenProcess(&H100000, True, taskId)
    ' Check for its signaled status; return to caller.
    WaitForProcess = WaitForSingleObject(procHandle, msecs) <> -1
    ' Close the handle.
    CloseHandle procHandle
End Function
'Private Sub mnKoneksi_Click()
'frmDelay.Show 1
'frmKoneksi.Show
'End Sub

Private Sub mnMaster_Click()
frmDelay.Show 1
masterKelas.Show

End Sub

Private Sub mnPsbaru_Click()
frmPass.Show
End Sub

Private Sub Tbar1_ButtonClick(ByVal Button As MSComctlLib.Button)
Select Case Button.Index
    Case 1
     If Len(cmbport.Text) = 0 Then
            MsgBox "Pilih Port koneksi", vbInformation + vbSystemModal, "Informasi"
     Else
        mulai
     End If
    Case 2
        DC
        Tbar1.Buttons(1).Enabled = True
         pbsignal.Value = 1
        
        Unload frmCommand
    Case 3
        If utama.com.PortOpen = True Then
            DC
            frmCommand.Listcommand.Clear
            mulai
        
              
        End If
        
        
End Select
End Sub

Private Sub tbar2_ButtonClick(ByVal Button As MSComctlLib.Button)
frmDelay.Show 1
Select Case Button.Index
        Case 1
            masterKelas.Show
        Case 2
            frmambil.Show
        Case 3
            frmbroadcast.Show
        Case 4
            frmPass.Show
        Case 5
            frmUbahPass.Show
        Case 6
            frmDbase.Show
         Case 7
            Me.AutoShowChildren = True
            frmCommand.Show
End Select
End Sub

Private Sub Timer1_Timer()
If status_konek = True Then
    Call ambilpesan
End If
End Sub
 
Private Sub Toolbar2_ButtonClick(ByVal Button As MSComctlLib.Button)

End Sub
Function mulai()
Dim s As String
Dim nocomport As String
On Error GoTo err_handler
nocomport = Mid(cmbport.Text, 4)
hp_konek = False: err_konek = False

With utama.com

    If .PortOpen = False Then
         utama.sbar.Panels(4).Text = "Tunggu..."
        .CommPort = nocomport
        .Settings = "19200,n,8,1"
        .DTREnable = True
        .RTSEnable = True
        .RThreshold = 1
        .InputLen = 1
        .PortOpen = True
        bacaMerkHp
        bacatipe
        provid
        status_konek = True
        status_signal = True
        
'        utama.Timer1 = True
'        utama.Timer2 = True

        Tbar1.Buttons(1).Enabled = False
        
        MsgBox "koneksi berhasil"
    Else
        .PortOpen = False
        status_konek = False
'        utama.Timer1 = False
        utama.sbar.Panels(4).Text = ""
        MsgBox "Tidak ada koneksi"
    End If
    utama.sbar.Panels(4).Text = "Status: Connect"
End With

Exit Function
err_handler:
    MsgBox Err.Number & Err.Description
End Function
Function DC()
With utama
If .com.PortOpen = True Then
    .Timer1 = False
     status_konek = False
     status_signal = False
    .com.PortOpen = False
    .sbar.Panels(4).Text = "Disconnect"
    .sbar.Panels(3).Text = ""
    .sbar.Panels(2).Text = ""
    .sbar.Panels(1).Text = ""
End If
End With
End Function

Private Sub Timer2_Timer()
If status_signal = True Then
    SignalStatus
Else
    status_signal = False
End If
End Sub
Public Function SignalStatus()
hp_konek = False: err_konek = False
utama.com.Output = "AT+CSQ" + Chr(13)
End Function

