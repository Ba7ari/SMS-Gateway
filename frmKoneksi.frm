VERSION 5.00
Object = "{531653E6-69A2-40A4-8734-91C78B33DD08}#1.0#0"; "XPButton.ocx"
Begin VB.Form frmKoneksi 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Koneksi hp"
   ClientHeight    =   1140
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   3255
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   1140
   ScaleWidth      =   3255
   ShowInTaskbar   =   0   'False
   Begin ucXPButton.XPButton cmdConnect 
      Height          =   375
      Left            =   1920
      TabIndex        =   2
      Top             =   120
      Width           =   1215
      _ExtentX        =   2143
      _ExtentY        =   661
      BTYPE           =   5
      TX              =   "Connect"
      ENAB            =   -1  'True
      BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      COLTYPE         =   1
      FOCUSR          =   -1  'True
      BCOL            =   13160660
      BCOLO           =   13160660
      FCOL            =   0
      FCOLO           =   0
      MCOL            =   12632256
      MPTR            =   1
      MICON           =   "frmKoneksi.frx":0000
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin VB.ComboBox cmbPort 
      Height          =   315
      Left            =   600
      TabIndex        =   1
      Top             =   120
      Width           =   1215
   End
   Begin ucXPButton.XPButton cmdDC 
      Height          =   375
      Left            =   1920
      TabIndex        =   3
      Top             =   600
      Width           =   1215
      _ExtentX        =   2143
      _ExtentY        =   661
      BTYPE           =   5
      TX              =   "Disconnect"
      ENAB            =   -1  'True
      BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      COLTYPE         =   1
      FOCUSR          =   -1  'True
      BCOL            =   13160660
      BCOLO           =   13160660
      FCOL            =   0
      FCOLO           =   0
      MCOL            =   12632256
      MPTR            =   1
      MICON           =   "frmKoneksi.frx":001C
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin VB.Label Label1 
      Caption         =   "Port :"
      Height          =   255
      Left            =   120
      TabIndex        =   0
      Top             =   240
      Width           =   735
   End
End
Attribute VB_Name = "frmKoneksi"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdConnect_Click()
'Dim s As String
'Dim nocomport As String
'On Error GoTo err_handler
'nocomport = Mid(cmbPort.Text, 4)
' hp_konek = False: err_konek = False

'With utama.com
'
'    If .PortOpen = False Then
'         utama.sbar.Panels(4).Text = "Tunggu..."
'        .CommPort = nocomport
'        .Settings = "19200,n,8,1"
'        .DTREnable = True
'        .RTSEnable = True
'        .RThreshold = 1
'        .InputLen = 1
'        .PortOpen = True
'        bacaMerkHp
'        bacatipe
'        provid
'        status_konek = True
'
'        utama.Timer1 = True
'        cmdConnect.Enabled = False
'        Me.Hide
'        MsgBox "koneksi berhasil"
'    Else
'        .PortOpen = False
'        status_konek = False
'        MsgBox "Tidak ada koneksi"
'    End If
'    utama.sbar.Panels(4).Text = "Status: Connect"
'End With
'
'Exit Sub
'err_handler:
'    MsgBox Err.Number & Err.Description
End Sub

Private Sub cmdDC_Click()
'With utama
'If .com.PortOpen = True Then
'    .Timer1 = False
'     status_konek = False
'    .com.PortOpen = False
'    cmdConnect.Enabled = True
'    .sbar.Panels(4).Text = "Disconnect"
'End If
'End With
End Sub

Private Sub Form_Load()
NoPort
End Sub
Function NoPort()
Dim i As Integer
For i = 1 To 10
    cmbPort.AddItem ("COM" & i)
Next
End Function




