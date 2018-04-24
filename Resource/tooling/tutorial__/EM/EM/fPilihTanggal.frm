VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form fPilihTanggal 
   BorderStyle     =   0  'None
   Caption         =   "Pilih Tanggal"
   ClientHeight    =   2820
   ClientLeft      =   0
   ClientTop       =   -105
   ClientWidth     =   5565
   BeginProperty Font 
      Name            =   "Arial"
      Size            =   9.75
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "fPilihTanggal.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2820
   ScaleWidth      =   5565
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin MSComCtl2.DTPicker DTPicker2 
      Height          =   330
      Left            =   2010
      TabIndex        =   7
      Top             =   1275
      Width           =   3270
      _ExtentX        =   5768
      _ExtentY        =   582
      _Version        =   393216
      CustomFormat    =   "ddd, d MMM yyyy"
      Format          =   51576835
      CurrentDate     =   38147
   End
   Begin MSComCtl2.DTPicker DTPicker1 
      Height          =   330
      Left            =   2010
      TabIndex        =   6
      Top             =   705
      Width           =   3270
      _ExtentX        =   5768
      _ExtentY        =   582
      _Version        =   393216
      CustomFormat    =   "ddd, d MMM yyyy"
      Format          =   51576835
      CurrentDate     =   38147
   End
   Begin EnterpriseManager.XpButton xpcmd 
      Height          =   465
      Index           =   0
      Left            =   2115
      TabIndex        =   4
      Top             =   2010
      Width           =   1140
      _ExtentX        =   2011
      _ExtentY        =   820
      Caption         =   "Pilih"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin EnterpriseManager.XpButton xpcmd 
      Height          =   465
      Index           =   1
      Left            =   3435
      TabIndex        =   5
      Top             =   2025
      Width           =   1140
      _ExtentX        =   2011
      _ExtentY        =   820
      Caption         =   "Batal"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.Label Label4 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Pilih Tanggal Batasan"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   405
      TabIndex        =   0
      Top             =   180
      Width           =   3870
   End
   Begin VB.Label Label2 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      Caption         =   "Sampai Tanggal:"
      Height          =   240
      Left            =   450
      TabIndex        =   3
      Top             =   1335
      Width           =   1485
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      Caption         =   "SMS Mulai Tanggal:"
      Height          =   240
      Left            =   120
      TabIndex        =   1
      Top             =   765
      Width           =   1815
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      BackColor       =   &H00C0C0C0&
      Height          =   600
      Left            =   0
      TabIndex        =   2
      Top             =   0
      Width           =   8325
   End
End
Attribute VB_Name = "fPilihTanggal"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Load()
DTPicker1.Value = fUtama.TglFilter1
DTPicker2.Value = fUtama.TglFilter2
fUtama.DateChanged = False
End Sub

Private Sub Label3_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
AForm.DragForm Me.hWnd
End Sub

Private Sub Label4_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
AForm.DragForm Me.hWnd
End Sub

Private Sub xpCmd_Click(Index As Integer)
fUtama.DateChanged = (Index = 0)
If Index = 0 Then
    fUtama.TglFilter1 = DTPicker1.Value
    fUtama.TglFilter2 = DTPicker2.Value
End If
Unload Me
End Sub
