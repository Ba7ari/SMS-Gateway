VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form fSendSMS 
   BorderStyle     =   0  'None
   Caption         =   "Send SMS"
   ClientHeight    =   3900
   ClientLeft      =   270
   ClientTop       =   1260
   ClientWidth     =   5205
   BeginProperty Font 
      Name            =   "Arial"
      Size            =   9.75
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "fSendSMS.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3900
   ScaleWidth      =   5205
   ShowInTaskbar   =   0   'False
   Begin MSComctlLib.ListView lvPB 
      Height          =   2460
      Left            =   1905
      TabIndex        =   7
      Top             =   1365
      Visible         =   0   'False
      Width           =   3060
      _ExtentX        =   5398
      _ExtentY        =   4339
      View            =   3
      LabelEdit       =   1
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      FullRowSelect   =   -1  'True
      GridLines       =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   0
      NumItems        =   2
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "Nama"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "Nomor"
         Object.Width           =   2540
      EndProperty
   End
   Begin EnterpriseManager.XpButton xpCmd 
      Height          =   360
      Index           =   0
      Left            =   3840
      TabIndex        =   2
      Top             =   1065
      Width           =   1095
      _ExtentX        =   1931
      _ExtentY        =   635
      Caption         =   "Phonebook..."
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
   Begin VB.TextBox txtPesan 
      Height          =   1590
      Left            =   1395
      MaxLength       =   160
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   1
      Top             =   1560
      Width           =   3525
   End
   Begin VB.TextBox txtNomor 
      Height          =   315
      Left            =   1410
      TabIndex        =   0
      Top             =   1110
      Width           =   2325
   End
   Begin EnterpriseManager.XpButton xpCmd 
      Height          =   360
      Index           =   1
      Left            =   1365
      TabIndex        =   3
      Top             =   3330
      Width           =   1095
      _ExtentX        =   1931
      _ExtentY        =   635
      Caption         =   "Kirim"
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
   Begin EnterpriseManager.XpButton xpCmd 
      Height          =   360
      Index           =   2
      Left            =   2790
      TabIndex        =   4
      Top             =   3330
      Width           =   1095
      _ExtentX        =   1931
      _ExtentY        =   635
      Caption         =   "Tutup"
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
      Caption         =   "Kirim SMS"
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
      Left            =   690
      TabIndex        =   8
      Top             =   195
      Width           =   3870
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Pesan"
      Height          =   240
      Left            =   720
      TabIndex        =   6
      Top             =   1605
      Width           =   555
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Nomor Tujuan"
      Height          =   240
      Left            =   75
      TabIndex        =   5
      Top             =   1140
      Width           =   1200
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      BackColor       =   &H00C0C0C0&
      Height          =   600
      Index           =   0
      Left            =   0
      TabIndex        =   9
      Top             =   0
      Width           =   5160
   End
End
Attribute VB_Name = "fSendSMS"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Form_Load()
Dim RS As ADODB.Recordset
Dim l As ListItem

On Error GoTo KirimErr

Set RS = New ADODB.Recordset

With lvPB
    .Visible = False
    RS.Open "select * from phonebook order by nama asc", OC
    While Not RS.EOF
        Set l = .ListItems.Add(, , RS("nomer"))
        l.SubItems(1) = RS("nama")
        RS.MoveNext
    Wend
    RS.Close
End With
Set RS = Nothing

txtPesan = ""
txtPesan_Change


Exit Sub

KirimErr:
    TampilkanPesan "Terjadi kesalahan" & vbCrLf & _
                    "Pesan: " & Err.Description & vbCrLf & _
                    "Kode: " & Err.Number

End Sub

Private Sub Label3_MouseDown(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
AForm.DragForm Me.hWnd
End Sub

Private Sub Label4_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
AForm.DragForm Me.hWnd
End Sub

Private Sub lvPB_Click()
If Not (lvPB.SelectedItem Is Nothing) Then
    If txtNomor <> "" Then
        txtNomor = txtNomor & ";"
    End If
    txtNomor = txtNomor & lvPB.SelectedItem.Text
End If
lvPB.Visible = False
End Sub

Private Sub lvPB_LostFocus()
lvPB.Visible = False
End Sub

Private Sub txtNomor_Change()
txtPesan_Change
End Sub

Private Sub txtNomor_GotFocus()
lvPB.Visible = False
End Sub

Private Sub txtPesan_Change()
xpcmd(1).Enabled = ((Len(txtNomor) > 0) And (Len(txtPesan) > 0))
End Sub

Private Sub txtPesan_GotFocus()
lvPB.Visible = False
End Sub

Private Sub xpCmd_Click(Index As Integer)
Dim X
Dim i As Integer, c As Integer

Select Case Index
Case 0
    lvPB.Visible = True
    
Case 1
    X = Split(txtNomor, ";")
    c = 0
    For i = LBound(X) To UBound(X)
        If IsNumberValid(Trim(X(i))) Then
            KirimSMS X(i), txtPesan
            c = c + 1
        Else
            TampilkanPesan "Nomor " & X(i) & " tidak valid"
        End If
    Next
    TampilkanPesan c & " pesan dimasukkan ke Outbox"
    Unload Me
Case 2
    Unload Me
End Select
End Sub
