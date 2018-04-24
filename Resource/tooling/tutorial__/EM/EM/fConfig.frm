VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Begin VB.Form fConfig 
   BorderStyle     =   0  'None
   Caption         =   "fConfig"
   ClientHeight    =   6855
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   8115
   BeginProperty Font 
      Name            =   "Square721 BT"
      Size            =   9.75
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "fConfig.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   6855
   ScaleWidth      =   8115
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin EnterpriseManager.XpButton cmdTutup 
      Height          =   495
      Left            =   5655
      TabIndex        =   1
      Top             =   6240
      Width           =   2190
      _ExtentX        =   3863
      _ExtentY        =   873
      Caption         =   "Tutup Setting"
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
   Begin TabDlg.SSTab SSTab1 
      Height          =   5205
      Left            =   195
      TabIndex        =   0
      Top             =   900
      Width           =   7680
      _ExtentX        =   13547
      _ExtentY        =   9181
      _Version        =   393216
      Tabs            =   2
      TabHeight       =   520
      ForeColor       =   -2147483630
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Square721 BT"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      TabCaption(0)   =   "Phonebook"
      TabPicture(0)   =   "fConfig.frx":000C
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "Frame3"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).ControlCount=   1
      TabCaption(1)   =   "Data Pengguna"
      TabPicture(1)   =   "fConfig.frx":0028
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "Frame5"
      Tab(1).ControlCount=   1
      Begin VB.Frame Frame5 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   4770
         Left            =   -74925
         TabIndex        =   12
         Top             =   345
         Width           =   7530
         Begin VB.TextBox txtUser 
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   300
            IMEMode         =   3  'DISABLE
            Index           =   0
            Left            =   5580
            TabIndex        =   19
            Top             =   810
            Width           =   1485
         End
         Begin VB.TextBox txtUser 
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   300
            IMEMode         =   3  'DISABLE
            Index           =   1
            Left            =   5580
            TabIndex        =   17
            Top             =   1335
            Width           =   1485
         End
         Begin VB.CheckBox chkAktif 
            Caption         =   "Aktif"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   345
            Left            =   5610
            TabIndex        =   15
            Top             =   2835
            Width           =   1395
         End
         Begin VB.TextBox txtUser 
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   300
            IMEMode         =   3  'DISABLE
            Index           =   2
            Left            =   5595
            PasswordChar    =   "O"
            TabIndex        =   14
            Top             =   1860
            Width           =   1485
         End
         Begin VB.TextBox txtUser 
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   300
            IMEMode         =   3  'DISABLE
            Index           =   3
            Left            =   5595
            PasswordChar    =   "O"
            TabIndex        =   13
            Top             =   2385
            Width           =   1485
         End
         Begin EnterpriseManager.XpButton cmdUser 
            Height          =   360
            Index           =   0
            Left            =   4080
            TabIndex        =   16
            Top             =   3975
            Width           =   945
            _ExtentX        =   1667
            _ExtentY        =   635
            Caption         =   "Simpan"
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
         Begin MSComctlLib.ListView LVUser 
            Height          =   3675
            Left            =   270
            TabIndex        =   18
            Top             =   660
            Width           =   3600
            _ExtentX        =   6350
            _ExtentY        =   6482
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
            Appearance      =   1
            NumItems        =   3
            BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               Text            =   "Username"
               Object.Width           =   1764
            EndProperty
            BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   1
               Text            =   "Nama"
               Object.Width           =   2540
            EndProperty
            BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   2
               Text            =   "Status"
               Object.Width           =   1411
            EndProperty
         End
         Begin EnterpriseManager.XpButton cmdUser 
            Height          =   360
            Index           =   1
            Left            =   5205
            TabIndex        =   20
            Top             =   3975
            Width           =   945
            _ExtentX        =   1667
            _ExtentY        =   635
            Caption         =   "Baru"
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
         Begin EnterpriseManager.XpButton cmdUser 
            Height          =   360
            Index           =   2
            Left            =   6285
            TabIndex        =   21
            Top             =   3975
            Width           =   945
            _ExtentX        =   1667
            _ExtentY        =   635
            Caption         =   "Hapus"
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
         Begin EnterpriseManager.XpButton cmdUser 
            Height          =   360
            Index           =   3
            Left            =   5595
            TabIndex        =   22
            Top             =   1830
            Width           =   1500
            _ExtentX        =   2646
            _ExtentY        =   635
            Caption         =   "Ganti Password"
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
         Begin VB.Label Label3 
            AutoSize        =   -1  'True
            Caption         =   "Nama Lengkap"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   240
            Index           =   1
            Left            =   4005
            TabIndex        =   28
            Top             =   1375
            Width           =   1305
         End
         Begin VB.Label Label4 
            AutoSize        =   -1  'True
            Caption         =   "Data Pengguna"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   12
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Left            =   285
            TabIndex        =   27
            Top             =   270
            Width           =   1770
         End
         Begin VB.Label Label3 
            AutoSize        =   -1  'True
            Caption         =   "Nama Pengenal"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   240
            Index           =   0
            Left            =   4005
            TabIndex        =   26
            Top             =   870
            Width           =   1380
         End
         Begin VB.Label Label3 
            AutoSize        =   -1  'True
            Caption         =   "Status"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   240
            Index           =   2
            Left            =   4035
            TabIndex        =   25
            Top             =   2895
            Width           =   570
         End
         Begin VB.Label Label3 
            AutoSize        =   -1  'True
            Caption         =   "Password"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   240
            Index           =   3
            Left            =   4005
            TabIndex        =   24
            Top             =   1880
            Width           =   855
         End
         Begin VB.Label Label3 
            AutoSize        =   -1  'True
            Caption         =   "Ulangi Password"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   240
            Index           =   4
            Left            =   4005
            TabIndex        =   23
            Top             =   2400
            Width           =   1455
         End
      End
      Begin VB.Frame Frame3 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   4770
         Left            =   75
         TabIndex        =   2
         Top             =   345
         Width           =   7530
         Begin VB.TextBox txtCari 
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Left            =   1365
            TabIndex        =   30
            Top             =   4350
            Width           =   2490
         End
         Begin VB.TextBox txtPB 
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   300
            IMEMode         =   3  'DISABLE
            Index           =   1
            Left            =   5565
            TabIndex        =   4
            Top             =   1335
            Width           =   1485
         End
         Begin VB.TextBox txtPB 
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   300
            IMEMode         =   3  'DISABLE
            Index           =   0
            Left            =   5565
            TabIndex        =   3
            Top             =   810
            Width           =   1485
         End
         Begin EnterpriseManager.XpButton cmdPB 
            Height          =   360
            Index           =   0
            Left            =   4080
            TabIndex        =   5
            Top             =   3975
            Width           =   945
            _ExtentX        =   1667
            _ExtentY        =   635
            Caption         =   "Simpan"
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
         Begin MSComctlLib.ListView LVPB 
            Height          =   3675
            Left            =   270
            TabIndex        =   6
            Top             =   660
            Width           =   3600
            _ExtentX        =   6350
            _ExtentY        =   6482
            View            =   3
            LabelEdit       =   1
            LabelWrap       =   -1  'True
            HideSelection   =   0   'False
            FullRowSelect   =   -1  'True
            GridLines       =   -1  'True
            _Version        =   393217
            ForeColor       =   -2147483640
            BackColor       =   -2147483643
            BorderStyle     =   1
            Appearance      =   1
            NumItems        =   2
            BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               Text            =   "Nomor"
               Object.Width           =   2540
            EndProperty
            BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   1
               Text            =   "Nama"
               Object.Width           =   2540
            EndProperty
         End
         Begin EnterpriseManager.XpButton cmdPB 
            Height          =   360
            Index           =   1
            Left            =   5205
            TabIndex        =   7
            Top             =   3975
            Width           =   945
            _ExtentX        =   1667
            _ExtentY        =   635
            Caption         =   "Baru"
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
         Begin EnterpriseManager.XpButton cmdPB 
            Height          =   360
            Index           =   2
            Left            =   6285
            TabIndex        =   8
            Top             =   3975
            Width           =   945
            _ExtentX        =   1667
            _ExtentY        =   635
            Caption         =   "Hapus"
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
         Begin VB.Label Label3 
            AutoSize        =   -1  'True
            Caption         =   "Cari Nama"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   240
            Index           =   7
            Left            =   300
            TabIndex        =   29
            Top             =   4395
            Width           =   915
         End
         Begin VB.Label Label3 
            AutoSize        =   -1  'True
            Caption         =   "Nomor Telpon"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   240
            Index           =   5
            Left            =   3990
            TabIndex        =   11
            Top             =   870
            Width           =   1200
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            Caption         =   "Data BukuTelpon"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   12
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Left            =   285
            TabIndex        =   10
            Top             =   270
            Width           =   1995
         End
         Begin VB.Label Label3 
            AutoSize        =   -1  'True
            Caption         =   "Nama"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   240
            Index           =   6
            Left            =   3990
            TabIndex        =   9
            Top             =   1380
            Width           =   510
         End
      End
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Setting Aplikasi"
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
      Left            =   2235
      TabIndex        =   31
      Top             =   135
      Width           =   3870
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      BackColor       =   &H00C0C0C0&
      Height          =   600
      Index           =   8
      Left            =   0
      TabIndex        =   32
      Top             =   0
      Width           =   8070
   End
End
Attribute VB_Name = "fConfig"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim RS As ADODB.Recordset
Dim PBBaru As Boolean

Private Sub cmdPB_Click(Index As Integer)
On Error GoTo KirimErr

Select Case Index
Case 0
    If Len(txtPB(0)) = 0 Then
        TampilkanPesan "Nomor telpon masih kosong"
        Exit Sub
    End If
    If Not IsNumberValid(txtPB(0)) Then
        TampilkanPesan "Nomor telpon tidak valid"
        Exit Sub
    End If
    If Len(txtPB(1)) = 0 Then
        TampilkanPesan "Nama pengenal telpon masih kosong"
        Exit Sub
    End If
    
    RS.Open "select * from phonebook where nomer regexp " & VldTeks(GetMainNumber(txtPB(0))), OC
    If RS.EOF Then
        OC.Execute "insert into phonebook(nomer,nama) values(" & _
                    VldTeks(txtPB(0)) & "," & VldTeks(txtPB(1)) & ")"
        TampilkanPesan "Data buku telpon telah tersimpan"
    Else
        If PBBaru Then
            If TampilkanTanya("Nomor " & VldTeks(txtPB(0)) & " sudah ada dengan nama " & VldTeks(RS("nama")) & vbCrLf & _
                        " apakah akan ditumpuki?") = msgdll.BS_Button1 Then
                    OC.Execute "update phonebook set nomer=" & VldTeks(txtPB(0)) & ",nama=" & VldTeks(txtPB(1)) & _
                            " where id=" & RS("id")
                    TampilkanPesan "Data buku telpon telah tersimpan"
            End If
        Else
                    OC.Execute "update phonebook set nomer=" & VldTeks(txtPB(0)) & ",nama=" & VldTeks(txtPB(1)) & _
                                " where id=" & RS("id")
                    TampilkanPesan "Data buku telpon telah tersimpan"
        End If
    End If
    RS.Close
    BacaPB
    DataBaruPB
    
Case 1
    DataBaruPB
    
Case 2
    If Len(txtPB(0)) = 0 Then
        TampilkanPesan "Nomor masih kosong"
        Exit Sub
    End If
    RS.Open "select id from phonebook where nomer regexp " & VldTeks(GetMainNumber(txtPB(0))), OC
    If RS.EOF Then
        TampilkanPesan "Nomor " & VldTeks(txtPB(0)) & " tidak ditemukan"
    Else
        If TampilkanTanya("Yakin nomor " & VldTeks(txtUser(0)) & " akan dihapus?") = BS_Button1 Then
            OC.Execute "delete from phonebook where id=" & RS("id")
            TampilkanPesan "Nomor telpon telah dihapus"
            DataBaruPB
        End If
    End If
    RS.Close
    BacaPB

End Select

Exit Sub

KirimErr:
    TampilkanPesan "Terjadi kesalahan" & vbCrLf & _
                    "Pesan: " & Err.Description & vbCrLf & _
                    "Kode: " & Err.Number

End Sub

Private Sub cmdTutup_Click()
Unload Me
End Sub

Private Sub cmdUser_Click(Index As Integer)
On Error GoTo KirimErr

Select Case Index
Case 0
    If Len(txtUser(0)) = 0 Then
        TampilkanPesan "Nama user masih kosong"
        Exit Sub
    End If
    If Len(txtUser(1)) = 0 Then
        TampilkanPesan "Nama lengkap masih kosong"
        Exit Sub
    End If
    
    RS.Open "select username from pengguna where username=" & VldTeks(txtUser(0)), OC
    If RS.EOF Then
    
        If Len(txtUser(2)) = 0 Then
            TampilkanPesan "Password masih kosong"
            RS.Close
            Exit Sub
        End If
        If txtUser(2) <> txtUser(3) Then
            TampilkanPesan "Kedua password harus diisi sama"
            RS.Close
            Exit Sub
        End If
    
        OC.Execute "insert into pengguna(username,nama,password,aktif) values(" & _
                    VldTeks(txtUser(0)) & "," & VldTeks(txtUser(1)) & ",password(" & VldTeks(txtUser(2)) & ")," & chkAktif.Value & ")"
    Else
        OC.Execute "update pengguna set nama=" & VldTeks(txtUser(1)) & ",aktif=" & chkAktif.Value & _
                    " where username=" & VldTeks(txtUser(0))
    End If
    RS.Close
    TampilkanPesan "Data pengguna telah tersimpan"
    BacaUser
    DataBaru
    
Case 1
    DataBaru
    
Case 2
    If Len(txtUser(0)) = 0 Then
        TampilkanPesan "Nama user masih kosong"
        Exit Sub
    End If
    RS.Open "select username from pengguna where username=" & VldTeks(txtUser(0)), OC
    If RS.EOF Then
        TampilkanPesan "Nama pengenal " & VldTeks(txtUser(0)) & " tidak ditemukan"
    Else
        If TampilkanTanya("Yakin nama pengenal " & VldTeks(txtUser(0)) & " akan dihapus?") = BS_Button1 Then
            OC.Execute "delete from pengguna where username=" & VldTeks(txtUser(0))
            TampilkanPesan "Nama pengenal telah dihapus"
            DataBaru
        End If
    End If
    RS.Close
    BacaUser
    
Case 3
    fGantiPwd.namauser = txtUser(0)
    fGantiPwd.Show vbModal, Me
End Select

Exit Sub

KirimErr:
    TampilkanPesan "Terjadi kesalahan" & vbCrLf & _
                    "Pesan: " & Err.Description & vbCrLf & _
                    "Kode: " & Err.Number
End Sub

Private Sub Form_Load()
Set RS = New ADODB.Recordset
BacaUser
BacaPB
End Sub

Private Sub Form_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
AForm.DragForm Me.hWnd
End Sub

Sub BacaUser()
Dim l As ListItem
LVUser.ListItems.Clear

On Error GoTo KirimErr

RS.Open "select * from pengguna", OC
While Not RS.EOF
    Set l = LVUser.ListItems.Add(, , RS("username"))
    l.SubItems(1) = RS("nama")
    l.SubItems(2) = IIf(RS("aktif") = 1, "Aktif", "Non Aktif")
    RS.MoveNext
Wend
RS.Close
If LVUser.ListItems.Count > 0 Then LVUser_ItemClick LVUser.ListItems(1)

Exit Sub

KirimErr:
    TampilkanPesan "Terjadi kesalahan" & vbCrLf & _
                    "Pesan: " & Err.Description & vbCrLf & _
                    "Kode: " & Err.Number
End Sub

Sub BacaPB()
Dim l As ListItem
On Error GoTo KirimErr

lvPB.ListItems.Clear
RS.Open "select * from phonebook order by nama", OC
While Not RS.EOF
    Set l = lvPB.ListItems.Add(, "s" & RS("id"), RS("nomer"))
    l.SubItems(1) = RS("nama")
    RS.MoveNext
Wend
RS.Close
If lvPB.ListItems.Count > 0 Then LVPB_ItemClick lvPB.ListItems(1)

Exit Sub

KirimErr:
    TampilkanPesan "Terjadi kesalahan" & vbCrLf & _
                    "Pesan: " & Err.Description & vbCrLf & _
                    "Kode: " & Err.Number
End Sub

Private Sub Form_Unload(Cancel As Integer)
Set RS = Nothing
End Sub

Private Sub Label2_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
AForm.DragForm Me.hWnd
End Sub

Private Sub Label3_MouseDown(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
AForm.DragForm Me.hWnd
End Sub

Private Sub LVPB_ItemClick(ByVal Item As MSComctlLib.ListItem)
txtPB(0) = Item.Text
txtPB(1) = Item.SubItems(1)
PBBaru = False
End Sub

Private Sub LVUser_ItemClick(ByVal Item As MSComctlLib.ListItem)
txtUser(0) = Item.Text
txtUser(1) = Item.SubItems(1)
chkAktif.Value = IIf(Item.SubItems(2) = "Aktif", 1, 0)
txtUser(0).Enabled = False
txtUser(2).Visible = False
txtUser(3).Visible = False
cmdUser(3).Visible = True
Label3(4).Visible = False
cmdUser(2).Enabled = (UCase(Item.Text) <> "ADMIN")
End Sub

Sub DataBaru()
Dim a As TextBox
For Each a In txtUser
    a = ""
Next
chkAktif.Value = 0
txtUser(0).Enabled = True
txtUser(2).Visible = True
txtUser(3).Visible = True
cmdUser(3).Visible = False
Label3(4).Visible = True
End Sub

Sub DataBaruPB()
txtPB(0) = ""
txtPB(1) = ""
PBBaru = True
End Sub

Private Sub txtCari_Change()
Dim i As Integer
If txtCari = "" Then Exit Sub
RS.Open "select nomer from phonebook where nama like '%" & txtCari & "%'", OC
If Not RS.EOF Then
  For i = 1 To lvPB.ListItems.Count
    If GetMainNumber(lvPB.ListItems(i).Text) = GetMainNumber(RS("nomer")) Then
      LVPB_ItemClick lvPB.ListItems(i)
    End If
  Next
End If
RS.Close
End Sub
