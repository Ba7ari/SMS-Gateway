VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{531653E6-69A2-40A4-8734-91C78B33DD08}#1.0#0"; "XPButton.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form masterKelas 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Input Maskapai"
   ClientHeight    =   4755
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   10035
   Icon            =   "masterKelas.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   4755
   ScaleWidth      =   10035
   ShowInTaskbar   =   0   'False
   Begin TabDlg.SSTab SSTab1 
      Height          =   4575
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   9885
      _ExtentX        =   17436
      _ExtentY        =   8070
      _Version        =   393216
      TabHeight       =   520
      TabCaption(0)   =   "Input Route"
      TabPicture(0)   =   "masterKelas.frx":A4CA
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "Shape2"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "cmdbatalInputRoute"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "cmdsimpanRoute"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "cmdubahRoute"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).Control(4)=   "cmdHpsRoute"
      Tab(0).Control(4).Enabled=   0   'False
      Tab(0).Control(5)=   "frmsroute"
      Tab(0).Control(5).Enabled=   0   'False
      Tab(0).Control(6)=   "LVRoute"
      Tab(0).Control(6).Enabled=   0   'False
      Tab(0).Control(7)=   "frmEditRoute"
      Tab(0).Control(7).Enabled=   0   'False
      Tab(0).ControlCount=   8
      TabCaption(1)   =   "Input Flight"
      TabPicture(1)   =   "masterKelas.frx":A4E6
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "Shape1"
      Tab(1).Control(0).Enabled=   0   'False
      Tab(1).Control(1)=   "cmdbatalInputFlight"
      Tab(1).Control(1).Enabled=   0   'False
      Tab(1).Control(2)=   "LVMaskapai"
      Tab(1).Control(2).Enabled=   0   'False
      Tab(1).Control(3)=   "cmdHapusMaskapai"
      Tab(1).Control(3).Enabled=   0   'False
      Tab(1).Control(4)=   "cmdUbahMaskapai"
      Tab(1).Control(4).Enabled=   0   'False
      Tab(1).Control(5)=   "cmdSimpanMaskapai"
      Tab(1).Control(5).Enabled=   0   'False
      Tab(1).Control(6)=   "Frame1"
      Tab(1).Control(6).Enabled=   0   'False
      Tab(1).Control(7)=   "cmbEditFlt"
      Tab(1).Control(7).Enabled=   0   'False
      Tab(1).ControlCount=   8
      TabCaption(2)   =   "Input Tarif"
      TabPicture(2)   =   "masterKelas.frx":A502
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "cmdKeluar"
      Tab(2).Control(0).Enabled=   0   'False
      Tab(2).Control(1)=   "cmdBatal"
      Tab(2).Control(1).Enabled=   0   'False
      Tab(2).Control(2)=   "cmdDetail"
      Tab(2).Control(2).Enabled=   0   'False
      Tab(2).Control(3)=   "cmdHapusClass"
      Tab(2).Control(3).Enabled=   0   'False
      Tab(2).Control(4)=   "cmdUbahClass"
      Tab(2).Control(4).Enabled=   0   'False
      Tab(2).Control(5)=   "cmdSimpanClass"
      Tab(2).Control(5).Enabled=   0   'False
      Tab(2).Control(6)=   "Frame2"
      Tab(2).Control(6).Enabled=   0   'False
      Tab(2).ControlCount=   7
      Begin VB.ComboBox cmbEditFlt 
         Height          =   315
         Left            =   -73560
         TabIndex        =   52
         Top             =   1320
         Visible         =   0   'False
         Width           =   975
      End
      Begin VB.Frame frmEditRoute 
         Height          =   1335
         Left            =   3720
         TabIndex        =   47
         Top             =   3600
         Visible         =   0   'False
         Width           =   3375
         Begin VB.TextBox txtUbahNmRoute 
            Height          =   318
            Left            =   1320
            TabIndex        =   51
            Top             =   600
            Width           =   1455
         End
         Begin VB.ComboBox cmbUbahkdRoute 
            Height          =   315
            Left            =   1320
            TabIndex        =   48
            Top             =   240
            Width           =   975
         End
         Begin VB.Label Label17 
            Caption         =   "Nama Route"
            Height          =   255
            Left            =   240
            TabIndex        =   50
            Top             =   720
            Width           =   975
         End
         Begin VB.Label Label16 
            Caption         =   "Kd.Route"
            Height          =   255
            Left            =   240
            TabIndex        =   49
            Top             =   360
            Width           =   735
         End
      End
      Begin VB.Frame Frame2 
         Height          =   3495
         Left            =   -74760
         TabIndex        =   23
         Top             =   480
         Width           =   9495
         Begin VB.ComboBox cmbNmClass 
            Height          =   315
            ItemData        =   "masterKelas.frx":A51E
            Left            =   3840
            List            =   "masterKelas.frx":A52B
            TabIndex        =   58
            Text            =   "Flexi"
            Top             =   240
            Width           =   1095
         End
         Begin VB.TextBox txtppn2 
            Height          =   315
            Left            =   1800
            TabIndex        =   45
            Top             =   1680
            Width           =   1095
         End
         Begin VB.TextBox txtstock 
            Height          =   315
            Left            =   5880
            TabIndex        =   40
            Top             =   1680
            Width           =   615
         End
         Begin VB.TextBox txtselling 
            Height          =   315
            Left            =   4440
            TabIndex        =   38
            Top             =   1680
            Width           =   1335
         End
         Begin VB.TextBox txtiwjr 
            Height          =   315
            Left            =   3000
            TabIndex        =   35
            Top             =   1680
            Width           =   1335
         End
         Begin VB.TextBox txtppn 
            Height          =   315
            Left            =   1320
            TabIndex        =   34
            Top             =   1680
            Width           =   495
         End
         Begin VB.TextBox txtFuel 
            Height          =   315
            Left            =   120
            TabIndex        =   32
            Top             =   1680
            Width           =   1095
         End
         Begin VB.TextBox txtbasic 
            Height          =   315
            Left            =   1320
            TabIndex        =   30
            Top             =   960
            Width           =   1335
         End
         Begin VB.ComboBox cmbNoflt 
            Height          =   315
            Left            =   1320
            TabIndex        =   28
            Top             =   600
            Width           =   1215
         End
         Begin VB.TextBox txtKdKelas 
            Height          =   315
            Left            =   1320
            MaxLength       =   3
            TabIndex        =   25
            Top             =   240
            Width           =   375
         End
         Begin MSComctlLib.ListView LVClass 
            Height          =   1335
            Left            =   120
            TabIndex        =   41
            Top             =   2040
            Width           =   9255
            _ExtentX        =   16325
            _ExtentY        =   2355
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
            NumItems        =   9
            BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               Text            =   "Class"
               Object.Width           =   882
            EndProperty
            BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   1
               Text            =   "Nm.Class"
               Object.Width           =   1764
            EndProperty
            BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   2
               Text            =   "No.flt"
               Object.Width           =   2540
            EndProperty
            BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   3
               Text            =   "Fare Basic"
               Object.Width           =   2540
            EndProperty
            BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   4
               Text            =   "Fuel Surc."
               Object.Width           =   882
            EndProperty
            BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   5
               Text            =   "PPN"
               Object.Width           =   2540
            EndProperty
            BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   6
               Text            =   "IWJR"
               Object.Width           =   2540
            EndProperty
            BeginProperty ColumnHeader(8) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   7
               Text            =   "Selling Price"
               Object.Width           =   2540
            EndProperty
            BeginProperty ColumnHeader(9) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   8
               Text            =   "Stock Class"
               Object.Width           =   2540
            EndProperty
         End
         Begin VB.Label lblTarifRoute 
            BorderStyle     =   1  'Fixed Single
            Height          =   315
            Left            =   3240
            TabIndex        =   54
            Top             =   720
            Width           =   1335
         End
         Begin VB.Label Label18 
            Caption         =   "Route"
            Height          =   255
            Left            =   2640
            TabIndex        =   53
            Top             =   720
            Width           =   495
         End
         Begin VB.Label Label15 
            Caption         =   "Stock Class"
            Height          =   375
            Left            =   5760
            TabIndex        =   39
            Top             =   1200
            Width           =   615
         End
         Begin VB.Label Label14 
            Caption         =   "Selling Price"
            Height          =   255
            Left            =   4440
            TabIndex        =   37
            Top             =   1440
            Width           =   1215
         End
         Begin VB.Label Label13 
            Caption         =   "IWJR ( Asuransi )"
            Height          =   255
            Left            =   3000
            TabIndex        =   36
            Top             =   1440
            Width           =   1455
         End
         Begin VB.Label Label12 
            Caption         =   "PPN"
            Height          =   255
            Left            =   1800
            TabIndex        =   33
            Top             =   1440
            Width           =   495
         End
         Begin VB.Label Label11 
            Caption         =   "Fuel Surc."
            Height          =   255
            Left            =   240
            TabIndex        =   31
            Top             =   1440
            Width           =   975
         End
         Begin VB.Label Label10 
            Caption         =   "Fare Basic"
            Height          =   255
            Left            =   240
            TabIndex        =   29
            Top             =   1080
            Width           =   975
         End
         Begin VB.Label Label9 
            Caption         =   "No.flt"
            Height          =   375
            Left            =   240
            TabIndex        =   27
            Top             =   720
            Width           =   975
         End
         Begin VB.Label Label8 
            Caption         =   "Nama Class"
            Height          =   375
            Left            =   2760
            TabIndex        =   26
            Top             =   360
            Width           =   975
         End
         Begin VB.Label Label7 
            Caption         =   "Kode Class"
            Height          =   375
            Left            =   240
            TabIndex        =   24
            Top             =   360
            Width           =   975
         End
      End
      Begin MSComctlLib.ListView LVRoute 
         Height          =   2895
         Left            =   4680
         TabIndex        =   18
         Top             =   480
         Width           =   5055
         _ExtentX        =   8916
         _ExtentY        =   5106
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
         NumItems        =   2
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "Kode"
            Object.Width           =   1235
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "Nama Route"
            Object.Width           =   4410
         EndProperty
      End
      Begin VB.Frame frmsroute 
         Height          =   1215
         Left            =   240
         TabIndex        =   10
         Top             =   600
         Width           =   3015
         Begin VB.TextBox txtnmRoute 
            Height          =   318
            Left            =   1320
            TabIndex        =   12
            Top             =   600
            Width           =   1455
         End
         Begin VB.TextBox txtKdRoute 
            Height          =   318
            Left            =   1320
            MaxLength       =   3
            TabIndex        =   11
            Top             =   240
            Width           =   615
         End
         Begin VB.Label Label6 
            Caption         =   "Nama Route"
            Height          =   375
            Left            =   240
            TabIndex        =   14
            Top             =   720
            Width           =   1095
         End
         Begin VB.Label Label5 
            Caption         =   "Kd.Route"
            Height          =   255
            Left            =   240
            TabIndex        =   13
            Top             =   360
            Width           =   735
         End
      End
      Begin VB.Frame Frame1 
         Height          =   2055
         Left            =   -74880
         TabIndex        =   1
         Top             =   1080
         Width           =   3855
         Begin VB.TextBox txtflt 
            Height          =   318
            Left            =   1320
            TabIndex        =   3
            Text            =   "RI"
            Top             =   240
            Width           =   1215
         End
         Begin VB.ComboBox cmbrouteMaskapai 
            Height          =   315
            Left            =   1320
            TabIndex        =   2
            Top             =   600
            Width           =   975
         End
         Begin MSMask.MaskEdBox mskjadwal 
            Height          =   315
            Left            =   1320
            TabIndex        =   4
            Top             =   1320
            Width           =   615
            _ExtentX        =   1085
            _ExtentY        =   556
            _Version        =   393216
            MaxLength       =   5
            Mask            =   "&&:&&"
            PromptChar      =   "_"
         End
         Begin VB.Label lblroute 
            BorderStyle     =   1  'Fixed Single
            Height          =   315
            Left            =   1320
            TabIndex        =   6
            Top             =   960
            Width           =   1095
         End
         Begin VB.Label Label1 
            Caption         =   "No.Flt"
            Height          =   255
            Left            =   240
            TabIndex        =   9
            Top             =   360
            Width           =   735
         End
         Begin VB.Label Label2 
            Caption         =   "Route"
            Height          =   255
            Left            =   240
            TabIndex        =   8
            Top             =   720
            Width           =   615
         End
         Begin VB.Label Label3 
            Caption         =   "Nama Route."
            Height          =   255
            Left            =   240
            TabIndex        =   7
            Top             =   1080
            Width           =   1095
         End
         Begin VB.Label Label4 
            Caption         =   "Jadwal Brkt."
            Height          =   255
            Left            =   240
            TabIndex        =   5
            Top             =   1440
            Width           =   1095
         End
      End
      Begin ucXPButton.XPButton cmdHpsRoute 
         Height          =   375
         Left            =   3600
         TabIndex        =   15
         Top             =   1920
         Width           =   975
         _ExtentX        =   1720
         _ExtentY        =   661
         BTYPE           =   5
         TX              =   "Hapus"
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
         BCOL            =   14802912
         BCOLO           =   14802912
         FCOL            =   0
         FCOLO           =   0
         MCOL            =   12632256
         MPTR            =   1
         MICON           =   "masterKelas.frx":A545
         PICN            =   "masterKelas.frx":A561
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin ucXPButton.XPButton cmdubahRoute 
         Height          =   375
         Left            =   1440
         TabIndex        =   16
         Top             =   1920
         Width           =   975
         _ExtentX        =   1720
         _ExtentY        =   661
         BTYPE           =   5
         TX              =   "&Ubah"
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
         BCOL            =   14802912
         BCOLO           =   14802912
         FCOL            =   0
         FCOLO           =   0
         MCOL            =   12632256
         MPTR            =   1
         MICON           =   "masterKelas.frx":27163
         PICN            =   "masterKelas.frx":2717F
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin ucXPButton.XPButton cmdsimpanRoute 
         Height          =   375
         Left            =   240
         TabIndex        =   17
         Top             =   1920
         Width           =   1095
         _ExtentX        =   1931
         _ExtentY        =   661
         BTYPE           =   5
         TX              =   "Simpan"
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
         BCOL            =   14802912
         BCOLO           =   14802912
         FCOL            =   0
         FCOLO           =   0
         MCOL            =   12632256
         MPTR            =   1
         MICON           =   "masterKelas.frx":43D81
         PICN            =   "masterKelas.frx":43D9D
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin ucXPButton.XPButton cmdSimpanMaskapai 
         Height          =   375
         Left            =   -74880
         TabIndex        =   19
         Top             =   3480
         Width           =   1095
         _ExtentX        =   1931
         _ExtentY        =   661
         BTYPE           =   5
         TX              =   "Simpan"
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
         BCOL            =   14802912
         BCOLO           =   14802912
         FCOL            =   0
         FCOLO           =   0
         MCOL            =   12632256
         MPTR            =   1
         MICON           =   "masterKelas.frx":6099F
         PICN            =   "masterKelas.frx":609BB
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin ucXPButton.XPButton cmdUbahMaskapai 
         Height          =   375
         Left            =   -73680
         TabIndex        =   20
         Top             =   3480
         Width           =   1095
         _ExtentX        =   1931
         _ExtentY        =   661
         BTYPE           =   5
         TX              =   "Ubah"
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
         BCOL            =   14802912
         BCOLO           =   14802912
         FCOL            =   0
         FCOLO           =   0
         MCOL            =   12632256
         MPTR            =   1
         MICON           =   "masterKelas.frx":7D5BD
         PICN            =   "masterKelas.frx":7D5D9
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin ucXPButton.XPButton cmdHapusMaskapai 
         Height          =   375
         Left            =   -71280
         TabIndex        =   21
         Top             =   3480
         Width           =   1095
         _ExtentX        =   1931
         _ExtentY        =   661
         BTYPE           =   5
         TX              =   "Hapus"
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
         BCOL            =   14802912
         BCOLO           =   14802912
         FCOL            =   0
         FCOLO           =   0
         MCOL            =   12632256
         MPTR            =   1
         MICON           =   "masterKelas.frx":9A1DB
         PICN            =   "masterKelas.frx":9A1F7
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin ucXPButton.XPButton cmdSimpanClass 
         Height          =   375
         Left            =   -74760
         TabIndex        =   42
         Top             =   4080
         Width           =   1095
         _ExtentX        =   1931
         _ExtentY        =   661
         BTYPE           =   5
         TX              =   "Simpan"
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
         BCOL            =   14802912
         BCOLO           =   14802912
         FCOL            =   0
         FCOLO           =   0
         MCOL            =   12632256
         MPTR            =   1
         MICON           =   "masterKelas.frx":B6DF9
         PICN            =   "masterKelas.frx":B6E15
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin ucXPButton.XPButton cmdUbahClass 
         Height          =   375
         Left            =   -73560
         TabIndex        =   43
         Top             =   4080
         Width           =   1095
         _ExtentX        =   1931
         _ExtentY        =   661
         BTYPE           =   5
         TX              =   "Cari"
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
         BCOL            =   14802912
         BCOLO           =   14802912
         FCOL            =   0
         FCOLO           =   0
         MCOL            =   12632256
         MPTR            =   1
         MICON           =   "masterKelas.frx":D3A17
         PICN            =   "masterKelas.frx":D3A33
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin ucXPButton.XPButton cmdHapusClass 
         Height          =   375
         Left            =   -69960
         TabIndex        =   44
         Top             =   4080
         Width           =   1095
         _ExtentX        =   1931
         _ExtentY        =   661
         BTYPE           =   5
         TX              =   "Hapus"
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
         BCOL            =   14802912
         BCOLO           =   14802912
         FCOL            =   0
         FCOLO           =   0
         MCOL            =   12632256
         MPTR            =   1
         MICON           =   "masterKelas.frx":F0635
         PICN            =   "masterKelas.frx":F0651
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin ucXPButton.XPButton cmdDetail 
         Height          =   375
         Left            =   -72360
         TabIndex        =   46
         Top             =   4080
         Width           =   1095
         _ExtentX        =   1931
         _ExtentY        =   661
         BTYPE           =   5
         TX              =   "Detail"
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
         BCOL            =   14802912
         BCOLO           =   14802912
         FCOL            =   0
         FCOLO           =   0
         MCOL            =   12632256
         MPTR            =   1
         MICON           =   "masterKelas.frx":10D253
         PICN            =   "masterKelas.frx":10D26F
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin MSComctlLib.ListView LVMaskapai 
         Height          =   2415
         Left            =   -70800
         TabIndex        =   22
         Top             =   960
         Width           =   5415
         _ExtentX        =   9551
         _ExtentY        =   4260
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
            Text            =   "No.flt"
            Object.Width           =   1235
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "Route"
            Object.Width           =   4410
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   2
            Text            =   "Jadwal"
            Object.Width           =   2540
         EndProperty
      End
      Begin ucXPButton.XPButton cmdBatal 
         Height          =   375
         Left            =   -71160
         TabIndex        =   55
         Top             =   4080
         Width           =   1095
         _ExtentX        =   1931
         _ExtentY        =   661
         BTYPE           =   5
         TX              =   "&Batal"
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
         BCOL            =   14802912
         BCOLO           =   14802912
         FCOL            =   0
         FCOLO           =   0
         MCOL            =   12632256
         MPTR            =   1
         MICON           =   "masterKelas.frx":10D6C1
         PICN            =   "masterKelas.frx":10D6DD
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin ucXPButton.XPButton cmdbatalInputFlight 
         Height          =   375
         Left            =   -72480
         TabIndex        =   56
         Top             =   3480
         Width           =   1095
         _ExtentX        =   1931
         _ExtentY        =   661
         BTYPE           =   5
         TX              =   "&Batal"
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
         BCOL            =   14802912
         BCOLO           =   14802912
         FCOL            =   0
         FCOLO           =   0
         MCOL            =   12632256
         MPTR            =   1
         MICON           =   "masterKelas.frx":10FABF
         PICN            =   "masterKelas.frx":10FADB
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin ucXPButton.XPButton cmdbatalInputRoute 
         Height          =   375
         Left            =   2520
         TabIndex        =   57
         Top             =   1920
         Width           =   975
         _ExtentX        =   1720
         _ExtentY        =   661
         BTYPE           =   5
         TX              =   "&Batal"
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
         BCOL            =   14802912
         BCOLO           =   14802912
         FCOL            =   0
         FCOLO           =   0
         MCOL            =   12632256
         MPTR            =   1
         MICON           =   "masterKelas.frx":11AF15
         PICN            =   "masterKelas.frx":11AF31
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin ucXPButton.XPButton cmdKeluar 
         Height          =   375
         Left            =   -68760
         TabIndex        =   59
         Top             =   4080
         Width           =   1095
         _ExtentX        =   1931
         _ExtentY        =   661
         BTYPE           =   5
         TX              =   "&Keluar"
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
         BCOL            =   14802912
         BCOLO           =   14802912
         FCOL            =   0
         FCOLO           =   0
         MCOL            =   12632256
         MPTR            =   1
         MICON           =   "masterKelas.frx":12636B
         PICN            =   "masterKelas.frx":126387
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin VB.Shape Shape2 
         BackColor       =   &H00E0E0E0&
         BackStyle       =   1  'Opaque
         BorderStyle     =   6  'Inside Solid
         Height          =   735
         Left            =   120
         Top             =   3600
         Width           =   9615
      End
      Begin VB.Shape Shape1 
         BackColor       =   &H00E0E0E0&
         BackStyle       =   1  'Opaque
         BorderStyle     =   6  'Inside Solid
         Height          =   495
         Left            =   -74880
         Top             =   360
         Width           =   9615
      End
   End
End
Attribute VB_Name = "masterKelas"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim clsRoute As New ClassRoute
Dim clsmaskapai As New ClassMaskapai
Dim j As Integer
Dim noflt As String
Dim cData() As String, bagicdata() As String

Private Sub cmbEditFlt_Click()
Dim x As Integer
Dim xData() As String, bagixData() As String
On Error Resume Next
xData = clsmaskapai.cari_maskapai(cmbEditFlt.Text)
For x = 0 To UBound(xData)
        bagixData = Split(xData(x), "#")
        cmbrouteMaskapai.Text = bagixData(0)
        mskjadwal.Mask = ""
        mskjadwal.Width = 800
        mskjadwal.Text = bagixData(1)
'       MsgBox bagixData(1)
Next

End Sub

Private Sub cmbnoflt_Click()
lblTarifRoute.Caption = clsRoute.cari_Flight(cmbNoflt.Text)
txtbasic.SetFocus
End Sub

Private Sub cmbNoflt_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    txtbasic.SetFocus
End If
End Sub

Private Sub cmbrouteMaskapai_Click()
Dim i As Integer
Dim xData() As String
Dim xGetData() As String
On Error Resume Next 'Jika data kosong maka pass..
xData = clsRoute.CariRoute(cmbrouteMaskapai.Text)
For i = 0 To UBound(xData)
    xGetData = Split(xData(i))
    lblroute.Caption = xGetData(1)
Next i
mskjadwal.SetFocus
End Sub

Private Sub cmbrouteMaskapai_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    mskjadwal.SetFocus
End If
End Sub

Private Sub cmbUbahkdRoute_Click()
Dim clsRoute As New ClassRoute
Dim i As Integer
Dim xDataRoute() As String, bagiDataRoute() As String
xDataRoute = clsRoute.CariRoute(cmbUbahkdRoute.Text)
For i = 0 To UBound(xDataRoute)
        bagiDataRoute = Split(xDataRoute(i), " ")
        txtUbahNmRoute.Text = bagiDataRoute(1)
Next i
End Sub

Private Sub cmbUbahkdRoute_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    txtUbahNmRoute.SetFocus
End If
End Sub

Private Sub cmdBatal_Click()
bersihMaster_tarif
LVClass.ListItems.Clear
End Sub

Private Sub cmdbatalInputFlight_Click()
bersihkan_tabInputFlight 'panggil modul pembersih tab inputflight

End Sub

Private Sub cmdbatalInputRoute_Click()
bersihkanLayar
txtKdRoute.SetFocus
cmdubahRoute.Caption = "&Ubah"
cmdsimpanRoute.Enabled = True
frmEditRoute.Visible = False
frmsroute.Visible = True
End Sub

Private Sub cmdDetail_Click()
frmDelay.Show 1
frmDetailMaskapai.Show

End Sub

Private Sub cmdHapusClass_Click()
Dim clsmaskapai As New ClassMaskapai
Dim inBox As String
Dim bagiDat() As String, i As Integer
On Error GoTo pesan_error
inBox = InputBox("Masukkan No.Flight<spasi>Class", "Hapus No.Flight")
If Len(inBox) = 0 Then Exit Sub
'ReDim Preserve bagiDat(i)
bagiDat = Split(inBox, " ")
clsmaskapai.Hapus_NoFlt bagiDat(0), bagiDat(1)


Exit Sub
pesan_error:
    MsgBox Err.Description & Err.Number, vbCritical, "Gagal menghapus"

End Sub

Private Sub cmdHapusMaskapai_Click()
Dim clsmaskapai As New ClassMaskapai

If cmdHapusMaskapai.Caption = "Hapus" Then
'        ambilRoute_
        ambilNoflt
        OpsiUbahMaskapai True
        cmdHapusMaskapai.Enabled = True
        cmdHapusMaskapai.Caption = "OK"
Else
    
    If MsgBox("Data Penerbangan akan dihapus,Anda yakin?", vbOKCancel + vbQuestion + vbSystemModal, "Konfirmasi") = vbOK Then
    
        With clsmaskapai
                .cariHps_maskapai cmbEditFlt.Text, cmbrouteMaskapai.Text
        End With
        
        
    End If
    ambilMaskapai
    OpsiUbahMaskapai False
    cmdHapusMaskapai.Caption = "Hapus"
End If

End Sub

Private Sub cmdHpsRoute_Click()
Dim inMsg As String
inMsg = InputBox("Masukkan Kode Route yang ingin dihapus", "Hapus Route")
If Len(inMsg) = 0 Then Exit Sub
If MsgBox("Ingin menghapus Route '" & inMsg & "'?", vbOKCancel + vbQuestion) = vbOK Then
    With clsRoute
        .hapus_Route inMsg
    End With
    ambilRoute_
End If
End Sub


Private Sub cmdKeluar_Click()
Unload Me
End Sub

Private Sub cmdSimpanClass_Click()
Dim i As Integer
Dim clsClass As New ClassKelas
If LVClass.ListItems.Count = 0 Then
        MsgBox "Data kosong!,tidak dapat menyimpan data", vbCritical + vbSystemModal, "Informasi"
Else
    If MsgBox("Yakinkah data anda sudah terisi dengan benar?", vbOKCancel + vbQuestion, "Konfirmasi") = vbOK Then
        With clsClass
            For i = 1 To LVClass.ListItems.Count
                .kdClass = LVClass.ListItems.Item(i)
                    With LVClass.ListItems.Item(i)
                        clsClass.noflt = .ListSubItems.Item(2)
                        clsClass.simpan_kelas .ListSubItems(1), .ListSubItems(3), .ListSubItems(4), .ListSubItems(5), .ListSubItems(6), .ListSubItems(7), .ListSubItems(8)
                    End With
            Next
        End With
        bersihkanLayar
        LVClass.ListItems.Clear
        lblTarifRoute.Caption = ""
        txtKdKelas.SetFocus
    End If
End If
End Sub

Private Sub cmdSimpanMaskapai_Click()
Dim clsmaskapai As New ClassMaskapai
If kosong_box_maskapai Then
    MsgBox "Jadwal penerbangan mohon diisi", vbInformation + vbSystemModal, "Informasi"
Else
With clsmaskapai
    .simpanMaskapai txtflt.Text, cmbrouteMaskapai.Text, mskjadwal.Text
End With
'bersihkanLayar
'txtflt.Text = "RI"
'txtflt.SelStart = 2
'lblroute.Caption = ""
'ambilMaskapai
'txtflt.SetFocus
bersihkan_tabInputFlight 'panggil modul pembersih tab inputflight
End If
End Sub

Private Sub cmdsimpanRoute_Click()
If kosong_box Then
    MsgBox "isi data anda"
    txtKdRoute.SetFocus
Else
    With clsRoute
        .simpan_route Trim$(txtKdRoute.Text), HilangkanSpasiTengah(txtnmRoute.Text)
    End With
    bersihkanLayar
    ambilRoute_
    txtKdRoute.SetFocus
    
End If
End Sub

Function bersihkanLayar()
Dim ctrl As Control
For Each ctrl In Me.Controls
    If TypeOf ctrl Is TextBox Then
        ctrl.Text = ""
    End If
    If TypeOf ctrl Is MaskEdBox Then
        ctrl.Text = "__:__"
        ctrl.Mask = "&&:&&"
    End If
    If TypeOf ctrl Is ComboBox Then
        ctrl.Text = ""
    End If
    
Next
End Function
Function kosong_box() As Boolean
Dim ctrl As Control
Dim frm As Frame
If txtKdRoute.Text = "" Then
        kosong_box = True
ElseIf txtnmRoute.Text = "" Then
End If
End Function
Function kosong_box_maskapai() As Boolean
If txtflt.Text = "" Then
    kosong_box_maskapai = True
ElseIf cmbrouteMaskapai.Text = "" Then
    kosong_box_maskapai = True
ElseIf mskjadwal.Text = "__:__" Then
    kosong_box_maskapai = True
End If
End Function

Private Sub cmdUbahClass_Click()
Dim SetEnable As Boolean
Dim clsmaskapai As New ClassMaskapai
Dim xData() As String, xClass() As String


Dim i As Integer
On Error Resume Next
'xInput = InputBox("Masukkan No.Flt<spasi>Kd.Class", "Ubah Class")
'xData = Split(xInput, " ")


Select Case cmdUbahClass.Caption
        Case Is = "Cari"
                cmdUbahClass.Caption = "OK"
                EnableFalse_InputTarif False
        Case Is = "Ubah"
            clsmaskapai.ubah_noflt_class cmbNoflt.Text, _
                                                    txtKdKelas.Text, _
                                                    Val(txtbasic.Text), _
                                                    Val(txtFuel.Text), _
                                                    Val(txtppn.Text), _
                                                    Val(txtiwjr.Text), _
                                                    Val(txtselling.Text), _
                                                    Val(txtstock.Text), _
                                                    cmbNmClass.Text
            
             EnableFalse_InputTarif True
             bersihMaster_tarif
             cmdUbahClass.Caption = "Cari"
        Case Is = "OK"
                    xClass = clsmaskapai.cari_noflt_class(cmbNoflt.Text, txtKdKelas.Text)
                    For i = 0 To UBound(xClass)
                        xData = Split(xClass(i), "#")
                        On Error Resume Next
                        cmbNmClass.Text = xData(0)
                        txtbasic.Text = xData(1)
                        txtFuel.Text = xData(2)
                        txtppn.Text = xData(3)
                        txtiwjr.Text = xData(4)
                        txtselling.Text = xData(5)
                        txtstock.Text = xData(6)
                    Next
                    cmdUbahClass.Caption = "Ubah"
                   
End Select

End Sub

Private Sub cmdUbahMaskapai_Click()

On Error Resume Next 'JIKA DATA TIDAK ADA MAKA LIWATIN

If cmdUbahMaskapai.Caption = "Ubah" Then
'    cmbEditFlt.Clear
'    cData = clsMaskapai.panggilMaskapai
'    For j = 0 To UBound(cData)
'        bagicdata = Split(cData(j))
'        cmbEditFlt.AddItem (bagicdata(0))
'    Next
    ambilRoute_
    ambilNoflt
    OpsiUbahMaskapai True
    cmdHapusMaskapai.Enabled = True
    cmdUbahMaskapai.Caption = "OK"
Else
    If MsgBox("Data Flight '" & cmbrouteMaskapai.Text & "' akan diubah?", vbQuestion + vbYesNo, "Konfirmasi Ubah") = vbOK Then
        noflt = cmbEditFlt.Text
        clsmaskapai.ubah_maskapai noflt, cmbrouteMaskapai.Text, mskjadwal.Text
        bersihkan_tabInputFlight
        cmbEditFlt.Clear
        cmbrouteMaskapai.Clear: mskjadwal.Text = "__:__"
        mskjadwal.Mask = "&&:&&"
        mskjadwal.Width = 615
        OpsiUbahMaskapai False
        cmdUbahMaskapai.Caption = "Ubah"
    End If
End If
End Sub

Private Sub cmdubahRoute_Click()
Dim i As Integer
Dim xData() As String
If cmdubahRoute.Caption = "&Ubah" Then
        cmdubahRoute.Caption = "OK"
        cmdsimpanRoute.Enabled = False
        frmEditRoute.Width = frmsroute.Width
        frmEditRoute.Height = frmsroute.Height
        frmEditRoute.Left = frmsroute.Left
        frmEditRoute.Top = frmsroute.Top
        frmEditRoute.Visible = True
        
Else
     If MsgBox("Yakin,ingin mengubah Data Route?", vbOKCancel + vbQuestion, "Konfirmasi") = vbOK Then
        cmdubahRoute.Caption = "&Ubah"
        ubahRoute
        frmEditRoute.Visible = False
        ambilRoute_
        cmdsimpanRoute.Enabled = True
     End If
        
End If
End Sub



Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyEscape Then
        mskjadwal.Width = 615
        OpsiUbahMaskapai False
        cmdUbahMaskapai.Caption = "Ubah"
End If
End Sub

Private Sub Form_Load()
Call centerscreen(Me)
Me.MousePointer = vbHourglass
ambilRoute_
ambilMaskapai
Me.MousePointer = vbDefault
End Sub

Function ambilRoute_()
Dim clsRoute As New ClassRoute
Dim i As Integer
Dim xroute() As String, xbagi() As String
On Error Resume Next 'Jika data yang diambil tidak ada maka pass
LVRoute.ListItems.Clear
cmbrouteMaskapai.Clear
cmbUbahkdRoute.Clear
xroute = clsRoute.ambilRoute
For i = 0 To UBound(xroute)
        xbagi = Split(xroute(i), "#")
        With LVRoute.ListItems.Add(, , xbagi(0))
                cmbrouteMaskapai.AddItem (xbagi(0))
                cmbUbahkdRoute.AddItem (xbagi(0))
                .ListSubItems.Add , , xbagi(1)
        End With
Next
End Function
Function ambilMaskapai()
Dim clsmaskapai As New ClassMaskapai
Dim i As Integer, xMaskapai() As String, bagiMaskapai() As String
On Error Resume Next
LVMaskapai.ListItems.Clear
cmbNoflt.Clear
With clsmaskapai
    xMaskapai = .panggilMaskapai
    For i = 0 To UBound(xMaskapai)
            bagiMaskapai = Split(xMaskapai(i), "#")
            With LVMaskapai.ListItems.Add(, , bagiMaskapai(0)) 'no_flt
                    .ListSubItems.Add , , bagiMaskapai(1) 'kd_route
                    .ListSubItems.Add , , bagiMaskapai(2)
                     .ListSubItems.Add , , bagiMaskapai(3)
            End With
            cmbNoflt.AddItem (bagiMaskapai(0))
    Next
End With
End Function

Private Sub Text5_Change()

End Sub

Private Sub Text5_KeyPress(KeyAscii As Integer)

End Sub

Private Sub mskjadwal_KeyPress(KeyAscii As Integer)
If cekdigit(KeyAscii) = True Then
    KeyAscii = 0
ElseIf KeyAscii = 13 Then cmdSimpanMaskapai_Click
    
End If
End Sub

Private Sub txtbasic_KeyPress(KeyAscii As Integer)
If cekdigit(KeyAscii) = True Then
    KeyAscii = 0
ElseIf KeyAscii = 13 Then
    If Len(txtbasic.Text) = 0 Then
            txtbasic.Text = "0"
    End If
    txtFuel.SetFocus
End If
End Sub

Private Sub txtflt_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then cmbrouteMaskapai.SetFocus
End Sub

Private Sub txtFuel_KeyPress(KeyAscii As Integer)
If cekdigit(KeyAscii) = True Then
    KeyAscii = 0
ElseIf KeyAscii = 13 Then
    If Len(txtFuel.Text) = 0 Then
        txtFuel.Text = "0"
    End If
        txtppn.SetFocus
End If
End Sub

Private Sub txtiwjr_Change()
hitungTotalSP
End Sub

Private Sub txtiwjr_KeyPress(KeyAscii As Integer)
If cekdigit(KeyAscii) = True Then
    KeyAscii = 0
ElseIf KeyAscii = 13 Then
    If Len(txtiwjr.Text) = 0 Then
            txtiwjr.Text = "0"
    End If
    txtstock.SetFocus
End If
End Sub

Private Sub txtKdKelas_Change()
txtKdKelas.SelStart = Len(txtKdKelas.Text)
txtKdKelas.Text = UCase(txtKdKelas.Text)
End Sub

Private Sub txtKdKelas_KeyPress(KeyAscii As Integer)
If IsNumeric(Chr(KeyAscii)) = True Then
    KeyAscii = 0
ElseIf KeyAscii = 13 Then
    cmbNmClass.SetFocus
End If
End Sub

Private Sub txtKdRoute_Change()
txtKdRoute.SelStart = Len(txtKdRoute.Text)
txtKdRoute.Text = UCase(txtKdRoute.Text)
End Sub

Private Sub txtKdRoute_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then txtnmRoute.SetFocus
End Sub

Private Sub txtnmRoute_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then cmdsimpanRoute_Click
End Sub

Private Sub txtppn_Change()
hitungPajak
End Sub

Private Sub txtppn_KeyPress(KeyAscii As Integer)
If cekdigit(KeyAscii) = True Then
    KeyAscii = 0
ElseIf KeyAscii = 13 Then
    If Len(txtppn.Text) = 0 Then
        txtppn.Text = "0"
    End If
    txtppn2.SetFocus
End If
End Sub

Private Sub txtppn2_KeyPress(KeyAscii As Integer)
If cekdigit(KeyAscii) = True Then
    KeyAscii = 0
ElseIf KeyAscii = 13 Then
    txtiwjr.SetFocus
End If
End Sub

Private Sub txtselling_KeyPress(KeyAscii As Integer)
If cekdigit(KeyAscii) = True Then
    KeyAscii = 0
End If
End Sub

Private Sub txtstock_KeyPress(KeyAscii As Integer)
If cekdigit(KeyAscii) = True Then KeyAscii = 0

If KeyAscii = 13 Then
    If kosong_box_Class = True Then
            MsgBox "Mohon periksa kelengkapan data", vbInformation + vbSystemModal, "Informasi"
    Else
        With LVClass.ListItems.Add(, , txtKdKelas.Text)
            .ListSubItems.Add , , cmbNmClass.Text
            .ListSubItems.Add , , cmbNoflt.Text
            .ListSubItems.Add , , txtbasic.Text
            .ListSubItems.Add , , txtFuel.Text
            .ListSubItems.Add , , txtppn.Text
            .ListSubItems.Add , , txtiwjr.Text
            .ListSubItems.Add , , txtselling.Text
            .ListSubItems.Add , , txtstock.Text
        End With
        bersihMaster_tarif
    End If
    
End If
End Sub
Function hitungPajak()
txtppn2.Text = Val(txtbasic.Text) * Val(txtppn.Text) / 100
End Function
Function hitungTotalSP()
'Hitung jumlah Total keseluruhan
Dim fb As Long, fs As Long, ppn As Long, iwjr As Long, sp As Long

fb = Val(txtbasic.Text)
fs = Val(txtFuel.Text)
ppn = Val(txtppn2.Text)
iwjr = Val(txtiwjr.Text)
sp = fb + fs + ppn + iwjr

'txtsp.Text = (Val(txtiwjr.Text) + Val(txtfb.Text) + Val(txtfs.Text) + Val(txtppn2.Text))
txtselling.Text = sp
End Function
Function kosong_box_Class() As Boolean
If txtKdKelas.Text = "" Then
        kosong_box_Class = True
ElseIf cmbNmClass.Text = "" Then
        kosong_box_Class = True
ElseIf cmbNoflt.Text = "" Then
        kosong_box_Class = True
ElseIf txtbasic.Text = "" Then
        kosong_box_Class = True
ElseIf txtFuel.Text = "" Then
        kosong_box_Class = True
ElseIf txtppn2.Text = "" Then
        kosong_box_Class = True
ElseIf txtppn.Text = "" Then
        kosong_box_Class = True
ElseIf txtiwjr.Text = "" Then
        kosong_box_Class = True
ElseIf txtselling.Text = "" Then
        kosong_box_Class = True
ElseIf txtstock.Text = "" Then
        kosong_box_Class = True
'ElseIf LVClass.ListItems.Count = 0 Then
'        kosong_box_Class = True
End If
End Function
Function ubahRoute()

Dim clsRoute As New ClassRoute
clsRoute.EditRoute cmbUbahkdRoute.Text, txtUbahNmRoute.Text

End Function
Private Function OpsiUbahMaskapai(ByVal NonAktif As Boolean)
    cmbEditFlt.Left = 1440
    cmbEditFlt.Top = 1320
 If NonAktif Then
    cmbEditFlt.Visible = True
    txtflt.Visible = False
    cmdHapusMaskapai.Enabled = False
    cmdSimpanMaskapai.Enabled = False
 Else
    cmbEditFlt.Visible = False
    txtflt.Visible = True
    cmdHapusMaskapai.Enabled = True
    cmdSimpanMaskapai.Enabled = True
End If
End Function
Function ambilNoflt()
cmbEditFlt.Clear
        cData = clsmaskapai.panggilMaskapai
        For j = 0 To UBound(cData)
            bagicdata = Split(cData(j), "#")
            cmbEditFlt.AddItem (bagicdata(0))
        Next
End Function
Function bersihMaster_tarif()
txtstock.Text = ""

txtiwjr.Text = ""
txtppn2.Text = ""
txtppn.Text = ""
txtFuel.Text = ""
txtbasic.Text = ""
cmbNoflt.Text = ""
txtKdKelas.Text = ""
cmbNmClass.Text = ""
lblTarifRoute.Caption = ""
txtselling.Text = ""

txtKdKelas.SetFocus
End Function
Function HilangkanSpasiTengah(ByVal xkarakter As String) As String
Dim hasil As String
Dim tmp() As String, tmp2() As String
Dim i As Integer, j As Integer
ReDim Preserve tmp2(i)
tmp2(i) = xkarakter
tmp = Split(tmp2(i), " ")
For j = 0 To UBound(tmp)
    hasil = hasil & Trim$(tmp(j))
Next j
HilangkanSpasiTengah = hasil
End Function
Function EnableFalse_InputTarif(ByVal sTaTus As Boolean)
If sTaTus = False Then
        cmdSimpanClass.Enabled = False
'        cmdUbahClass.Enabled = False
        cmdHapusClass.Enabled = False
Else
        cmdSimpanClass.Enabled = True
'        cmdUbahClass.Enabled = True
        cmdHapusClass.Enabled = True
End If
End Function
Function xInputBox(ByVal noflt As String, ByVal kdClass As String)
Dim inputan As String
Dim xbagi() As String
inputan = InputBox("Masukkan No.Flt<spasi>KdClass", "Update Class")
xbagi = Split(inputan, " ")
noflt = xbagi(0)
kdClass = xbagi(1)
End Function
Function bersihkan_tabInputFlight()
bersihkanLayar
txtflt.Text = "RI"
txtflt.SelStart = 2
lblroute.Caption = ""
ambilMaskapai
'txtflt.SetFocus
cmdUbahMaskapai.Caption = "Ubah"
cmdHapusMaskapai.Caption = "Hapus"
cmdSimpanMaskapai.Enabled = True
cmbEditFlt.Visible = False
txtflt.Visible = True
End Function

