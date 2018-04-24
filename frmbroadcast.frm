VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{531653E6-69A2-40A4-8734-91C78B33DD08}#1.0#0"; "XPButton.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmbroadcast 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Broadcast"
   ClientHeight    =   5040
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   7065
   Icon            =   "frmbroadcast.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   5040
   ScaleWidth      =   7065
   ShowInTaskbar   =   0   'False
   Begin MSComctlLib.ImageList Imagelist1 
      Left            =   6000
      Top             =   3480
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   61
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmbroadcast.frx":0E42
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmbroadcast.frx":115C
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmbroadcast.frx":1476
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmbroadcast.frx":3C28
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmbroadcast.frx":5932
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmbroadcast.frx":5C4C
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmbroadcast.frx":5DA6
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmbroadcast.frx":5F00
            Key             =   ""
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmbroadcast.frx":605A
            Key             =   ""
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmbroadcast.frx":6374
            Key             =   ""
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmbroadcast.frx":71C6
            Key             =   ""
         EndProperty
         BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmbroadcast.frx":8018
            Key             =   ""
         EndProperty
         BeginProperty ListImage13 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmbroadcast.frx":8EC2
            Key             =   ""
         EndProperty
         BeginProperty ListImage14 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmbroadcast.frx":9D14
            Key             =   ""
         EndProperty
         BeginProperty ListImage15 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmbroadcast.frx":C4C6
            Key             =   ""
         EndProperty
         BeginProperty ListImage16 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmbroadcast.frx":C7E0
            Key             =   ""
         EndProperty
         BeginProperty ListImage17 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmbroadcast.frx":D632
            Key             =   ""
         EndProperty
         BeginProperty ListImage18 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmbroadcast.frx":D94C
            Key             =   ""
         EndProperty
         BeginProperty ListImage19 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmbroadcast.frx":EFA6
            Key             =   ""
         EndProperty
         BeginProperty ListImage20 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmbroadcast.frx":F2C0
            Key             =   ""
         EndProperty
         BeginProperty ListImage21 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmbroadcast.frx":F5DA
            Key             =   ""
         EndProperty
         BeginProperty ListImage22 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmbroadcast.frx":F8F4
            Key             =   ""
         EndProperty
         BeginProperty ListImage23 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmbroadcast.frx":FC0E
            Key             =   ""
         EndProperty
         BeginProperty ListImage24 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmbroadcast.frx":104E8
            Key             =   ""
         EndProperty
         BeginProperty ListImage25 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmbroadcast.frx":10DC2
            Key             =   ""
         EndProperty
         BeginProperty ListImage26 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmbroadcast.frx":11C14
            Key             =   ""
         EndProperty
         BeginProperty ListImage27 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmbroadcast.frx":12A66
            Key             =   ""
         EndProperty
         BeginProperty ListImage28 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmbroadcast.frx":138B8
            Key             =   ""
         EndProperty
         BeginProperty ListImage29 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmbroadcast.frx":14192
            Key             =   ""
         EndProperty
         BeginProperty ListImage30 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmbroadcast.frx":14FE4
            Key             =   ""
         EndProperty
         BeginProperty ListImage31 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmbroadcast.frx":16CEE
            Key             =   ""
         EndProperty
         BeginProperty ListImage32 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmbroadcast.frx":17008
            Key             =   ""
         EndProperty
         BeginProperty ListImage33 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmbroadcast.frx":17E5A
            Key             =   ""
         EndProperty
         BeginProperty ListImage34 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmbroadcast.frx":1A60C
            Key             =   ""
         EndProperty
         BeginProperty ListImage35 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmbroadcast.frx":1CDBE
            Key             =   ""
         EndProperty
         BeginProperty ListImage36 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmbroadcast.frx":1EAC8
            Key             =   ""
         EndProperty
         BeginProperty ListImage37 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmbroadcast.frx":2127A
            Key             =   ""
         EndProperty
         BeginProperty ListImage38 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmbroadcast.frx":22F84
            Key             =   ""
         EndProperty
         BeginProperty ListImage39 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmbroadcast.frx":2385E
            Key             =   ""
         EndProperty
         BeginProperty ListImage40 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmbroadcast.frx":246B0
            Key             =   ""
         EndProperty
         BeginProperty ListImage41 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmbroadcast.frx":25D0A
            Key             =   ""
         EndProperty
         BeginProperty ListImage42 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmbroadcast.frx":265E4
            Key             =   ""
         EndProperty
         BeginProperty ListImage43 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmbroadcast.frx":26EBE
            Key             =   ""
         EndProperty
         BeginProperty ListImage44 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmbroadcast.frx":27D10
            Key             =   ""
         EndProperty
         BeginProperty ListImage45 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmbroadcast.frx":2936A
            Key             =   ""
         EndProperty
         BeginProperty ListImage46 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmbroadcast.frx":2A1BC
            Key             =   ""
         EndProperty
         BeginProperty ListImage47 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmbroadcast.frx":2A4D6
            Key             =   ""
         EndProperty
         BeginProperty ListImage48 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmbroadcast.frx":2ADB0
            Key             =   ""
         EndProperty
         BeginProperty ListImage49 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmbroadcast.frx":2B0CA
            Key             =   ""
         EndProperty
         BeginProperty ListImage50 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmbroadcast.frx":2B3E4
            Key             =   ""
         EndProperty
         BeginProperty ListImage51 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmbroadcast.frx":2B6FE
            Key             =   ""
         EndProperty
         BeginProperty ListImage52 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmbroadcast.frx":2B858
            Key             =   ""
         EndProperty
         BeginProperty ListImage53 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmbroadcast.frx":2BB72
            Key             =   ""
         EndProperty
         BeginProperty ListImage54 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmbroadcast.frx":2BE8C
            Key             =   ""
         EndProperty
         BeginProperty ListImage55 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmbroadcast.frx":2C1A6
            Key             =   ""
         EndProperty
         BeginProperty ListImage56 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmbroadcast.frx":2C4C0
            Key             =   ""
         EndProperty
         BeginProperty ListImage57 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmbroadcast.frx":2C7DA
            Key             =   ""
         EndProperty
         BeginProperty ListImage58 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmbroadcast.frx":2CAF4
            Key             =   ""
         EndProperty
         BeginProperty ListImage59 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmbroadcast.frx":2CC4E
            Key             =   ""
         EndProperty
         BeginProperty ListImage60 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmbroadcast.frx":2CDA8
            Key             =   ""
         EndProperty
         BeginProperty ListImage61 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmbroadcast.frx":2CF02
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin TabDlg.SSTab tbMenu 
      Height          =   4935
      Left            =   120
      TabIndex        =   0
      Top             =   0
      Width           =   6855
      _ExtentX        =   12091
      _ExtentY        =   8705
      _Version        =   393216
      Tabs            =   2
      TabHeight       =   520
      TabCaption(0)   =   "Broadcast"
      TabPicture(0)   =   "frmbroadcast.frx":2D05C
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "Shape1"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "cmdTambah"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "cmdclear"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "cmdBatal"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).Control(4)=   "cmdkirim"
      Tab(0).Control(4).Enabled=   0   'False
      Tab(0).Control(5)=   "Frame2"
      Tab(0).Control(5).Enabled=   0   'False
      Tab(0).ControlCount=   6
      TabCaption(1)   =   "Tabel Pelanggan"
      TabPicture(1)   =   "frmbroadcast.frx":2D078
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "Frame3"
      Tab(1).Control(1)=   "lvpelanggan"
      Tab(1).ControlCount=   2
      Begin VB.Frame Frame3 
         Caption         =   "Pilih Opsi"
         Height          =   1095
         Left            =   -74760
         TabIndex        =   18
         Top             =   3600
         Width           =   2775
         Begin VB.OptionButton optPilih 
            Caption         =   "Ubah"
            Height          =   375
            Index           =   1
            Left            =   240
            TabIndex        =   20
            Top             =   480
            Width           =   1095
         End
         Begin VB.OptionButton optPilih 
            Caption         =   "Hapus"
            Height          =   375
            Index           =   0
            Left            =   240
            TabIndex        =   19
            Top             =   240
            Width           =   1095
         End
         Begin ucXPButton.XPButton cmdUpdate 
            Height          =   375
            Left            =   1920
            TabIndex        =   21
            ToolTipText     =   "Tambah anggota baru"
            Top             =   600
            Width           =   735
            _ExtentX        =   1296
            _ExtentY        =   661
            BTYPE           =   5
            TX              =   "&OK"
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
            MICON           =   "frmbroadcast.frx":2D094
            PICN            =   "frmbroadcast.frx":2D0B0
            UMCOL           =   -1  'True
            SOFT            =   0   'False
            PICPOS          =   0
            NGREY           =   0   'False
            FX              =   0
            HAND            =   0   'False
            CHECK           =   0   'False
            VALUE           =   0   'False
         End
      End
      Begin VB.Frame Frame2 
         BorderStyle     =   0  'None
         Height          =   3735
         Left            =   120
         TabIndex        =   1
         Top             =   360
         Width           =   6615
         Begin VB.Frame framBroad 
            Height          =   3615
            Left            =   0
            TabIndex        =   2
            Top             =   0
            Width           =   6495
            Begin VB.Frame framPuts 
               Height          =   2055
               Left            =   3480
               TabIndex        =   3
               Top             =   240
               Width           =   2895
               Begin VB.TextBox txtno 
                  Enabled         =   0   'False
                  Height          =   318
                  Left            =   120
                  MaxLength       =   15
                  TabIndex        =   6
                  Text            =   "62"
                  ToolTipText     =   "Input nomor tambahan"
                  Top             =   1560
                  Width           =   1455
               End
               Begin VB.TextBox txtNm 
                  Enabled         =   0   'False
                  Height          =   318
                  Left            =   120
                  MaxLength       =   15
                  TabIndex        =   5
                  ToolTipText     =   "Input nomor tambahan"
                  Top             =   360
                  Width           =   1335
               End
               Begin VB.TextBox txtAlamat 
                  Enabled         =   0   'False
                  Height          =   315
                  Left            =   120
                  MaxLength       =   15
                  TabIndex        =   4
                  ToolTipText     =   "Input nomor tambahan"
                  Top             =   960
                  Width           =   2535
               End
               Begin VB.Label Label2 
                  BackColor       =   &H00C0C0C0&
                  BackStyle       =   0  'Transparent
                  Caption         =   "Nama:"
                  Height          =   255
                  Left            =   120
                  TabIndex        =   9
                  Top             =   120
                  Width           =   855
               End
               Begin VB.Label Label3 
                  BackColor       =   &H00C0C0C0&
                  BackStyle       =   0  'Transparent
                  Caption         =   "Alamat :"
                  Height          =   255
                  Left            =   120
                  TabIndex        =   8
                  Top             =   720
                  Width           =   855
               End
               Begin VB.Label Label4 
                  BackColor       =   &H00C0C0C0&
                  BackStyle       =   0  'Transparent
                  Caption         =   "Hp :"
                  Height          =   255
                  Left            =   120
                  TabIndex        =   7
                  Top             =   1320
                  Width           =   855
               End
            End
            Begin VB.Frame Frame1 
               Caption         =   "Pesan  :"
               Height          =   1095
               Left            =   120
               TabIndex        =   10
               Top             =   2400
               Width           =   6255
               Begin VB.TextBox txtPesan 
                  Height          =   735
                  Left            =   120
                  MaxLength       =   159
                  MultiLine       =   -1  'True
                  TabIndex        =   11
                  ToolTipText     =   "Ketik pesan sms anda disini"
                  Top             =   240
                  Width           =   6015
               End
            End
            Begin MSComctlLib.ListView lv 
               Height          =   2175
               Left            =   120
               TabIndex        =   12
               Top             =   240
               Width           =   3255
               _ExtentX        =   5741
               _ExtentY        =   3836
               View            =   3
               LabelWrap       =   -1  'True
               HideSelection   =   -1  'True
               Checkboxes      =   -1  'True
               FullRowSelect   =   -1  'True
               GridLines       =   -1  'True
               _Version        =   393217
               SmallIcons      =   "Imagelist1"
               ForeColor       =   -2147483640
               BackColor       =   -2147483643
               BorderStyle     =   1
               Appearance      =   1
               NumItems        =   2
               BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
                  Text            =   "Hp"
                  Object.Width           =   3528
               EndProperty
               BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
                  SubItemIndex    =   1
                  Text            =   "Nama"
                  Object.Width           =   5292
               EndProperty
            End
         End
      End
      Begin ucXPButton.XPButton cmdkirim 
         Height          =   375
         Left            =   600
         TabIndex        =   13
         ToolTipText     =   "Kirim SMS ke semua anggota"
         Top             =   4200
         Width           =   1095
         _ExtentX        =   1931
         _ExtentY        =   661
         BTYPE           =   5
         TX              =   "&Kirim"
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
         MICON           =   "frmbroadcast.frx":2D64A
         PICN            =   "frmbroadcast.frx":2D666
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin ucXPButton.XPButton cmdBatal 
         Height          =   375
         Left            =   4920
         TabIndex        =   14
         ToolTipText     =   "Batalkan pengiriman"
         Top             =   4200
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
         BCOL            =   13160660
         BCOLO           =   13160660
         FCOL            =   0
         FCOLO           =   0
         MCOL            =   12632256
         MPTR            =   1
         MICON           =   "frmbroadcast.frx":2DC00
         PICN            =   "frmbroadcast.frx":2DC1C
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin ucXPButton.XPButton cmdclear 
         Height          =   375
         Left            =   2040
         TabIndex        =   15
         ToolTipText     =   "Bersihkan layar ini"
         Top             =   4200
         Width           =   1095
         _ExtentX        =   1931
         _ExtentY        =   661
         BTYPE           =   5
         TX              =   "&Clear"
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
         MICON           =   "frmbroadcast.frx":39056
         PICN            =   "frmbroadcast.frx":39072
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin ucXPButton.XPButton cmdTambah 
         Height          =   375
         Left            =   3480
         TabIndex        =   16
         ToolTipText     =   "Tambah anggota baru"
         Top             =   4200
         Width           =   1095
         _ExtentX        =   1931
         _ExtentY        =   661
         BTYPE           =   5
         TX              =   "&Tambah"
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
         MICON           =   "frmbroadcast.frx":5660C
         PICN            =   "frmbroadcast.frx":56628
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin MSComctlLib.ListView lvpelanggan 
         Height          =   2895
         Left            =   -74760
         TabIndex        =   17
         Top             =   480
         Width           =   6375
         _ExtentX        =   11245
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
         NumItems        =   3
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "Hp"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "Nama"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   2
            Text            =   "Alamat"
            Object.Width           =   6174
         EndProperty
      End
      Begin VB.Shape Shape1 
         BackColor       =   &H80000014&
         BackStyle       =   1  'Opaque
         Height          =   615
         Left            =   120
         Top             =   4080
         Width           =   6495
      End
   End
End
Attribute VB_Name = "frmbroadcast"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Option Base 1
Dim i As Integer
Dim clsPlg As New ClassPelanggan

Private Sub cmdBatal_Click()
'status_konek = True
If cmdTambah.Caption = "&Simpan" Then
     enabled_tombol False
     cmdTambah.Caption = "&Tambah"
Else
    Unload Me
End If
End Sub

Private Sub cmdclear_Click()
Call bersihkan_TabBroadcast
End Sub

Private Sub cmdkirim_Click()
If Cek_status_tabel = True Then
    If utama.com.PortOpen = False Then
        MsgBox "Pastikan kabel data anda konek ke komputer", vbCritical + vbSystemModal, "Koneksi gagal"
    Else
        BroadCast
    End If
Else
    MsgBox "Tidak dapat melakukan Broadcast,pilih nomor hp yg ada ditabel", vbInformation + vbSystemModal, "Informasi" '
End If
End Sub

Private Sub cmdTambah_Click()
Dim nohp As String, xnama As String, xalamat As String
Dim clsPlg As New ClassPelanggan

If cmdTambah.Caption = "&Tambah" Then
    cmdTambah.Caption = "&Simpan"
    enabled_tombol True
    cmdBatal.Enabled = True
    txtNm.SetFocus
Else
    With clsPlg
            .simpan_pelanggan_manual txtno.Text, txtNm.Text, txtAlamat.Text
        End With
    Call bersihkan_TabBroadcast
    cmdTambah.Caption = "&Tambah"
    enabled_tombol False
    panggilData
    defaultIcon
End If
End Sub

Private Sub cmdUpdate_Click()
Dim no_telp As String, nmplg As String, almtplg As String
Dim clsPlg As New ClassPelanggan
Dim j As Integer
If optPilih.Item(0).Value = True Then
            no_telp = InputBox("Masukkan nomor Hp pelanggan", "Hapus data pelanggan")
            If Len(no_telp) = 0 Then Exit Sub
            clsPlg.hapus_nomor_pelanggan no_telp, , , 1
            
ElseIf optPilih.Item(1).Value = True Then
    no_telp = InputBox("Masukkan nomor Hp pelanggan", "Hapus data pelanggan")
    If Len(no_telp) = 0 Then Exit Sub
    nmplg = InputBox("Nama Pelanggan", "Ubah Nama")
    almtplg = InputBox("Alamat Pelanggan", "Ubah Alamat")
    clsPlg.hapus_nomor_pelanggan no_telp, nmplg, almtplg, 2
End If
End Sub
Private Sub Form_Load()
Call centerscreen(Me)
panggilData
 defaultIcon
End Sub
Function panggilData()
Dim i As Integer
Dim s() As String, bagiData() As String
On Error Resume Next
s = clsPlg.AddDataPelanggan
lv.ListItems.Clear
For i = 0 To UBound(s)
    bagiData = Split(s(i), "#")
    With lv.ListItems.Add(, , bagiData(0))
        .ListSubItems.Add , , bagiData(1)
    End With
Next
End Function
Private Sub lv_ItemCheck(ByVal Item As MSComctlLib.ListItem)
Set lv.SelectedItem = Item
If Item.Checked Then
    lv.ListItems.Item(lv.SelectedItem.Index).SmallIcon = 6
Else
    lv.ListItems.Item(lv.SelectedItem.Index).SmallIcon = 7
    lv.ListItems.Item(lv.SelectedItem.Index).Bold = True
End If
    
End Sub
Function defaultIcon()
Dim i As Integer
For i = 1 To lv.ListItems.Count
    lv.ListItems.Item(i).SmallIcon = 7
Next
End Function
Function BroadCast()
Dim i As Integer, x As Integer
Dim smsSend As String

If Len(txtPesan.Text) = 0 Then
        MsgBox "Isi pesan broadcast", vbInformation + vbSystemModal, "Informasi"
Else

        frmkirim.Show
        ReDim arrKirim(1)
        arrKirim(1) = ""
        bagiKirim (Trim(txtPesan.Text))
        status_konek = False: utama.Timer1 = False 'tunda auto respons
        status_signal = False: utama.Timer2 = False 'tunda pengecekan signal
        hp_konek = False: err_konek = False
        
         For i = 1 To lv.ListItems.Count
                
           If lv.ListItems.Item(i).Checked = True Then
                  For x = 1 To UBound(arrKirim)
                      If Trim(arrKirim(x)) <> "" Then
                            SMSDibalas lv.ListItems(i), arrKirim(x)
                            frmkirim.pbar1.Value = frmkirim.pbar1.Value + i
                            Do While Not hp_konek = True
                                    DoEvents
                                    Tunggu
                                    If err_konek = True Then Exit Do
                            Loop
                      End If
                    Next
           End If
        Next
        frmkirim.pbar1.Value = frmkirim.pbar1.Max
        MsgBox "Pesan telah terkirim", vbInformation + vbSystemModal, "Laporan sms"
        Tunda 0.1
        status_konek = True: utama.Timer1 = True 'mulai auto respons
        status_signal = True: utama.Timer2 = True 'lanjut pengecekan signal
End If
End Function

Private Sub txtalmt_KeyPress(KeyAscii As Integer)

End Sub


Private Sub XPButton1_Click()

End Sub

Private Sub tbMenu_Click(PreviousTab As Integer)
Dim i As Integer
Dim s() As String, bagiData() As String
On Error Resume Next
'MsgBox PreviousTab
If tbMenu.Tab = 1 Then
        s = clsPlg.AddDataPelanggan
        lvpelanggan.ListItems.Clear
        For i = 0 To UBound(s)
        bagiData = Split(s(i), "#")
        With lvpelanggan.ListItems.Add(, , bagiData(0))
            .ListSubItems.Add , , bagiData(1)
            .ListSubItems.Add , , bagiData(2)
        End With
        Next
ElseIf tbMenu.Tab = 0 Then
        panggilData
        defaultIcon
End If
End Sub

Private Sub txtAlamat_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then txtno.SetFocus
End Sub

Private Sub txtNm_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then txtAlamat.SetFocus
End Sub

Private Sub txtno_KeyPress(KeyAscii As Integer)
If cekdigit(KeyAscii) = True Then
    KeyAscii = 0
ElseIf KeyAscii = 13 Then
       If Len(txtno.Text) = 2 Then
        MsgBox "Mohon isi No.hp"
       Else
         cmdTambah_Click
         lv.ListItems.Add , , txtno.Text
            
       
      End If

End If
End Sub

Private Sub txtPesan_KeyPress(KeyAscii As Integer)
Dim p As String
i = Len(txtPesan.Text) + 1
p = i & " /150 Char"
i = i + 1
Frame1.Caption = "Pesan :" & p
End Sub
Function Cek_status_tabel() As Boolean
Dim i As Integer
    For i = 1 To lv.ListItems.Count
        If lv.ListItems.Item(i).Checked = True Then
            Cek_status_tabel = True
        End If
    Next
End Function
Function UncekAll() As Boolean
Dim i As Integer
For i = 1 To lv.ListItems.Count
    If lv.ListItems.Item(i).Checked = True Then
            lv.ListItems.Item(i).Checked = False
    End If
Next
End Function
Function enabled_tombol(ByVal eStatus As Boolean)
If eStatus Then
    cmdkirim.Enabled = False
    cmdclear.Enabled = False
    cmdBatal.Enabled = False
    txtNm.Enabled = True
    txtAlamat.Enabled = True
    txtno.Enabled = True
Else
    cmdkirim.Enabled = True
    cmdclear.Enabled = True
    cmdBatal.Enabled = True
     txtNm.Enabled = False
    txtAlamat.Enabled = False
    txtno.Enabled = False
End If
End Function
Function bersihkan_TabBroadcast()
Call clearscreen(Me)
UncekAll
txtno.Text = "62"
Frame1.Caption = "Pesan  :"
End Function
