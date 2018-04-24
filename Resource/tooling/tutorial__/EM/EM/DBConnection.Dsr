VERSION 5.00
Begin {C0E45035-5775-11D0-B388-00A0C9055D8E} DBConnection 
   ClientHeight    =   11130
   ClientLeft      =   0
   ClientTop       =   390
   ClientWidth     =   15360
   _ExtentX        =   13891
   _ExtentY        =   12197
   FolderFlags     =   1
   TypeLibGuid     =   "{D3A61422-B2B9-11D8-81CC-08002BE6B944}"
   TypeInfoGuid    =   "{D3A61423-B2B9-11D8-81CC-08002BE6B944}"
   TypeInfoCookie  =   0
   Version         =   4
   NumConnections  =   1
   BeginProperty Connection1 
      ConnectionName  =   "Koneksi1"
      ConnDispId      =   1001
      SourceOfData    =   3
      ConnectionSource=   "Provider=MSDASQL.1;Persist Security Info=False;Data Source=SMSAR"
      DesignSaveAuth  =   -1  'True
      Expanded        =   -1  'True
      QuoteChar       =   96
      SeparatorChar   =   46
   EndProperty
   NumRecordsets   =   1
   BeginProperty Recordset1 
      CommandName     =   "sqlPesan"
      CommDispId      =   1038
      RsDispId        =   1041
      CommandText     =   "select * from temppesan"
      ActiveConnectionName=   "Koneksi1"
      CommandType     =   1
      Expanded        =   -1  'True
      IsRSReturning   =   -1  'True
      NumFields       =   3
      BeginProperty Field1 
         Precision       =   0
         Size            =   20
         Scale           =   0
         Type            =   200
         Name            =   "pengirim"
         Caption         =   "pengirim"
      EndProperty
      BeginProperty Field2 
         Precision       =   19
         Size            =   16
         Scale           =   0
         Type            =   135
         Name            =   "waktu"
         Caption         =   "waktu"
      EndProperty
      BeginProperty Field3 
         Precision       =   0
         Size            =   160
         Scale           =   0
         Type            =   200
         Name            =   "teks"
         Caption         =   "teks"
      EndProperty
      NumGroups       =   0
      ParamCount      =   0
      RelationCount   =   0
      AggregateCount  =   0
   EndProperty
End
Attribute VB_Name = "DBConnection"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub DataEnvironment_Initialize()
On Error GoTo KoneksiErr

Koneksi1.ConnectionString = "DSN=SMSAR"
Koneksi1.Open

Exit Sub
   
KoneksiErr:
    TampilkanPesan "Koneksi ke Database Server tidak terlaksana" & vbCrLf & _
                    "Pesan: " & Err.Description & vbCrLf & _
                    "Kode: " & Err.Number
    End
End Sub
