VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ClassPass"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Dim rsPass As ADODB.Recordset
Dim sUser As String, spass As String
Dim passBenar As Boolean
Public Function simpan_pass(ByVal username As String, ByVal pass As String)
Set rsPass = New ADODB.Recordset
rsPass.Open "SELECT * FROM xpass WHERE xusername LIKE '" & username & "'", cnAdmin, 1, 3
With rsPass
    If .RecordCount <> 0 Then
            MsgBox "Username dan Password anda sudah terpakai", vbInformation + vbSystemModal, "Informasi"
    Else
            .AddNew
                    !xusername = username
                    !xpassword = xCrypt(pass)
            .Update
    End If
End With
rsPass.Close
cnAdmin.Close
End Function
Public Function cek_password(ByVal username As String, Optional ByVal pl As String)
Dim i As Integer
Dim passLama As String
passAda = False
Set rsPass = New ADODB.Recordset
rsPass.Open "SELECT * FROM xpass WHERE xusername LIKE '" & username & "'", cnAdmin, 1, 3
    With rsPass
            If .RecordCount = 0 Then
                    MsgBox "Username dan Password tidak ditemukan", vbInformation + vbSystemModal, "Informasi"
            Else
                If pl = xDesc(!xpassword) Then
                    xuser = xDesc(!xpassword)
                    passAda = True
                Else
                    MsgBox "Username dan Password tidak ditemukan"
                End If
            End If
    End With
rsPass.Close
cnAdmin.Close
End Function
Public Function xCrypt(ByVal karakter As String) As String
Dim x
Dim i As Integer
Dim newPass As String, Encrypt As String
'Encrypt = ""
newPass = Len(karakter)
x = 1
Do While x <= newPass
    Encrypt = CStr(Encrypt) + Chr((Asc(Mid(Trim(karakter), x, 1)) + 17))
    x = x + 1
Loop
xCrypt = Encrypt
End Function
Public Function xDesc(ByVal karakter As String) As String
Dim x
Dim i As Integer
Dim oldPass As String, Desc As String
oldPass = Len(karakter)
x = 1
Do While x <= oldPass
    Desc = Desc + Chr((Asc(Mid(Trim(karakter), x, 1)) - 17))
    x = x + 1
Loop
xDesc = Desc
End Function

Public Property Get xuser() As String
xuser = sUser
End Property

Public Property Let xuser(ByVal vNewValue As String)
sUser = vNewValue
End Property

Public Function update_pass(ByVal xuser As String, ByVal xpass As String, Optional ByVal xnewpass As String, Optional ByVal opsi As Integer)
Set rsPass = New ADODB.Recordset
rsPass.Open "SELECT * FROM xpass WHERE xusername LIKE '" & xuser & "'", cnAdmin, 1, 3
With rsPass
    If .RecordCount <> 0 Then
        If xpass = xDesc(!xpassword) Then
            Select Case opsi
                    Case 1 'ubah
                        !xpassword = xCrypt(xnewpass) 'Encrypt passwordnya
                        .Update
                    Case 2 'delete
                            Do While Not .EOF
                                .Delete
                                .MoveNext
                            Loop
            End Select
         Else
            MsgBox "data tidak ditemukan"
        End If
    End If
End With
rsPass.Close
cnAdmin.Close
End Function

Public Property Get passAda() As Boolean
passAda = passBenar
End Property

Public Property Let passAda(ByVal vNewValue As Boolean)
passBenar = vNewValue
End Property
