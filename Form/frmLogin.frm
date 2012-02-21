VERSION 5.00
Begin VB.Form frmLogin 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "调查问卷信息处理系统－登录"
   ClientHeight    =   3135
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4485
   Icon            =   "frmLogin.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "frmLogin.frx":0E42
   ScaleHeight     =   209
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   299
   StartUpPosition =   2  '屏幕中心
   Begin VB.CommandButton Command1 
      Caption         =   "Command1"
      Height          =   390
      Left            =   225
      TabIndex        =   6
      Top             =   2655
      Visible         =   0   'False
      Width           =   675
   End
   Begin 调查问卷信息查询系统.FTextBox FTextBox1 
      Height          =   300
      Left            =   1440
      TabIndex        =   1
      Top             =   1440
      Width           =   2295
      _ExtentX        =   4048
      _ExtentY        =   529
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "宋体"
         Size            =   9
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      FontName        =   "宋体"
      FontSize        =   9
   End
   Begin 调查问卷信息查询系统.FTextBox txtPW 
      Height          =   300
      Left            =   1440
      TabIndex        =   3
      Top             =   1920
      Width           =   2295
      _ExtentX        =   4048
      _ExtentY        =   529
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "宋体"
         Size            =   9
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      FontName        =   "宋体"
      FontSize        =   9
      PasswordChar    =   "*"
   End
   Begin 调查问卷信息查询系统.XPButton cmdOK 
      Height          =   375
      Left            =   1800
      TabIndex        =   4
      Top             =   2625
      Width           =   1095
      _ExtentX        =   1931
      _ExtentY        =   661
      Caption         =   "登录(&L)"
      CapAlign        =   2
      BackStyle       =   2
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "宋体"
         Size            =   9
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      cGradient       =   0
      Mode            =   0
      Value           =   0   'False
      cBack           =   -2147483633
   End
   Begin 调查问卷信息查询系统.XPButton cmdExit 
      Height          =   375
      Left            =   3120
      TabIndex        =   5
      Top             =   2640
      Width           =   1095
      _ExtentX        =   1931
      _ExtentY        =   661
      Caption         =   "退出(&Q)"
      CapAlign        =   2
      BackStyle       =   2
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "宋体"
         Size            =   9
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      cGradient       =   0
      Mode            =   0
      Value           =   0   'False
      cBack           =   -2147483633
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "密  码："
      Height          =   180
      Left            =   600
      TabIndex        =   2
      Top             =   1980
      Width           =   720
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "用户名："
      Height          =   180
      Left            =   600
      TabIndex        =   0
      Top             =   1500
      Width           =   720
   End
   Begin VB.Shape Shape1 
      BackColor       =   &H00FFF8F0&
      BorderColor     =   &H00C5742F&
      Height          =   1335
      Left            =   270
      Top             =   1140
      Width           =   3900
   End
End
Attribute VB_Name = "frmLogin"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub cmdExit_Click()
    Unload Me
End Sub
Private Sub cmdOK_Click()
    If FTextBox1.Text = "" Then
        MsgBox "请填写用户名。", vbInformation
        FTextBox1.SetFocus
        Exit Sub
    End If
    If txtPW.Text = "" Then
        MsgBox "请填写密码。", vbInformation
        txtPW.SetFocus
        Exit Sub
    End If
'On Error GoTo aaaa
    Dim rs As New ADODB.Recordset, strMD5 As String
    If Conn.State <> 0 Then Conn.Close
    DBConnect
    rs.Open "Select * From userInfo Where uID='" & FTextBox1.Text & "'", Conn, 1, 1
    If Not rs.EOF Then
            If StrComp(rs("uID"), FTextBox1.Text, 1) = 0 And StrComp(rs("uPWD"), GetMD5(txtPW.Text), 1) = 0 Then
                curID = rs("uID")
                curSchool = rs("uSchool")
                curClass = rs("uClass")
                'SaveUserList
                frmMain.Icon = Me.Icon
                
                If Mid(curID, 2, 1) = "1" Then
                    If Mid(curID, 5, 1) < 4 Then
                        curInputForm = 1      '小学1-3年级
                    Else
                        curInputForm = 2     '小学4-6年级
                    End If
                Else
                        curInputForm = 3     '初中、高中
                End If
                
                Unload Me
                frmMain.Show
                Exit Sub
            End If
    End If
    MsgBox "用户名或密码错误，登陆失败！", vbCritical
    rs.Close
    Conn.Close
Exit Sub
aaaa:
    MsgBox Err.Description, vbCritical
    If Conn.State = 1 Then Conn.Close
End Sub

Private Sub cmdServer_Click()
    With frmServer
        .txtServer.Text = strSQLServer
        .txtUser.Text = strSQLUser
        If strSQLPW <> "" Then .lbPW.Visible = True
        .txtDB.Text = IIf(strSQLDB <> "", strSQLDB, "SuperMarketdb")
        .Show 1
    End With
End Sub

Private Sub Command1_Click()
    Dim rs As ADODB.Recordset
    Set rs = New ADODB.Recordset
    DBConnect
    rs.Open "select * from zdfx order by uid", Conn, 1, 1
    Do While Not rs.EOF
        sql = "insert into userinfo(uid,uPWD,uDesc,uSchool,uClass) values('" & _
                       rs("uid") & "','" & rs("uPWD") & "','" & rs("uDesc") & _
                       "','" & rs("uSchool") & "','" & rs("uClass") & "')"
        Conn.Execute sql
        rs.MoveNext
        
    Loop
    rs.Close
    Exit Sub
    
    
    rs.Open "select * from zdfx", Conn, 3, 3
    n = 0
    Do While Not rs.EOF
        temp = getPWD
        n = n + 1
        rs("uPWD") = GetMD5(temp)
        rs("uDesc") = temp
        rs.Update
        rs.MoveNext
    Loop
    rs.Close
    Set rs = Nothing
    MsgBox n & "条记录！"
    
End Sub
Private Function getPWD() As String
        strpwd = ""
        For i = 1 To 6
           strpwd = strpwd & Chr(Int(Rnd * 25 + 97))
        Next
        getPWD = strpwd

End Function

Private Sub Form_Activate()
On Error Resume Next
    FTextBox1.SetFocus
    
End Sub

Private Sub FTextBox1_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 13 Then txtPW.SetFocus
    
End Sub

Private Sub txtPW_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 13 Then
        cmdOK_Click
    End If
End Sub
