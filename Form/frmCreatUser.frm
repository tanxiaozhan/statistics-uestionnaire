VERSION 5.00
Begin VB.Form frmCreateUser 
   Caption         =   "�����û�"
   ClientHeight    =   3090
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   5295
   Icon            =   "frmCreatUser.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3090
   ScaleWidth      =   5295
   StartUpPosition =   2  '��Ļ����
   Begin VB.PictureBox Picture1 
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   240
      Left            =   4125
      Picture         =   "frmCreatUser.frx":058A
      ScaleHeight     =   240
      ScaleWidth      =   240
      TabIndex        =   9
      ToolTipText     =   "����Ϊ1-16λ�����ġ�Ӣ�Ļ��������"
      Top             =   1110
      Width           =   240
   End
   Begin �����ʾ���Ϣ��ѯϵͳ.FTextBox txtPW 
      Height          =   300
      Left            =   1740
      TabIndex        =   2
      Top             =   1500
      Width           =   2295
      _ExtentX        =   4048
      _ExtentY        =   529
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "����"
         Size            =   9
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      FontName        =   "����"
      FontSize        =   9
      PasswordChar    =   "*"
   End
   Begin �����ʾ���Ϣ��ѯϵͳ.XPButton cmdCreate 
      Height          =   375
      Left            =   2580
      TabIndex        =   4
      Top             =   2580
      Width           =   1095
      _ExtentX        =   1931
      _ExtentY        =   661
      Caption         =   "����"
      CapAlign        =   2
      BackStyle       =   2
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "����"
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
   Begin �����ʾ���Ϣ��ѯϵͳ.XPButton cmdExit 
      Height          =   375
      Left            =   3900
      TabIndex        =   5
      Top             =   2580
      Width           =   1095
      _ExtentX        =   1931
      _ExtentY        =   661
      Caption         =   "�˳�"
      CapAlign        =   2
      BackStyle       =   2
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "����"
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
   Begin �����ʾ���Ϣ��ѯϵͳ.FTextBox txtUser 
      Height          =   300
      Left            =   1740
      TabIndex        =   1
      ToolTipText     =   "�ɳ���1-16λ�����ġ�Ӣ�Ļ��������"
      Top             =   1080
      Width           =   2295
      _ExtentX        =   4048
      _ExtentY        =   529
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "����"
         Size            =   9
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      FontName        =   "����"
      FontSize        =   9
   End
   Begin �����ʾ���Ϣ��ѯϵͳ.FTextBox txtConform 
      Height          =   300
      Left            =   1740
      TabIndex        =   3
      Top             =   1920
      Width           =   2295
      _ExtentX        =   4048
      _ExtentY        =   529
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "����"
         Size            =   9
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      FontName        =   "����"
      FontSize        =   9
      PasswordChar    =   "*"
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "ȷ�����룺"
      Height          =   180
      Left            =   780
      TabIndex        =   8
      Top             =   1995
      Width           =   900
   End
   Begin VB.Shape Shape1 
      BackColor       =   &H00FFF8F0&
      BorderColor     =   &H00C5742F&
      Height          =   1575
      Left            =   240
      Top             =   840
      Width           =   4860
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "�� �� ����"
      Height          =   180
      Left            =   780
      TabIndex        =   7
      Top             =   1125
      Width           =   900
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "��    �룺"
      Height          =   180
      Left            =   780
      TabIndex        =   6
      Top             =   1560
      Width           =   900
   End
   Begin VB.Label Label1 
      Caption         =   "��һ��ʹ�ñ�ϵͳ���贴�������û��ʺš�"
      BeginProperty Font 
         Name            =   "����"
         Size            =   12
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   480
      TabIndex        =   0
      Top             =   240
      Width           =   5055
   End
End
Attribute VB_Name = "frmCreateUser"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdCreate_Click()
    On Error GoTo ErrProcess
    Dim strsql As String
    
    If Me.txtPW.Text <> Me.txtConform.Text Then
        MsgBox "������������벻һ�£����������롣", vbExclamation
        Me.txtConform.SetFocus
        Exit Sub
        
    Else
        strsql = "insert into userInfo(uID,uPWD,uRight) " & " " & _
                               "values('" & Me.txtUser.Text & "','" & GetMD5(Me.txtPW.Text) & "','1')"

        DBConnect
        Conn.Execute strsql
        If Conn.State > 0 Then
            Conn.Close
        End If
        
    End If
    frmLogin.Show
    Unload Me
    Exit Sub
ErrProcess:
    MsgBox Err.Description, , "������ʾ"

End Sub

Private Sub cmdExit_Click()
    End
End Sub

Private Sub Form_Activate()
    Me.txtUser.SetFocus
End Sub

Private Sub Picture1_Click()
    MsgBox "����Ϊ1-16λ�����ġ�Ӣ�Ļ�������ϡ�", vbInformation, "�����û�"
End Sub

Private Sub txtConform_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 13 Then
        Me.cmdCreate.SetFocus
    End If
End Sub
Private Sub txtPW_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 13 Then
        Me.txtConform.SetFocus
    End If
End Sub

Private Sub txtUser_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 13 Then
        Me.txtPW.SetFocus
    End If
End Sub
