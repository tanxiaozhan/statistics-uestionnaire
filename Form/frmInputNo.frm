VERSION 5.00
Begin VB.Form frmInputNo 
   Caption         =   "¼����"
   ClientHeight    =   2055
   ClientLeft      =   8265
   ClientTop       =   6390
   ClientWidth     =   3645
   ControlBox      =   0   'False
   LinkTopic       =   "Form2"
   ScaleHeight     =   2055
   ScaleWidth      =   3645
   StartUpPosition =   1  '����������
   Begin �����ʾ���Ϣ��ѯϵͳ.XPButton XPButton2 
      Height          =   375
      Left            =   2145
      TabIndex        =   3
      Top             =   1290
      Width           =   1155
      _ExtentX        =   2037
      _ExtentY        =   661
      Caption         =   " ȡ  �� "
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
   Begin �����ʾ���Ϣ��ѯϵͳ.XPButton XPButton1 
      Height          =   375
      Left            =   525
      TabIndex        =   2
      Top             =   1275
      Width           =   1155
      _ExtentX        =   2037
      _ExtentY        =   661
      Caption         =   "ȷ  ��"
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
   Begin �����ʾ���Ϣ��ѯϵͳ.FTextBox txtNo 
      Height          =   300
      Left            =   2310
      TabIndex        =   1
      Top             =   465
      Width           =   900
      _ExtentX        =   1588
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
      isNumber        =   -1  'True
      MaxLength       =   3
      afterdecimal    =   0
   End
   Begin VB.Label Label1 
      Caption         =   "�������ʾ��ţ�"
      BeginProperty Font 
         Name            =   "����"
         Size            =   10.5
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   495
      TabIndex        =   0
      Top             =   510
      Width           =   1695
   End
End
Attribute VB_Name = "frmInputNo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub txtNo_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 13 Then
        XPButton1.SetFocus
    End If
End Sub

Private Sub XPButton1_Click()
    Dim rs As ADODB.Recordset
    Set rs = New ADODB.Recordset
    
    curNo = Trim(txtNo.Text)
    If curNo = "" Then
        MsgBox "����������ʾ��ţ�", vbExclamation, "���"
        txtNo.SetFocus
        Exit Sub
    End If
    DBConnect
    sql = "select mNo from main where mNo=" & curNo & " and mClass='" & curID & "'"
    rs.Open sql, Conn, 1, 1
    If rs.RecordCount > 0 Then
        MsgBox "�˱���Ѵ��ڣ������������ţ�", vbExclamation, "���"
        txtNo.SetFocus
        Exit Sub
    End If
    
    
    Unload Me
    
    Select Case curInputForm
        Case 1:
            frmInputMainSCL.Show     'Сѧ1-3�꼶
        Case 2:
            frmInputMainSCH.Show     'Сѧ4-6�꼶
        Case 3:
            frmInputMainMS.Show      '���С�����
    End Select
    
End Sub

Private Sub XPButton2_Click()
    frmMain.cmdLeft_Click 1
    frmMain.picStudentList.Visible = False
    frmMain.mnuStudentList.Checked = False
    If GetINI("Main", "Guide") = "n" Then
        frmMain.picLeft.Visible = False
        frmMain.mnuGuide.Checked = False
    End If
    
    Unload Me
End Sub
