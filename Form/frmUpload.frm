VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{48E59290-9880-11CF-9754-00AA00C00908}#1.0#0"; "msinet.ocx"
Begin VB.Form frmUpload 
   Caption         =   "�����ϴ�"
   ClientHeight    =   2640
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   7395
   Icon            =   "frmUpload.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2640
   ScaleWidth      =   7395
   StartUpPosition =   1  '����������
   Begin VB.Timer TimerUpload 
      Enabled         =   0   'False
      Interval        =   500
      Left            =   5220
      Top             =   165
   End
   Begin MSComctlLib.ProgressBar PBar1 
      Height          =   450
      Left            =   255
      TabIndex        =   3
      Top             =   1080
      Visible         =   0   'False
      Width           =   6930
      _ExtentX        =   12224
      _ExtentY        =   794
      _Version        =   393216
      Appearance      =   0
      Scrolling       =   1
   End
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   1000
      Left            =   4380
      Top             =   135
   End
   Begin �����ʾ���Ϣ��ѯϵͳ.XPButton cmdExit 
      Height          =   540
      Left            =   4260
      TabIndex        =   2
      Top             =   1830
      Width           =   1425
      _ExtentX        =   2514
      _ExtentY        =   953
      Caption         =   " ��  �� "
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
   Begin �����ʾ���Ϣ��ѯϵͳ.XPButton cmdUpload 
      Height          =   540
      Left            =   1620
      TabIndex        =   1
      Top             =   1830
      Width           =   1425
      _ExtentX        =   2514
      _ExtentY        =   953
      Caption         =   "�ϴ�����"
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
      Enabled         =   0   'False
      cBack           =   -2147483633
   End
   Begin InetCtlsObjects.Inet Inet1 
      Left            =   2490
      Top             =   30
      _ExtentX        =   1005
      _ExtentY        =   1005
      _Version        =   393216
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Caption         =   "Label1"
      BeginProperty Font 
         Name            =   "����"
         Size            =   10.5
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   300
      Left            =   270
      TabIndex        =   0
      Top             =   720
      Width           =   6915
   End
End
Attribute VB_Name = "frmUpload"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim strHTTPText As String
Dim intTimeCount As Integer
Dim uploadData(5000, 2) As String
Dim lngCount As Long      '��ʱ������


Private Sub cmdExit_Click()
    Inet1.Cancel
    
    Unload Me
End Sub

Private Sub cmdUpload_Click()
    '�ϴ�����
    
    Dim strCeti, strURL, strHttp As String    'strCeti-��֤�ַ�
    Dim sql As String
    Dim rs As ADODB.Recordset
        
    strCeti = "HPEDUTWQS" '��������������ί�ʾ����
    strURL = "http://jspx.hpedu.gov.cn:8080/qs/receive.php"
    Set rs = New ADODB.Recordset
    sql = "select * from main where mClass='" & curID & "'  order by mNo"
    
    cmdUpload.Enabled = False
    
    DBConnect
    rs.Open sql, Conn, 1, 1
    
    If rs.RecordCount < 1 Then MsgBox "δ¼�����ݣ��ϴ�����ʧ�ܣ�": Exit Sub
    
    PBar1.Max = rs.RecordCount
    
    PBar1.Visible = True
    Label1.ForeColor = RGB(0, 0, 0)
    Label1.caption = "�����ϴ����ݣ����Ժ�..."
    
    strURL = strURL & "?ceti=" & strCeti
    Do While Not rs.EOF
    
        strHttp = strURL & "&class=" & rs("mClass") & "&no=" & rs("mNo") & "&anwser=" & rs("mAnwser")
        strHTTPText = ""
        Inet1.Cancel
        
        Inet1.Execute strHttp, "GET"
        TimerUpload.Enabled = True
        lngCount = 0
        
        
        
        Do                  '��������ϴ��Ƿ�ɹ�
            DoEvents
            
            If lngCount > 60 Then
                MsgBox "�������ӳ�ʱ���ϴ�����ʧ�ܣ�", , "�ϴ�����"
                cmdUpload.Enabled = False
                Exit Sub
            End If
            
            If InStr(1, strHTTPText, "���ݸ�ʽ����") > 0 Or InStr(1, strHTTPText, "����������") > 0 Then
                MsgBox "�ϴ�����ʱ���ִ����ϴ�����ʧ�ܣ�", , "�ϴ�����"
                cmdUpload.Enabled = False
                Exit Sub
            End If
            
            
            If InStr(1, strHTTPText, "�ϴ��ɹ�") > 0 Then
                PBar1.value = PBar1.value + 1
                Exit Do
            End If
        Loop
        TimerUpload.Enabled = False
        
        rs.MoveNext
    
    Loop
    
    If PBar1.value = rs.RecordCount Then Label1.caption = "�����ϴ��ɹ���"
    
    rs.Close
    Set rs = Nothing
    Conn.Close
    

End Sub

Private Sub Form_Load()
    intTimeCount = 0
    strHTTPText = ""
    Label1.ForeColor = RGB(255, 0, 0)
    Label1.caption = "���ڽ����������Ӳ��ԣ����Ժ�..."
        
    Inet1.Execute "http://jspx.hpedu.gov.cn:8080/qs/conntest.php", "GET"
    
    Timer1.Enabled = True
    
End Sub

Private Sub Inet1_StateChanged(ByVal State As Integer)
    
    Dim strBuffer As String
    
    On Error Resume Next
    
    Select Case State
        
        Case icResponseCompleted
            
            Do  '�ӻ�������ȡ����
                DoEvents
                
                strBuffer = Inet1.GetChunk(512)
                strHTTPText = strHTTPText & strBuffer
                
            Loop Until Len(strBuffer) = 0
            
    End Select
    
End Sub

Private Sub Timer1_Timer()
    intTimeCount = intTimeCount + 1
    Label1.caption = "���ڽ����������Ӳ��ԣ����Ժ�...   " & intTimeCount & "��"
    If intTimeCount > 120 Then
        Timer1.Enabled = False
        Label1.ForeColor = RGB(255, 0, 0)
        Label1.caption = "��������ʧ�ܣ���������״̬��"
        cmdUpload.Enabled = False
    End If
    
    If InStr(1, strHTTPText, "�ʾ��ѯϵͳ�������Ӳ���ҳ") > 0 Then
        Timer1.Enabled = False
        Label1.caption = "���������������뵥�����ϴ����ݡ���ť�ϴ����ݣ�"
        Label1.ForeColor = RGB(0, 190, 0)
        cmdUpload.Enabled = True
        
    End If
End Sub

Private Sub TimerUpload_Timer()
    lngCount = lngCount + 1
End Sub
