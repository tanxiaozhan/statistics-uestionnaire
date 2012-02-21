VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{48E59290-9880-11CF-9754-00AA00C00908}#1.0#0"; "msinet.ocx"
Begin VB.Form frmUpload 
   Caption         =   "数据上传"
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
   StartUpPosition =   1  '所有者中心
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
   Begin 调查问卷信息查询系统.XPButton cmdExit 
      Height          =   540
      Left            =   4260
      TabIndex        =   2
      Top             =   1830
      Width           =   1425
      _ExtentX        =   2514
      _ExtentY        =   953
      Caption         =   " 返  回 "
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
   Begin 调查问卷信息查询系统.XPButton cmdUpload 
      Height          =   540
      Left            =   1620
      TabIndex        =   1
      Top             =   1830
      Width           =   1425
      _ExtentX        =   2514
      _ExtentY        =   953
      Caption         =   "上传数据"
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
         Name            =   "宋体"
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
Dim lngCount As Long      '定时器计数


Private Sub cmdExit_Click()
    Inet1.Cancel
    
    Unload Me
End Sub

Private Sub cmdUpload_Click()
    '上传数据
    
    Dim strCeti, strURL, strHttp As String    'strCeti-认证字符
    Dim sql As String
    Dim rs As ADODB.Recordset
        
    strCeti = "HPEDUTWQS" '黄埔区教育局团委问卷调查
    strURL = "http://jspx.hpedu.gov.cn:8080/qs/receive.php"
    Set rs = New ADODB.Recordset
    sql = "select * from main where mClass='" & curID & "'  order by mNo"
    
    cmdUpload.Enabled = False
    
    DBConnect
    rs.Open sql, Conn, 1, 1
    
    If rs.RecordCount < 1 Then MsgBox "未录入数据，上传数据失败！": Exit Sub
    
    PBar1.Max = rs.RecordCount
    
    PBar1.Visible = True
    Label1.ForeColor = RGB(0, 0, 0)
    Label1.caption = "正在上传数据，请稍候..."
    
    strURL = strURL & "?ceti=" & strCeti
    Do While Not rs.EOF
    
        strHttp = strURL & "&class=" & rs("mClass") & "&no=" & rs("mNo") & "&anwser=" & rs("mAnwser")
        strHTTPText = ""
        Inet1.Cancel
        
        Inet1.Execute strHttp, "GET"
        TimerUpload.Enabled = True
        lngCount = 0
        
        
        
        Do                  '检查数据上传是否成功
            DoEvents
            
            If lngCount > 60 Then
                MsgBox "网络连接超时，上传数据失败！", , "上传数据"
                cmdUpload.Enabled = False
                Exit Sub
            End If
            
            If InStr(1, strHTTPText, "数据格式错误") > 0 Or InStr(1, strHTTPText, "参数不完整") > 0 Then
                MsgBox "上传数据时出现错误，上传数据失败！", , "上传数据"
                cmdUpload.Enabled = False
                Exit Sub
            End If
            
            
            If InStr(1, strHTTPText, "上传成功") > 0 Then
                PBar1.value = PBar1.value + 1
                Exit Do
            End If
        Loop
        TimerUpload.Enabled = False
        
        rs.MoveNext
    
    Loop
    
    If PBar1.value = rs.RecordCount Then Label1.caption = "数据上传成功！"
    
    rs.Close
    Set rs = Nothing
    Conn.Close
    

End Sub

Private Sub Form_Load()
    intTimeCount = 0
    strHTTPText = ""
    Label1.ForeColor = RGB(255, 0, 0)
    Label1.caption = "正在进行网络连接测试，请稍候..."
        
    Inet1.Execute "http://jspx.hpedu.gov.cn:8080/qs/conntest.php", "GET"
    
    Timer1.Enabled = True
    
End Sub

Private Sub Inet1_StateChanged(ByVal State As Integer)
    
    Dim strBuffer As String
    
    On Error Resume Next
    
    Select Case State
        
        Case icResponseCompleted
            
            Do  '从缓冲区读取数据
                DoEvents
                
                strBuffer = Inet1.GetChunk(512)
                strHTTPText = strHTTPText & strBuffer
                
            Loop Until Len(strBuffer) = 0
            
    End Select
    
End Sub

Private Sub Timer1_Timer()
    intTimeCount = intTimeCount + 1
    Label1.caption = "正在进行网络连接测试，请稍候...   " & intTimeCount & "秒"
    If intTimeCount > 120 Then
        Timer1.Enabled = False
        Label1.ForeColor = RGB(255, 0, 0)
        Label1.caption = "网络连接失败，请检查网络状态！"
        cmdUpload.Enabled = False
    End If
    
    If InStr(1, strHTTPText, "问卷查询系统网络连接测试页") > 0 Then
        Timer1.Enabled = False
        Label1.caption = "网络连接正常，请单击“上传数据”按钮上传数据！"
        Label1.ForeColor = RGB(0, 190, 0)
        cmdUpload.Enabled = True
        
    End If
End Sub

Private Sub TimerUpload_Timer()
    lngCount = lngCount + 1
End Sub
