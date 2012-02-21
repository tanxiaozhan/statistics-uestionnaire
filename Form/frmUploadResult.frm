VERSION 5.00
Object = "{48E59290-9880-11CF-9754-00AA00C00908}#1.0#0"; "msinet.ocx"
Begin VB.Form frmUploadResult 
   Caption         =   "查看已收传数据"
   ClientHeight    =   3075
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   7845
   Icon            =   "frmUploadResult.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3075
   ScaleWidth      =   7845
   StartUpPosition =   1  '所有者中心
   Begin VB.Timer TimerUpload 
      Enabled         =   0   'False
      Interval        =   500
      Left            =   5220
      Top             =   165
   End
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   1000
      Left            =   4380
      Top             =   135
   End
   Begin 调查问卷信息查询系统.XPButton cmdExit 
      Height          =   540
      Left            =   3120
      TabIndex        =   1
      Top             =   2430
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
   Begin InetCtlsObjects.Inet Inet1 
      Left            =   2490
      Top             =   30
      _ExtentX        =   1005
      _ExtentY        =   1005
      _Version        =   393216
   End
   Begin VB.Label Label2 
      Caption         =   "已上传学生编号："
      Height          =   225
      Left            =   540
      TabIndex        =   2
      Top             =   330
      Visible         =   0   'False
      Width           =   1455
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
      Height          =   1650
      Left            =   570
      TabIndex        =   0
      Top             =   720
      Width           =   6645
   End
End
Attribute VB_Name = "frmUploadResult"
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
        Inet1.Cancel
        Timer1.Enabled = False
        Label1.caption = "网络连接正常，请单击“上传数据”按钮上传数据！"
        Label1.ForeColor = RGB(0, 190, 0)
        
        Inet1.Execute "http://jspx.hpedu.gov.cn:8080/qs/getUploadResult.php?uid=" & curID, "GET"
        
        TimerUpload.Enabled = True
        lngCount = 0
        
        
        
        Do                  '检查网页是否下载成功
            DoEvents
            
            If lngCount > 60 Then
                MsgBox "网络连接超时，查询失败！", , "查看已上传数据"
                cmdUpload.Enabled = False
                Exit Sub
            End If
            
            If InStr(1, strHTTPText, "成功查看已上传数据") > 0 Then
                '获取已上传学生编号
                endPos = InStr(1, strHTTPText, ",WTUDEPH")
                If endPos < 1 Then
                    Label1.caption = "该班未上传数据。"
                    Label1.ForeColor = RGB(255, 0, 0)
                    Exit Do
                End If
                    
                Label2.Visible = True
                
                startPos = InStr(1, strHTTPText, "HPEDUTW")
                strno = Mid(strHTTPText, startPos + 7, endPos - startPos - 7)
                arrno = Split(strno, ",")
                
                Label1.caption = ""
                Label1.Alignment = 0   '左对齐
                For i = 0 To UBound(arrno)
                    If arrno(i) < 10 Then
                        Label1.caption = Label1.caption & " " & arrno(i) & "   "
                    Else
                        Label1.caption = Label1.caption & arrno(i) & "   "
                    End If
                    
                    If (i + 1) Mod 10 = 0 Then Label1.caption = Label1.caption & Chr(13)
                
                Next
                
                Exit Do
            End If
        Loop
        Inet1.Cancel
        TimerUpload.Enabled = False
        
        
        
        
        
    End If
End Sub

Private Sub TimerUpload_Timer()
    lngCount = lngCount + 1
End Sub
