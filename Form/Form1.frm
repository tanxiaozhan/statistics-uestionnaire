VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "修改密码"
   ClientHeight    =   2970
   ClientLeft      =   1485
   ClientTop       =   1350
   ClientWidth     =   5730
   ControlBox      =   0   'False
   Icon            =   "Form1.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   2970
   ScaleWidth      =   5730
   StartUpPosition =   1  '所有者中心
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
    DBConnect
    sql = "create table main(mid int,mClass char(6),mNo int"
    For i = 1 To 8
        sql = sql & ",m1" & i & " char(1)"
    Next
    
    For i = 2 To 5
        For j = 1 To 15
            sql = sql & ",m" & i & j & " char(1)"
        Next
    Next
    
    sql = sql & ")"
    
    Conn.Execute "drop table main"
    Conn.Execute sql
    'Conn.Execute "create table main(mid int,mClass char(6))"
    'MsgBox sql
    
    Conn.Close
    

End Sub
