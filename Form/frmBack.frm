VERSION 5.00
Begin VB.Form frmBack 
   Caption         =   "Form1"
   ClientHeight    =   3090
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   ScaleHeight     =   3090
   ScaleWidth      =   4680
   StartUpPosition =   3  '窗口缺省
   Begin VB.PictureBox PicBack 
      Height          =   1110
      Left            =   240
      Picture         =   "frmBack.frx":0000
      ScaleHeight     =   1050
      ScaleWidth      =   2520
      TabIndex        =   0
      Top             =   330
      Width           =   2580
   End
End
Attribute VB_Name = "frmBack"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public Sub SetBack()
On Error Resume Next
    Me.PaintPicture PicBack.Picture, 0, 0, frmMain.Width, frmMain.Height
    Me.CurrentX = frmMain.ScaleWidth - 3500
    Me.CurrentY = frmMain.ScaleHeight - 1000
    frmMain.Picture = Me.Image
    frmMain.BackColor = frmMain.BackColor - 1 '为了刷MDI窗口,否则背景不会改变
End Sub
