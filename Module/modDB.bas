Attribute VB_Name = "modDB"
Public Conn  As New ADODB.Connection

'����ACCESS���ݿ�
Sub DBConnect()
    
    strconn = "Provider=Microsoft.Jet.OLEDB.4.0;jet oledb:database Password=txzlsh;Data Source=" & GetApp & "Data\data.qta"
    If Conn.State <> 0 Then Conn.Close
    Conn.Open strconn
    
End Sub

'��ȡSQL������������Ϣ
Public Sub readServer()
On Error GoTo aaaa
    Dim strTmp As String, strT() As String
    Open GetApp & "Files\sql.inf" For Input As #1
        If EOF(1) = False Then Line Input #1, strTmp
    Close #1
    strTmp = Trim(strTmp)
    If strTmp <> "" Then
        strT = Split(strTmp, "||")
        For i = 0 To 3
            strT(i) = strT(i)
        Next
        strSQLServer = strT(0)
        strSQLUser = strT(1)
        strSQLPW = strT(2)
        strSQLDB = strT(3)
    End If
Exit Sub
aaaa:
    strSQLServer = ""
    strSQLUser = ""
    strSQLPW = ""
    strSQLDB = ""
End Sub

'����SQL������������Ϣ
Public Sub SaveServer(ByVal strServer As String, ByVal strUser As String, ByVal strPass As String, ByVal strDataBase)
On Error GoTo aaaa
    Open GetApp & "Files\sql.inf" For Output As #1
        Print #1, strServer & "||" & strUser & "||" & strPass & "||" & strDataBase
    Close #1
Exit Sub
aaaa:
    MsgBox "���� SQL ��������Ϣʧ�ܣ�", vbCritical
End Sub
