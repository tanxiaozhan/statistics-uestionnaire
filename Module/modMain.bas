Attribute VB_Name = "modMain"
Public GetApp As String '����·��
Public curID As String    '��ǰ�û���
Public curUserLevel As Long    '��ǰ�û�����
Public DataOperateState As String      '����¼��/�༭״̬
Public mainID As Long   'ѡ���main���ݱ��¼��ID��,���ڱ༭�޸Ķ�Ӧ�ļ�¼
Public subID As Long
Public borrowID As Long   'ѡ��Ľ�֧��ļ�¼ID
Public incomeID As Long
Public curDOCType As Integer    '�����ĵ����ͣ�1-���㵥��2-��Ŀȷ�ϵ���3-��Ŀ��֧��
Public dblBalace As Double      '��֧���
Public strContractType() As String    '��ͬ����
Public strMode() As String
Public curList1Index As Integer     '��ͬ�б�λ��
Public curList2Index As Byte
Public curList3Index As Byte
Public curList4Index As Byte
Public curList5Index As Byte
Public bytAfterDec As Byte       '�������ı���С��λ��
Public color(2) As Long         '0-�б���ɫ��1-�б��ı�ɫ��2-�ѽ����ı�ɫ
Public curSchool As String     'ѧУ����
Public curClass As String      '����
Public curNo As String         '�ʾ���
Public question1(8) As Byte    '�ʾ�ÿ��ѡ����Ŀ
Public question2(15) As Byte
Public question3(15) As Byte
Public question4(15) As Byte
Public question5(15) As Byte
Public IsEdit As Boolean       '¼���ʾ����״̬
Public curInputForm As Byte      '¼�봰����������ѡ��1-Сѧ1-3�꼶��2-Сѧ4-6�꼶��3-��ѧ��¼�봰��



'�������
Public Sub Main()
'On Error Resume Next
    
    If App.PrevInstance Then
        End
        Exit Sub
    End If
    '��ñ���·��
    GetApp = App.Path: If Right$(GetApp, 1) <> "\" Then GetApp = GetApp & "\"
    
    '��ȡ��ʼ������
    'GetItemInfo
    
    '�ʾ�ÿ��ѡ����Ŀ
    '��һ����
    question1(1) = 10
    question1(2) = 5
    question1(3) = 6
    question1(4) = 4
    question1(5) = 4
    question1(6) = 4
    question1(7) = 3
    question1(8) = 4
    
    For i = 1 To 15
        question2(i) = 4
        question3(i) = 4
        question4(i) = 4
        question5(i) = 4
        
    Next
    
    '��������
    question3(6) = 12
    question3(10) = 3
    question3(11) = 3
    
    question4(2) = 3
    question4(3) = 3
    question4(5) = 3
    question4(14) = 5
    
    question5(1) = 8
    question5(2) = 5
    question5(3) = 3
    question5(4) = 3
    question5(5) = 3
    question5(12) = 3
    question5(13) = 3
    question5(14) = 3
    question5(15) = 3
    
    'frmMain.Show
    frmLogin.Show
End Sub

Public Sub SetSB(ByVal i&, ByVal strText$)
    frmMain.LBSB(i).caption = strText
End Sub

Public Function GetINI(ByVal s1 As String, s2 As String)
On Error Resume Next
    GetINI = GetSetting("MySuperMarket", s1, s2)
End Function

Public Function GetLongINI(ByVal s1 As String, s2 As String, Optional Def As Long = 0) As Long  '��ȡINI������ֵ
On Error GoTo aaaa
    Dim str As String
    str = GetINI(s1, s2)
    If str = "" Then
        GetLongINI = Def
    Else
        GetLongINI = CLng(str)
    End If
    Exit Function
aaaa:
    GetLongINI = Def
End Function

Public Sub SaveINI(ByVal s1 As String, s2 As String, s3 As String)
On Error Resume Next
    SaveSetting "MySuperMarket", s1, s2, s3
End Sub

Function FieldTypeIsChar(n As Long) As Boolean    '�ж��ֶ��Ƿ��������ͣ����ڲ����¼ʱ�Ƿ������
    Dim IsChar As Boolean
    
Select Case n
'case���� ֵ ˵��
'Case 0x2000
' p = AdArray '���������� ADOX���� 0x2000 һ����־ֵ��ͨ������һ���������ͳ�����ϣ�ָʾ���������͵����顣
Case 20, 128, 14, 5, 3, 205, 131, 4, 2, 16, 21, 19, 18, 17, 204
    IsChar = False
Case 8, 136, 129, 6, 7, 133, 134, 135, 205, 203, 200, 202
    IsChar = True
End Select

FieldTypeIsChar = IsChar

End Function

Function GetID(id As String) As String
    GetID = Left(id, Len(id) - 1)
End Function

Function coverToChinese(money As String) As String
    Dim x As String, y As String
    Const zimu = ".sbqwsbqysbqwsbq" '����λ�ô���
    Const letter = "0123456789sbqwy.zjf" '���庺����д
    Const upcase = "��Ҽ��������½��ƾ�ʰ��Ǫ����Ԫ���Ƿ�" '�����д����
    Dim temp As String
    temp = money
    If InStr(temp, ".") > 0 Then temp = Left(temp, InStr(temp, ".") - 1)

    If Len(temp) > 16 Then MsgBox "��Ŀ̫���޷����㣡������һ�������µ�����", vbCritical, "������ʾ": Exit Function  'ֻ��ת��һ����Ԫ������Ŀ�Ļ��ң�

    x = Format(money, "0.00") '��ʽ������
    y = ""
    For i = 1 To Len(x) - 3
        y = y & Mid(x, i, 1) & Mid(zimu, Len(x) - 2 - i, 1)
    Next
    If Right(x, 3) = ".00" Then
        y = y & "z"          '***Ԫ��
    Else
        y = y & Left(Right(x, 2), 1) & "j" & Right(x, 1) & "f"     '*Ԫ*��*��
    End If
    
    y = Replace(y, "0q", "0") '������ǧ(�磺40200���f��ǧ�㷡��)
    y = Replace(y, "0b", "0") '�������(�磺41000���fҼǧ���)
    y = Replace(y, "0s", "0") '������ʮ(�磺204������ʰ����)

    Do While y <> Replace(y, "00", "0")
        y = Replace(y, "00", "0") '����˫��(�磺1004ҼǪ������)
    Loop
    
    y = Replace(y, "0y", "y") '������|(�磺210�|     ����Ҽʮ��|)
    y = Replace(y, "0w", "w") '�������f(�磺210�f     ����Ҽʮ���f)
    y = IIf(Len(x) = 5 And Left(y, 1) = "1", Right(y, Len(y) - 1), y) '����Ҽʮ(�磺14Ҽʰ����10Ҽʰ)
    y = IIf(Len(x) = 4, Replace(y, "0.", ""), Replace(y, "0.", ".")) '������Ԫ(�磺20.00��ʰ��Բ��0.12��ԲҼ�Ƿ���)

    For i = 1 To 19
        y = Replace(y, Mid(letter, i, 1), Mid(upcase, i, 1)) '��д����
    Next
    coverToChinese = y
    
End Function
  Public Function DirExists(ByVal strDirName As String) As Integer
          Const strWILDCARD$ = "*.*"
        
          Dim strDummy     As String
    
          On Error Resume Next
          If Trim(strDirName) = "" Then
                DirExists = 0
                Exit Function
          End If
          strDummy = Dir$(strDirName & strWILDCARD, vbDirectory)
          DirExists = Not (strDummy = vbNullString)
    
          Err = 0
  End Function
  Public Sub GetItemInfo()
    Dim rs As ADODB.Recordset
    Dim strsql As String
    
    Set rs = New ADODB.Recordset
    DBConnect
      
      '��ͬ����
    strsql = "select * from ItemInfo where ItemType=1 order by ItemID"
    rs.Open strsql, Conn, 1, 1
    ReDim strContractType(IIf(rs.RecordCount > 0, rs.RecordCount - 1, 0), 1)
    For i = 1 To rs.RecordCount
        strContractType(i - 1, 0) = rs("ItemName")
        strContractType(i - 1, 1) = rs("ItemID")
        rs.MoveNext
    Next
    
    rs.Close
    strsql = "select * from ItemInfo where ItemType=2 order by ItemID"
    rs.Open strsql, Conn, 1, 1
    ReDim strMode(IIf(rs.RecordCount > 0, rs.RecordCount - 1, 0), 1)
    For i = 1 To rs.RecordCount
        strMode(i - 1, 0) = rs("ItemName")
        strMode(i - 1, 1) = rs("ItemID")
        rs.MoveNext
    Next
    
    'С��λ��
    Set rs = New ADODB.Recordset
    strsql = "select ItemValue from ItemInfo where ItemType=3"
    rs.Open strsql, Conn, 1, 1
    
    bytAfterDec = 3          '��λС��
    
    If Not rs.EOF Then
        If Not IsNull(rs("ItemValue")) Then bytAfterDec = rs("ItemValue")
    Else
        Conn.Execute "insert into ItemInfo(ItemType,ItemValue) values(3,3)"
    End If
    rs.Close
        
    '��ɫ
    strsql = "select * from ItemInfo where ItemType=4 order by ItemID"
    rs.Open strsql, Conn, 1, 1
    If rs.RecordCount <> 3 Then
        color(0) = "&Hfafafa"     'ȱʡ��ɫ
        color(1) = "&H000000"
        color(2) = "&H008000"
    Else
        color(0) = rs("ItemValue")
        rs.MoveNext
        color(1) = rs("ItemValue")
        rs.MoveNext
        color(2) = rs("ItemValue")
    End If
    
    rs.Close
    Conn.Close
    Set rs = Nothing
    Set Conn = Nothing

  End Sub
