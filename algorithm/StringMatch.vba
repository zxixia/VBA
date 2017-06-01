Private Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)
Sub StringMatch()

    Dim strString As String
    strString = "abcdefgabcdefg"
    MsgBox "�����ַ����ǣ�" & Chr(10) & strString, vbInformation, "��ʾ"
    
    Dim strCheck As String
    
    '=====================================
    '����Ҫ���ҵ��ַ���
     strCheck = InputBox("������", "����������ҵ��ַ���", 1)
    
    If strCheck <> "" Then
     MsgBox "��" & Chr(10) & strString & Chr(10) & "�в��ң�" & Chr(10) & strCheck, vbInformation, "��ʾ"
    Else
     MsgBox ("����Ϊ�գ����������к��ļ�")
     Exit Sub
    End If
    '=====================================
    Debug.Print strString & Chr(10) & strCheck
    
    
    '=====================================
    '�����㷨
    Dim shotArr() As Integer
    shotArr = naiveStringMatch(strString, strCheck)
    printResult strString, strCheck, shotArr, "�����㷨"
    '=====================================
    
    
    '=====================================
    'RabinKarp�㷨
    shotArr = rabinKarpMatch(strString, strCheck)
    printResult strString, strCheck, shotArr, "RabinKarp�㷨"
    '=====================================
    
    printPreAndPostfix (strString)
End Sub
'�����㷨
Function naiveStringMatch(strString As String, _
                          strCheck As String)
'strString �Ǳ����ҵĳ�ʼ�ַ���
'strCheck  ������Ĵ������ַ���
    
    lenString = Len(strString)
    lenCheck = Len(strCheck)
    
    
    'һ��������ż�������index������
    Dim shotArr(1 To 10) As Integer
    Dim shotCount
    shotCount = 1
    
    'MsgBox "�����ҳ�ʼ�ַ���������" & lenString & " �����ҵ��ַ���������" & lenCheck, vbInformation, "��ʾ"
    For i = 1 To (lenString - lenCheck + 1)
        '��������м���,Mid����Ҫ��1��ʼ
        'Debug.Print Mid(strString, i + 1, 1)
        For j = 1 To lenCheck
            strA = Mid(strString, i, lenCheck)
            'Debug.Print strA & " "; Mid(strCheck, j, 1)
            If Mid(strA, j, 1) <> Mid(strCheck, j, 1) Then
                '��������ַ�����һ�㲻һ��,����ѭ��
                Exit For
            End If
            If j = lenCheck Then
                '�ҵ��Ӵ�
                'MsgBox i, vbInformation, "��ʾ"
                shotArr(shotCount) = i
                shotCount = shotCount + 1
            End If
        Next j
    Next i
    naiveStringMatch = shotArr
End Function
'RabinKarp�㷨
Function rabinKarpMatch(strString As String, _
                        strCheck As String)
    lenCheck = Len(strCheck)
    lenString = Len(strString)
    Debug.Print strString & " " & Hash(strString)
    Debug.Print strCheck & " " & Hash(strCheck)
    
    'һ��������ż�������index������
    Dim shotArr(1 To 10) As Integer
    Dim shotCount
    shotCount = 1
    
    hashCheck = Hash(strCheck)
    For i = 1 To (lenString - lenCheck + 1)
        toCheck = Mid(strString, i, lenCheck)
        If Hash(toCheck) = hashCheck Then
            For j = 1 To lenCheck
                If Mid(toCheck, j, 1) <> Mid(strCheck, j, 1) Then
                    '��������ַ�����һ�㲻һ��,����ѭ��
                    Exit For
                End If
                If j = lenCheck Then
                    '�ҵ��Ӵ�
                    'MsgBox i, vbInformation, "��ʾ"
                    shotArr(shotCount) = i
                    shotCount = shotCount + 1
                End If
            Next j
        End If
    Next i
    rabinKarpMatch = shotArr
End Function
'ģ��hash����,��ʵ���Ǽ���һ���ַ�������������ַ���ASCII��ֵ�ĺ�
Function Hash(strInput)
    Dim count As Integer
    For i = 1 To Len(strInput)
        count = count + Asc(Mid(strInput, i, 1))
    Next i
    Hash = count
End Function

'�������ķ���
Function printResult(strString As String, _
                     strCheck As String, _
                     shotArr() As Integer, _
                     strTitle As String)
    lenCheck = Len(strCheck)
    lenString = Len(strString)
    
    Dim strResult As String
    Dim start As Integer
    start = 1
    
    For Each i In shotArr
        If i < 1 Then
            Exit For
        End If
        Debug.Print i
        If start <= i Then
            strResult = strResult & Mid(strString, start, i - Len(strResult) - 1) & UCase(strCheck)
        End If
        start = i + lenCheck
    Next i
    '���������ַ���
    strResult = strResult & Mid(strString, start, lenString - Len(strResult))
    MsgBox strString & Chr(10) & strResult, vbInformation, strTitle
    Debug.Print strString
    Debug.Print strResult
End Function
'���һ���ַ�����ǰ׺�ͺ�׺
'���ڵ���
Function printPreAndPostfix(strString As String)
    Dim prefix As String
    Dim postfix As String
    
    For i = 1 To Len(strString) - 1
        prefix = prefix & " " & Mid(strString, 1, i)
        '�����������������
        'postfix = postfix & " " & Mid(strString, Len(strString) - i + 1, i)
        postfix = postfix & " " & Mid(strString, i + 1, Len(strString) - i)
    Next i
    Debug.Print prefix
    Debug.Print postfix
    MsgBox "ǰ׺:" & prefix & Chr(10) & "��׺:" & postfix, vbInformation, "ǰ��׺"
End Function
                     

