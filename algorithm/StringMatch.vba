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
    
    printRecursiveGetNext ("abcdqwfabcZZabcdqwfabcdabcdqwfabcZZabcdqwfabcd")
    printGetNext ("abcdqwfabcZZabcdqwfabcdabcdqwfabcZZabcdqwfabcd")
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
        For J = 1 To lenCheck
            strA = Mid(strString, i, lenCheck)
            'Debug.Print strA & " "; Mid(strCheck, j, 1)
            If Mid(strA, J, 1) <> Mid(strCheck, J, 1) Then
                '��������ַ�����һ�㲻һ��,����ѭ��
                Exit For
            End If
            If J = lenCheck Then
                '�ҵ��Ӵ�
                'MsgBox i, vbInformation, "��ʾ"
                shotArr(shotCount) = i
                shotCount = shotCount + 1
            End If
        Next J
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
            For J = 1 To lenCheck
                If Mid(toCheck, J, 1) <> Mid(strCheck, J, 1) Then
                    '��������ַ�����һ�㲻һ��,����ѭ��
                    Exit For
                End If
                If J = lenCheck Then
                    '�ҵ��Ӵ�
                    'MsgBox i, vbInformation, "��ʾ"
                    shotArr(shotCount) = i
                    shotCount = shotCount + 1
                End If
            Next J
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

'
' =================================================================
'
' ����
' a) N[j]          ��ʾ����Ϊj���ַ���
' b) n[1...j]      ��ʾ����Ϊj���ַ�����Ӧ��1...jλ���ľ���ȡֵ
' c) Next(N[j])    ��ʾ����Ϊj���ַ�����Ӧ��ģʽֵ
' =================================================================

' ģʽֵ�������ȡ
' �ο���ҳ��http://www.ituring.com.cn/article/59881
' N[4] = abca , Next(N[4]) =  1,  ��ʾn[1] = n[4]
' N[5] = abcab, Next(N[5]) =  2,  ��ʾn[1] = n[4], n[2] = n[5]
' N[3] = abc  , Next(N[3]) =  0,  ��ʾû����βƥ��


' ������֪�ַ���N[j]��Ӧ��ģʽֵ, Next(N[j]) = i,������ӻ���ʾ����ͼ��ʾ
'
' N[j] = abcdqwfabc
'        123456789*
'          |      0
'          i=3
'
' ����ͼ��֪
' j=10,  ��ʾ���ַ����ĳ���
' i=3,   ��ʾ���ַ����ģʽƥ���ǰ׺����ĩλ������,Ҳ����ʾ
'        n[1...i] == n[(j+1-i)...j]
'        n[1] = n[8]  = "a"
'        n[2] = n[9]  = "b"
'        n[3] = n[10] = "c"
'


' ������֪�ַ���N[j]��ģʽֵ,Next(N[j]) = i
' ���潫�֡�2��������������������ַ���N[j+1]��ģʽֵ,Next(N[j+1])��ֵ

' ��1��,
'  n[j+1] == n[i+1]�����
'
' ���ַ��� N[11]="abcdqwfabcd"
'                 123456789**
'                          01
'
'
' ���Ӵ�   N[10]="abcdqwfabc", ��Next(N[10]) = 3 , ��j=10, ��i=3
'                 123456789*
'                   |      0
'                   i=3
'
' ͬʱ     (n[j+1] = n[10+1] = n[11]) == (n[i+1] = n[3+1] = n[4]) = "d"
' ���Ƴ�   Next(N[11]) = Next(N[10]) + 1
' Ҳ����
' ���     n[j+1] == n[i+1]
' ��       Next(N[j+1]) = Next(N[j]) + 1
'                       = Next(N[10]) + 1
'                       = i+1
'                       = 4
'
'
' N[11] = "abcdqwfabcd"
'          123456789**
'             |     01
'             i=4
'
' =================================================================
'
'
'
' ��2��,
'  n[j+1] != n[i+1]�����
'
'  ���ַ���
'  N[j+1=23] ="abcdqwfabcZZabcdqwfabcd"
'              123456789**************
'                       0123456789****
'                                 0123
'
'
'
'  N[j=22]   ="abcdqwfabcZZabcdqwfabc"
'              123456789*************
'                       0123456789***
'                       |         012
'                       i=10
'  j=22
'  i=10
'
'  A, (n[i+1] = n[10+1] = n[11] = "Z") != (n[j+1] = n[22+1] = n[23] = "d")
'  B, ����Next(N[i])
'
'
'  N[i=10] ="abcdqwfabc"
'            123456789*
'              |      0
'              k=3
'
'  i=10
'  k=3
'
'
'  N[22] ="abcdqwfabcZZabcdqwfabc"
'          123456789*************
'          |        0123456789***
'          |        |  |      012
'          ----------  |        |
'               A      ----------
'                          B
'
'
'  ��ΪNext(N[22]) = 10
'  ����A���� == B����
'  �ٵ�������Next(N[10]) = 3 = k��ֵ
'  �õ����¸���һ���Ļ���
'
'
'           a1     a2
'          ---    ---   b1     b2
'          | |    | |  ---    ---
'          | |    | |  | |    | |
'  N[22] ="abcdqwfabcZZabcdqwfabc"
'          123456789*************
'          |        0123456789***
'          |        |  |      012
'          ----------  |        |
'               A      ----------
'                           B
'
'  Ҳ���ǵõ�
'  a1 == a2 == b1 == b2
'
'  �ó�����Ҫ��a1==b2�ؼ���Ϣ
'  ��ʱ���a1�ĺ�һλ��b2�ĺ�һλ���
'  ��õ���Next(N[j+1])��ֵ
'
'
'
'
'           a1     a2
'          ---    ---   b1     b2
'          | |    | |  ---    ---
'          | |    | |  | |    | |
'  N[22] ="abcdqwfabcZZabcdqwfabcd"
'          123456789**************
'             |     0123456789****
'             |               0123
'             |                  |
'             a1��һλ           |
'             n[k+1]             |
'                                b2��һλ
'                                n[j+1]
'
'  ��� n[k+1] == n[j+1]
'  ��   Next(N[j+1]) = Next(N[k]) + 1
'                    = 3 + 1
'                    = 4
'
'  ��� n[k+1] != n[j+1]
'  �����ǰ��ġ�B, ����Next(N[i])��
'  ***********�����������****************
'  ֱ���ҵ�һ��k1,
'  ���� n[k1+1] = n[j+1], ��ʱNext(N[j+1]) = Next(N[k1]) + 1, �ο�ǰ���ͼʾ���
'
' =============================================================================================
'
'
' ������ʺ�ʹ�õݹ�ķ�ʽ������Next����
'
'
'
'
' ===================================================
' ��bug
'
' ���¼����ֵ����
' abcdqwfabcZZabcdqwfabcda, 0
'
'
' a, 0
' ab, 0
' abc, 0
' abcd, 0
' abcdq, 0
' abcdqw, 0
' abcdqwf, 0
' abcdqwfa, 1
' abcdqwfab, 2
' abcdqwfabc, 3
' abcdqwfabcZ, 0
' abcdqwfabcZZ, 0
' abcdqwfabcZZa, 1
' abcdqwfabcZZab, 2
' abcdqwfabcZZabc, 3
' abcdqwfabcZZabcd, 4
' abcdqwfabcZZabcdq, 5
' abcdqwfabcZZabcdqw, 6
' abcdqwfabcZZabcdqwf, 7
' abcdqwfabcZZabcdqwfa, 8
' abcdqwfabcZZabcdqwfab, 9
' abcdqwfabcZZabcdqwfabc, 10
' abcdqwfabcZZabcdqwfabcd, 4
' abcdqwfabcZZabcdqwfabcda, 0
' abcdqwfabcZZabcdqwfabcdab, 0
' abcdqwfabcZZabcdqwfabcdabc, 0
' abcdqwfabcZZabcdqwfabcdabcd, 0
' abcdqwfabcZZabcdqwfabcdabcdq, 0
' abcdqwfabcZZabcdqwfabcdabcdqw, 0
' abcdqwfabcZZabcdqwfabcdabcdqwf, 0
' abcdqwfabcZZabcdqwfabcdabcdqwfa, 1
' abcdqwfabcZZabcdqwfabcdabcdqwfab, 2
' abcdqwfabcZZabcdqwfabcdabcdqwfabc, 3
' abcdqwfabcZZabcdqwfabcdabcdqwfabcZ, 0
' abcdqwfabcZZabcdqwfabcdabcdqwfabcZZ, 0
' abcdqwfabcZZabcdqwfabcdabcdqwfabcZZa, 1
' abcdqwfabcZZabcdqwfabcdabcdqwfabcZZab, 2
' abcdqwfabcZZabcdqwfabcdabcdqwfabcZZabc, 3
' abcdqwfabcZZabcdqwfabcdabcdqwfabcZZabcd, 4
' abcdqwfabcZZabcdqwfabcdabcdqwfabcZZabcdq, 5
' abcdqwfabcZZabcdqwfabcdabcdqwfabcZZabcdqw, 6
' abcdqwfabcZZabcdqwfabcdabcdqwfabcZZabcdqwf, 7
' abcdqwfabcZZabcdqwfabcdabcdqwfabcZZabcdqwfa, 8
' abcdqwfabcZZabcdqwfabcdabcdqwfabcZZabcdqwfab, 9
' abcdqwfabcZZabcdqwfabcdabcdqwfabcZZabcdqwfabc, 10
' abcdqwfabcZZabcdqwfabcdabcdqwfabcZZabcdqwfabcd, 4
'
'
' ===================================================
'
Function RecursiveGetNext(strN As String, _
                          intJ As Integer)
    If intJ <= 1 Then
        '�ο�ǰ���ע��
        '���ǵݹ���ֹ������
        '����ֻ��һ���ַ���"a"
        '��Next�϶�Ϊ0
        RecursiveGetNext = 0
    Else
        '����N[j-1]��ģʽֵ
        'Next(N[j-1])
        Dim i As Integer
        i = RecursiveGetNext(strN, intJ - 1)
        If Mid(strN, i + 1, 1) = Mid(strN, intJ, 1) Then
            '��Ӧ��n[j+1] = n[i+1]
            '��Next(N[j+1]) = Next(N[j]) + 1
            RecursiveGetNext = i + 1
        Else
            Dim k As Integer
            k = RecursiveGetNext(strN, i)
            Do While k > 0 And Mid(strN, k + 1, 1) <> Mid(strN, intJ, 1)
                k = RecursiveGetNext(strN, k)
            Loop
            '
            '�ο�����Ĵ�������
            '���Ϊʲô���ﲻ��Ҫ�ж�k��ֵ
            'If k > 0 Then
                If Mid(strN, k + 1, 1) = Mid(strN, intJ, 1) Then
                    RecursiveGetNext = k + 1
                Else
                    RecursiveGetNext = 0
                End If
            'Else
                'RecursiveGetNext = 0
            'End If
        End If
    End If
    
End Function

'=============================================
'�ݹ麯���Ĵ�ӡ��������
Function printRecursiveGetNext(strString As String)
    lenStr = Len(strString)
    Dim i As Integer
    For i = 1 To lenStr
        Dim strTemp As String
        strTemp = Mid(strString, 1, i)
        Debug.Print strTemp & ", " & RecursiveGetNext(strTemp, i)
    Next i
End Function

'=============================================
'�ǵݹ�ļ���Next����
Function GetNext(strString As String)
    lenStr = Len(strString)
    ' ���ڴ�ż��������Nextֵ������
    Dim NextArray() As Integer
    ReDim NextArray(1 To lenStr)
    '��һλ�϶���0
    NextArray(1) = 0
    
    For J = 2 To lenStr
        ' Next(N[j-1])��ֵ
        i = NextArray(J - 1)
        If Mid(strString, i + 1, 1) = Mid(strString, J, 1) Then
            NextArray(J) = i + 1
        Else
            If i > 0 Then
                ' i����0
                ' ��ʾ��ǰ����һ�����Ի��ֵ�ģʽֵ
                '
                '
                ' a, 0                         j=1
                ' ab, 0                        j=2
                ' abc, 0                       j=3
                ' abcd, 0                      j=4
                ' abcdq, 0                     j=5
                ' abcdqw, 0                    j=6
                ' abcdqwf, 0                   j=7
                ' abcdqwfa, 1                  j=8
                ' abcdqwfab, 2                 j=9
                ' abcdqwfabc, 3                j=10
                ' abcdqwfabcZ, 0               j=11
                ' abcdqwfabcZZ, 0              j=12
                ' abcdqwfabcZZa, 1             j=13
                ' abcdqwfabcZZab, 2            j=14
                ' abcdqwfabcZZabc, 3           j=15
                ' abcdqwfabcZZabcd, 4          j=16
                ' abcdqwfabcZZabcdq, 5         j=17
                ' abcdqwfabcZZabcdqw, 6        j=18
                ' abcdqwfabcZZabcdqwf, 7       j=19
                ' abcdqwfabcZZabcdqwfa, 8      j=20
                ' abcdqwfabcZZabcdqwfab, 9     j=21
                ' abcdqwfabcZZabcdqwfabc, 10   j=22
                ' abcdqwfabcZZabcdqwfabcd, 4   j=23
                ' abcdqwfabcZZabcdqwfabcda, 1  j=24
                '
                '
                ' ����������,
                ' ���赱ǰ
                '         j=23,
                ' ��      j-1=22
                ' ��Ӧ��    i=10
                '
                ' ��ôNextArray(i=10)�д�ŵľ���
                '
                '
                ' N[22]
                '
                ' abcdqwfabcZZabcdqwfabc
                ' 123456789*************
                ' |        0123456789***
                ' |        |  |      012
                ' |        |  |        |
                ' ----------  ----------
                '     a            b
                '
                ' �����a���ֶ�Ӧ��Next(N[10])��ֵ����������
                '
                '
                k = NextArray(i)
                Do While k > 0 And Mid(strString, k + 1, 1) <> Mid(strString, J, 1)
                    k = NextArray(k)
                Loop
                
                If Mid(strString, k + 1, 1) = Mid(strString, J, 1) Then
                    NextArray(J) = k + 1
                Else
                    NextArray(J) = 0
                End If
                
            Else
            End If
        End If
    Next J
    GetNext = NextArray
End Function

'=============================================
'�ǵݹ麯���Ĵ�ӡ��������
Function printGetNext(strString As String)
    lenStr = Len(strString)
    Dim NextArray() As Integer
    ReDim NextArray(1 To lenStr)
    NextArray = GetNext(strString)
    
    For i = 1 To lenStr
        Dim strTemp As String
        strTemp = Mid(strString, 1, i)
        Debug.Print strTemp & ", " & NextArray(i)
    Next i
End Function





























                     

