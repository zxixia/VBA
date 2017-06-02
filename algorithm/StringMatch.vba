Private Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)
Sub StringMatch()

    Dim strString As String
    strString = "abcdefgabcdefg"
    MsgBox "查找字符串是：" & Chr(10) & strString, vbInformation, "提示"
    
    Dim strCheck As String
    
    '=====================================
    '输入要查找的字符串
     strCheck = InputBox("请输入", "请输入待查找的字符串", 1)
    
    If strCheck <> "" Then
     MsgBox "在" & Chr(10) & strString & Chr(10) & "中查找：" & Chr(10) & strCheck, vbInformation, "提示"
    Else
     MsgBox ("输入为空，请重新运行宏文件")
     Exit Sub
    End If
    '=====================================
    Debug.Print strString & Chr(10) & strCheck
    
    
    '=====================================
    '朴素算法
    Dim shotArr() As Integer
    shotArr = naiveStringMatch(strString, strCheck)
    printResult strString, strCheck, shotArr, "朴素算法"
    '=====================================
    
    
    '=====================================
    'RabinKarp算法
    shotArr = rabinKarpMatch(strString, strCheck)
    printResult strString, strCheck, shotArr, "RabinKarp算法"
    '=====================================
    
    printPreAndPostfix (strString)
End Sub
'朴素算法
Function naiveStringMatch(strString As String, _
                          strCheck As String)
'strString 是被查找的初始字符串
'strCheck  是输入的待查找字符串
    
    lenString = Len(strString)
    lenCheck = Len(strCheck)
    
    
    '一个用来存放检索到的index的数组
    Dim shotArr(1 To 10) As Integer
    Dim shotCount
    shotCount = 1
    
    'MsgBox "被查找初始字符串长度是" & lenString & " 待查找的字符串长度是" & lenCheck, vbInformation, "提示"
    For i = 1 To (lenString - lenCheck + 1)
        '调试输出中间结果,Mid函数要从1开始
        'Debug.Print Mid(strString, i + 1, 1)
        For j = 1 To lenCheck
            strA = Mid(strString, i, lenCheck)
            'Debug.Print strA & " "; Mid(strCheck, j, 1)
            If Mid(strA, j, 1) <> Mid(strCheck, j, 1) Then
                '如果两个字符串有一点不一样,跳出循环
                Exit For
            End If
            If j = lenCheck Then
                '找到子串
                'MsgBox i, vbInformation, "提示"
                shotArr(shotCount) = i
                shotCount = shotCount + 1
            End If
        Next j
    Next i
    naiveStringMatch = shotArr
End Function
'RabinKarp算法
Function rabinKarpMatch(strString As String, _
                        strCheck As String)
    lenCheck = Len(strCheck)
    lenString = Len(strString)
    Debug.Print strString & " " & Hash(strString)
    Debug.Print strCheck & " " & Hash(strCheck)
    
    '一个用来存放检索到的index的数组
    Dim shotArr(1 To 10) As Integer
    Dim shotCount
    shotCount = 1
    
    hashCheck = Hash(strCheck)
    For i = 1 To (lenString - lenCheck + 1)
        toCheck = Mid(strString, i, lenCheck)
        If Hash(toCheck) = hashCheck Then
            For j = 1 To lenCheck
                If Mid(toCheck, j, 1) <> Mid(strCheck, j, 1) Then
                    '如果两个字符串有一点不一样,跳出循环
                    Exit For
                End If
                If j = lenCheck Then
                    '找到子串
                    'MsgBox i, vbInformation, "提示"
                    shotArr(shotCount) = i
                    shotCount = shotCount + 1
                End If
            Next j
        End If
    Next i
    rabinKarpMatch = shotArr
End Function
'模仿hash函数,其实就是计算一个字符串输入的所有字符的ASCII的值的和
Function Hash(strInput)
    Dim count As Integer
    For i = 1 To Len(strInput)
        count = count + Asc(Mid(strInput, i, 1))
    Next i
    Hash = count
End Function

'输出结果的方法
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
    '补充最后的字符串
    strResult = strResult & Mid(strString, start, lenString - Len(strResult))
    MsgBox strString & Chr(10) & strResult, vbInformation, strTitle
    Debug.Print strString
    Debug.Print strResult
End Function
'输出一个字符串的前缀和后缀
'便于调试
Function printPreAndPostfix(strString As String)
    Dim prefix As String
    Dim postfix As String
    
    For i = 1 To Len(strString) - 1
        prefix = prefix & " " & Mid(strString, 1, i)
        '下面这个是逆序的输出
        'postfix = postfix & " " & Mid(strString, Len(strString) - i + 1, i)
        postfix = postfix & " " & Mid(strString, i + 1, Len(strString) - i)
    Next i
    Debug.Print prefix
    Debug.Print postfix
    MsgBox "前缀:" & prefix & Chr(10) & "后缀:" & postfix, vbInformation, "前后缀"
End Function


' 定义
' a) N[j]          表示长度为j的字符串
' b) n[0...(j-1)]  表示长度为j的字符串对应于0...(j-1)位处的具体取值
' c) Next(N[j])    表示长度为j的字符串对应的模式值
' =================================================================

' 模式值数组的求取
' 参考网页：http://www.ituring.com.cn/article/59881
' N[4] = abca , Next(N[4]) =  0,  表示n[0] = n[3]
' N[5] = abcab, Next(N[5]) =  1,  表示n[0] = n[4], n[1] = n[3]
' N[3] = abc  , Next(N[3]) = -1,  表示没有首尾匹配


' 假设已知字符串N[j]对应的模式值, Next(N(j)) = i,则其可视化表示如下图所示
'
' N[j] = abcdqwfabc
'        0123456789
'          i
'
' 由上图可知
' j=10,  表示该字符串的长度
' i=2,   表示该字符串最长模式匹配的前缀的最末位的坐标,也即表示
'        n[0...i] == n[(j-1-i)...(j-1)]
'        n[0] = n[7] = "a"
'        n[1] = n[8] = "b"
'        n[2] = n[9] = "c"
'


' 假设已知字符串N[j]的模式值,Next(N[j]) = i
' 下面将分【3】种情况来讨论如何求解字符串N[j+1]的模式值,Next(N[j+1])的值

' 【1】,
'  n[j] == n[i+1]的情况
'
' 设字符串 N[11]="abcdqwfabcd"
' 其子串   N[10]="abcdqwfabc", 且Next(N[10]) = 2 , 其j=10, 其i=2
' 同时     (n[j] = n[10]) == (n[i+1] = n[3]) = "d"
' 则推出   Next(N[11]) = Next(N[10]) + 1
' 也即是
' 如果     n[j] == n[i+1]
' 则       Next(N[j+1]) = Next(N[10]) + 1
'                       = i+1
'
'
'
' 【2】,
' n[j] != n[i+1]的情况




'假设字符串N[j]已经求出了next的值,next(N[j]) = i}，
'如何求N[j+1]的next值,下面分3种情况进行讨论


'1,
'假设这个数组N[12]是"abcdqwerabcd",共有12个字符
'则N[10]{对应的字符串是"abcdqwerab"}的next[10]=1,表示前缀"ab"和后缀"ab"





                     

