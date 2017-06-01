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
    Dim shotArr() As Integer
    shotArr = naiveStringMatch(strString, strCheck)
    printResult strString, strCheck, shotArr
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

Function printResult(strString As String, _
                     strCheck As String, _
                     shotArr() As Integer)
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
        If start < i Then
            strResult = strResult & Mid(strString, start, i - Len(strResult) - 1) & UCase(strCheck)
        End If
        start = i + lenCheck
    Next i
    '补充最后的字符串
    strResult = strResult & Mid(strString, start, lenString - Len(strResult))
    MsgBox strString & Chr(10) & strResult, vbInformation, "提示"
    Debug.Print strString
    Debug.Print strResult
End Function
                     

