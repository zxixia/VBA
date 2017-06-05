Sub Hannoi()
'汉诺塔
    Dim iOne(3) As Integer
    Dim iTwo(3) As Integer
    Dim iThree(3) As Integer
    iOne(1) = 0
    iOne(2) = 2
    iOne(3) = 3
    iTwo(1) = 0
    iTwo(2) = 0
    iTwo(3) = 1
    iThree(1) = 0
    iThree(2) = 1
    iThree(3) = 3
    printHannnoi 1, 3, iOne, iTwo, iThree
End Sub

'打印当前汉诺塔函数
'先画一个汉诺塔的示意图
'
'
' 初始
'
'   -     !     !
'  ---    !     !
' -----   !     !
'
'
' 第1步
'
'   !     !     !
'  ---    !     !
' -----   -     !
'
'
' 第2步
'
'   !     !     !
'   !     !     !
' -----   -    ---
'
'
' 第3步
'
'   !     !     !
'   !     !     -
' -----   !    ---
'
'
' 第4步
'
'   !     !     !
'   !     !     -
'   !   -----  ---
'
'
' 第5步
'
'   !     !     !
'   !     -     !
'   !   -----  ---
'
'
' 第6步
'
'   !     !     !
'   !     -     !
'  ---  -----   !
'
'
' 第7步
'
'   !     !     !
'   -     !     !
'  ---  -----   !
'
'
' 第8步
'
'   !     !     !
'   -     !     !
'  ---    !   -----
'
'
' 第9步
'
'   !     !     !
'   !     !     !
'  ---    -   -----
'
'
' 第10步
'
'   !     !     !
'   !     !    ---
'   !     -   -----
'
'
' 第11步
'
'   !     !     -
'   !     !    ---
'   !     !   -----
'
Function printHannnoi(iStep As Integer, _
                      iHannoiCount As Integer, _
                      iOne() As Integer, _
                      iTwo() As Integer, _
                      iThree() As Integer)
' iStep        打印第一行的第几步
' iHannoiCount 当前汉诺塔的饼数目
' iOne()       第1个柱子从上到下的饼数目
' iTwo()       第2个柱子从上到下的饼数目
' iThree()     第3个柱子从上到下的饼数目
'
'            1 2 3
' iOne()   = 0,1,2
' iTwo()   = 0,0,0
' iThree() = 0,0,3
'
'1    !     !     !
'2    -     !     !
'3   ---    !   -----
'
'
 Debug.Print
 Debug.Print "第" & iStep & "步"
 Debug.Print
 For i = 1 To iHannoiCount
    ' 逐层打印当前的汉诺塔布局
    Dim one As Integer
    Dim two As Integer
    Dim three As Integer
    one = iOne(i)
    two = iTwo(i)
    three = iThree(i)
    Debug.Print " " & generateHannoi(iHannoiCount, one) & " " & generateHannoi(iHannoiCount, two) & " " & generateHannoi(iHannoiCount, three)
 Next i
End Function

Function generateHannoi(iHannoiCount As Integer, _
                        i As Integer)
    maxCount = iHannoiCount * 2 - 1
    If i > 0 Then
        iCount = i * 2 - 1
        iSpace = (maxCount - iCount) / 2
        generateHannoi = String(iSpace, " ") & String(iCount, "-") & String(iSpace, " ")
    Else
        iSpace = (maxCount - 1) / 2
        generateHannoi = String(iSpace, " ") & String(1, "!") & String(iSpace, " ")
    End If
End Function

'
'
'
'
'
'
'
'
'
'
'
'
'
'
'
'
'
'
'
'
'
'
'
'
'
'
'
'
'
'
'
'
'
'
'
'
'
'
'
'
'
'
'
'
'
'
'
'
'
'
'
