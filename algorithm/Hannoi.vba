Sub Enter()
'汉诺塔
    Dim iOne(4) As Integer
    Dim iTwo(4) As Integer
    Dim iThree(4) As Integer
    iOne(1) = 0
    iOne(2) = 0
    iOne(3) = 3
    iOne(4) = 4
    
    iTwo(1) = 0
    iTwo(2) = 0
    iTwo(3) = 1
    iTwo(4) = 2
    
    iThree(1) = 0
    iThree(2) = 0
    iThree(3) = 0
    iThree(4) = 0
    printHannnoi 1, 4, iOne, iTwo, iThree
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
' -----   !     -
'
'
' 第2步
'
'   !     !     !
'   !     !     !
' -----  ---    -
'
'
' 第3步
'
'   !     !     !
'   !     -     !
' -----  ---    !
'
'
' 第4步
'
'   !     !     !
'   !     -     !
'   !    ---  -----
'
'
' 第5步
'
'   !     !     !
'   !     !     !
'   -    ---  -----
'
'
' 第6步
'
'   !     !     !
'   !     !    ---
'   -     !   -----
'
'
' 第7步
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
' 汉诺塔递归构造函数
'
' 定义
'               1   2   3
'          饼  塔  塔  塔
' 1，Hannoi(N, t1, t2, t3)，表示将N个饼借助于 t2，从 t1 移动到 t3
'
'
'             F   T
'            塔  塔
' 2，Move(n, tF, tT)，表示将第n个饼从 tF 的塔顶移动到 tT 的塔顶
'
'
' 假设有N个饼
'
' 1，Hannoi(N, t1, t2, t3)，表示将N个饼，借助于t2, 从t1 移动到 t3
'
' 2，把（N-1）个饼, 借助于 t3, 从 t1 移动到 t2
'    Hannoi(N-1, t1, t3, t2)
'
' 3，把第n个饼从 t1 移动到 t3
'    Move(n, t1, t3)
'
' 4，把（N-1）个饼, 借助于 t1, 从 t2 移动到 t3
'    Hannoi(N-1, t2, t1, t3)
'
'
' 假设只有一个饼
'
' 1，Move(1, t1, t3)，直接将该饼从t1 的塔顶,移动到 t3的塔顶即可
'
'
' 基于上述推理，得到如下的递归伪代码
'
'
' Hannoi(N, t1, t2, t3) {
'   if N == 1 {
'       Move(1, t1, t3);
'   } else {
'       Hannoi(N-1, t1, t3, t2);
'       Move(n, t1, t3);
'       Hannoi(N-1, t2, t1, t3);
'   }
' }
'
'
'
Function Hannoi(iStep As Integer, _
                iHannoiCount As Integer, _
                iOne() As Integer, _
                iTwo() As Integer, _
                iThree() As Integer)

    
End Function
' 汉诺塔的移动
Function moveHannoi(iNo As Integer, _
                    iFrom As Integer, _
                    iTo As Integer)

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
