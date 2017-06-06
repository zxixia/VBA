' 全局变量
' 记录当前步数
Public StepCount As Integer
Sub Enter()
    StepCount = 1
'汉诺塔
    Dim t1() As Integer
    Dim t2() As Integer
    Dim t3() As Integer
    Dim num As Integer
    '输入要查找的字符串
    num = InputBox("请输入", "请输入汉诺塔的饼数", "3")
    
    InitHannoi num, t1, t2, t3
    Hannoi num, t1, t2, t3
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
'          F   T
'         塔  塔
' 2，Move(tF, tT)，表示将 tF 的塔顶元素移动到 tT 的塔顶
'
'
' 假设有N个饼
'
' 1，Hannoi(N, t1, t2, t3)，表示将N个饼，借助于t2, 从t1 移动到 t3
'
' 2，把（N-1）个饼, 借助于 t3, 从 t1 移动到 t2
'    Hannoi(N-1, t1, t3, t2)
'
' 3，把第n个饼从 t1 的塔顶移动到 t3 的塔顶
'    Move(t1, t3)
'
' 4，把（N-1）个饼, 借助于 t1, 从 t2 移动到 t3
'    Hannoi(N-1, t2, t1, t3)
'
'
' 假设只有一个饼
'
' 1，Move(t1, t3)，直接将该饼从t1 的塔顶,移动到 t3的塔顶即可
'
'
' 基于上述推理，得到如下的递归伪代码
'
'
' Hannoi(N, t1, t2, t3) {
'   if N == 1 {
'       Move(t1, t3);
'   } else {
'       Hannoi(N-1, t1, t3, t2);
'       Move(t1, t3);
'       Hannoi(N-1, t2, t1, t3);
'   }
' }
'
'
'
' 注意,为了区分不同的塔,
' 这个塔的长度为N+1
' 其中第（N+1）位放的是1,2,3 用来区分是t1,t2,t3
Function Hannoi(N As Integer, _
                t1() As Integer, _
                t2() As Integer, _
                t3() As Integer)

    If N = 1 Then
        Move t1, t3
        printHelper t1, t2, t3
    Else
        Hannoi (N - 1), t1, t3, t2
        
        Move t1, t3
        printHelper t1, t2, t3
        
        Hannoi (N - 1), t2, t1, t3
    End If
End Function
' 汉诺塔的移动
Function Move(tFrom() As Integer, _
              tTo() As Integer)
' 将tFrom 塔顶的元素
' 移动到tTo 的塔顶
 
  ' 塔的（N+1） 放的是1,2,3 使用来区分t1,t2,t3的关键
  arrayLen = UBound(tFrom) - 1
  Top = 0
  
  For i = 1 To arrayLen
    If tFrom(i) > 0 Then
        Top = tFrom(i)
        tFrom(i) = 0
        Exit For
    End If
  Next i
  
  For j = arrayLen To 1 Step -1
    If tTo(j) < 1 Then
        tTo(j) = Top
        Exit For
    End If
  Next j
End Function

'
' 打印帮助类
Function printHelper(t1() As Integer, _
                     t2() As Integer, _
                     t3() As Integer)
    tagIndex = UBound(t1)
    Dim hannoiSize As Integer
    hannoiSize = tagIndex - 1
    
    t1Tag = t1(tagIndex)
    t2Tag = t2(tagIndex)
    t3Tag = t3(tagIndex)
    
    If t1Tag = 1 And t2Tag = 2 And t3Tag = 3 Then
        printHannnoi StepCount, hannoiSize, t1, t2, t3
    End If
    
    If t1Tag = 1 And t2Tag = 3 And t3Tag = 2 Then
        printHannnoi StepCount, hannoiSize, t1, t3, t2
    End If
    
    If t1Tag = 2 And t2Tag = 1 And t3Tag = 3 Then
        printHannnoi StepCount, hannoiSize, t2, t1, t3
    End If
    
    If t1Tag = 2 And t2Tag = 3 And t3Tag = 1 Then
        printHannnoi StepCount, hannoiSize, t3, t1, t2
    End If
    
    If t1Tag = 3 And t2Tag = 1 And t3Tag = 2 Then
        printHannnoi StepCount, hannoiSize, t2, t3, t1
    End If
    
    If t1Tag = 3 And t2Tag = 2 And t3Tag = 1 Then
        printHannnoi StepCount, hannoiSize, t3, t2, t1
    End If
    
    StepCount = StepCount + 1
End Function
'
'
' 生成汉诺塔帮助类
Function InitHannoi(size As Integer, _
                    t1() As Integer, _
                    t2() As Integer, _
                    t3() As Integer)
                    
   tagIndex = size + 1
   ReDim t1(1 To tagIndex)
   t1(tagIndex) = 1
   
   ReDim t2(1 To tagIndex)
   t2(tagIndex) = 2
   
   ReDim t3(1 To tagIndex)
   t3(tagIndex) = 3
   
   For i = 1 To size
      t1(i) = i
      t2(i) = 0
      t3(i) = 0
   Next i
End Function
