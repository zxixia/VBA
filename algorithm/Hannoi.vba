' ȫ�ֱ���
' ��¼��ǰ����
Public StepCount As Integer
Sub Enter()
    StepCount = 1
'��ŵ��
    Dim t1() As Integer
    Dim t2() As Integer
    Dim t3() As Integer
    Dim num As Integer
    '����Ҫ���ҵ��ַ���
    num = InputBox("������", "�����뺺ŵ���ı���", "3")
    
    InitHannoi num, t1, t2, t3
    Hannoi num, t1, t2, t3
End Sub

'��ӡ��ǰ��ŵ������
'�Ȼ�һ����ŵ����ʾ��ͼ
'
'
' ��ʼ
'
'   -     !     !
'  ---    !     !
' -----   !     !
'
'
' ��1��
'
'   !     !     !
'  ---    !     !
' -----   !     -
'
'
' ��2��
'
'   !     !     !
'   !     !     !
' -----  ---    -
'
'
' ��3��
'
'   !     !     !
'   !     -     !
' -----  ---    !
'
'
' ��4��
'
'   !     !     !
'   !     -     !
'   !    ---  -----
'
'
' ��5��
'
'   !     !     !
'   !     !     !
'   -    ---  -----
'
'
' ��6��
'
'   !     !     !
'   !     !    ---
'   -     !   -----
'
'
' ��7��
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
' iStep        ��ӡ��һ�еĵڼ���
' iHannoiCount ��ǰ��ŵ���ı���Ŀ
' iOne()       ��1�����Ӵ��ϵ��µı���Ŀ
' iTwo()       ��2�����Ӵ��ϵ��µı���Ŀ
' iThree()     ��3�����Ӵ��ϵ��µı���Ŀ
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
 Debug.Print "��" & iStep & "��"
 Debug.Print
 For i = 1 To iHannoiCount
    ' ����ӡ��ǰ�ĺ�ŵ������
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
' ��ŵ���ݹ鹹�캯��
'
' ����
'               1   2   3
'          ��  ��  ��  ��
' 1��Hannoi(N, t1, t2, t3)����ʾ��N���������� t2���� t1 �ƶ��� t3
'
'
'          F   T
'         ��  ��
' 2��Move(tF, tT)����ʾ�� tF ������Ԫ���ƶ��� tT ������
'
'
' ������N����
'
' 1��Hannoi(N, t1, t2, t3)����ʾ��N������������t2, ��t1 �ƶ��� t3
'
' 2���ѣ�N-1������, ������ t3, �� t1 �ƶ��� t2
'    Hannoi(N-1, t1, t3, t2)
'
' 3���ѵ�n������ t1 �������ƶ��� t3 ������
'    Move(t1, t3)
'
' 4���ѣ�N-1������, ������ t1, �� t2 �ƶ��� t3
'    Hannoi(N-1, t2, t1, t3)
'
'
' ����ֻ��һ����
'
' 1��Move(t1, t3)��ֱ�ӽ��ñ���t1 ������,�ƶ��� t3����������
'
'
' �������������õ����µĵݹ�α����
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
' ע��,Ϊ�����ֲ�ͬ����,
' ������ĳ���ΪN+1
' ���еڣ�N+1��λ�ŵ���1,2,3 ����������t1,t2,t3
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
' ��ŵ�����ƶ�
Function Move(tFrom() As Integer, _
              tTo() As Integer)
' ��tFrom ������Ԫ��
' �ƶ���tTo ������
 
  ' ���ģ�N+1�� �ŵ���1,2,3 ʹ��������t1,t2,t3�Ĺؼ�
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
' ��ӡ������
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
' ���ɺ�ŵ��������
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
