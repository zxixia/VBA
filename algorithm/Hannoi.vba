Sub Enter()
'��ŵ��
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
'             F   T
'            ��  ��
' 2��Move(n, tF, tT)����ʾ����n������ tF �������ƶ��� tT ������
'
'
' ������N����
'
' 1��Hannoi(N, t1, t2, t3)����ʾ��N������������t2, ��t1 �ƶ��� t3
'
' 2���ѣ�N-1������, ������ t3, �� t1 �ƶ��� t2
'    Hannoi(N-1, t1, t3, t2)
'
' 3���ѵ�n������ t1 �ƶ��� t3
'    Move(n, t1, t3)
'
' 4���ѣ�N-1������, ������ t1, �� t2 �ƶ��� t3
'    Hannoi(N-1, t2, t1, t3)
'
'
' ����ֻ��һ����
'
' 1��Move(1, t1, t3)��ֱ�ӽ��ñ���t1 ������,�ƶ��� t3����������
'
'
' �������������õ����µĵݹ�α����
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
' ��ŵ�����ƶ�
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
