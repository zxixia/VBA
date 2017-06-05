Sub Hannoi()
'��ŵ��
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
' -----   -     !
'
'
' ��2��
'
'   !     !     !
'   !     !     !
' -----   -    ---
'
'
' ��3��
'
'   !     !     !
'   !     !     -
' -----   !    ---
'
'
' ��4��
'
'   !     !     !
'   !     !     -
'   !   -----  ---
'
'
' ��5��
'
'   !     !     !
'   !     -     !
'   !   -----  ---
'
'
' ��6��
'
'   !     !     !
'   !     -     !
'  ---  -----   !
'
'
' ��7��
'
'   !     !     !
'   -     !     !
'  ---  -----   !
'
'
' ��8��
'
'   !     !     !
'   -     !     !
'  ---    !   -----
'
'
' ��9��
'
'   !     !     !
'   !     !     !
'  ---    -   -----
'
'
' ��10��
'
'   !     !     !
'   !     !    ---
'   !     -   -----
'
'
' ��11��
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
