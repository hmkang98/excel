Attribute VB_Name = "Module1"
Function Prime(num) As String
    Dim n As Integer
    If Not IsNumeric(num) Then
        Prime = "���ڴ� �Էµ��� �ʽ��ϴ�."
        Exit Function
    End If
    n = num
   If n <> num Then
        Prime = "�ùٸ� ���ڸ� �Է��ϼ���."
    ElseIf n < 1 Then
        Prime = "���� ������ �Է��� �����մϴ�."
    ElseIf (n = 1) Then
        Prime = "1�� �Ҽ��� �ռ����� �ƴմϴ�."
    ElseIf (n > 1) Then
        For i = 2 To n - 1 Step 1
            If (n Mod i = 0) Then
                Prime = n & "��(��) " & i & "�� ������ �Ҽ��� �ƴմϴ�."
                Exit Function
            End If
        Next i
        Prime = n & "��(��) �Ҽ��Դϴ�."
    End If
End Function
