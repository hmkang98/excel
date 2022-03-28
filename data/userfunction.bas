Attribute VB_Name = "Module1"
Function Prime(num) As String
    Dim n As Integer
    If Not IsNumeric(num) Then
        Prime = "문자는 입력되지 않습니다."
        Exit Function
    End If
    n = num
   If n <> num Then
        Prime = "올바른 숫자를 입력하세요."
    ElseIf n < 1 Then
        Prime = "양의 정수만 입력이 가능합니다."
    ElseIf (n = 1) Then
        Prime = "1은 소수도 합성수도 아닙니다."
    ElseIf (n > 1) Then
        For i = 2 To n - 1 Step 1
            If (n Mod i = 0) Then
                Prime = n & "은(는) " & i & "로 나누어 소수가 아닙니다."
                Exit Function
            End If
        Next i
        Prime = n & "은(는) 소수입니다."
    End If
End Function
