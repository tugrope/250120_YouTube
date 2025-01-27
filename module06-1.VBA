Function NormalizeText(targetCell As Range) As String
    Dim inputText As String
    Dim i As Long
    Dim resultText As String
    Dim currentChar As String

    ' 入力セルからテキストを取得
    inputText = targetCell.Text
    resultText = ""

    ' 文字列を1文字ずつ処理
    For i = 1 To Len(inputText)
        currentChar = Mid(inputText, i, 1)

        ' 半角英数字の判定と変換
        Select Case Asc(currentChar)
            Case 48 To 57  ' 半角数字 (0-9)
                resultText = resultText & StrConv(currentChar, vbWide)
            Case 65 To 90  ' 半角英大文字 (A-Z)
                resultText = resultText & StrConv(currentChar, vbWide)
            Case 97 To 122 ' 半角英小文字 (a-z)
                resultText = resultText & StrConv(currentChar, vbWide)
            Case Else      ' その他の文字はそのまま
                resultText = resultText & currentChar
        End Select
    Next i

    NormalizeText = resultText
End Function
