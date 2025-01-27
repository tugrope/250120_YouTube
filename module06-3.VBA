Function NormalizeText(targetCell As Range) As String
    Dim inputText As String
    Dim i As Long
    Dim resultText As String
    Dim currentChar As String

    ' 入力セルからテキストを取得
    inputText = targetCell.Text

    ' まず半角カタカナを全角カタカナに変換
    resultText = ConvertToFullWidthKana(inputText)

    ' 一時変数を初期化
    Dim tempResult As String
    tempResult = ""

    ' 文字列を1文字ずつ処理
    For i = 1 To Len(resultText)
        currentChar = Mid(resultText, i, 1)

        ' ハイフン類似文字の判定と変換
        Select Case AscW(currentChar)
            Case &H2D     ' 半角ハイフンマイナス -
                tempResult = tempResult & ChrW(&H30FC)
            Case &H2212   ' 全角マイナス −
                tempResult = tempResult & ChrW(&H30FC)
            Case &HFF0D   ' 全角ハイフンマイナス －
                tempResult = tempResult & ChrW(&H30FC)
            Case &H2014   ' 全角ダッシュ —
                tempResult = tempResult & ChrW(&H30FC)
            Case &H2015   ' 全角ホリゾンタルバー ―
                tempResult = tempResult & ChrW(&H30FC)
            Case 48 To 57  ' 半角数字 (0-9)
                tempResult = tempResult & StrConv(currentChar, vbWide)
            Case 65 To 90  ' 半角英大文字 (A-Z)
                tempResult = tempResult & StrConv(currentChar, vbWide)
            Case 97 To 122 ' 半角英小文字 (a-z)
                tempResult = tempResult & StrConv(currentChar, vbWide)
            Case Else      ' その他の文字はそのまま
                tempResult = tempResult & currentChar
        End Select
    Next i

    NormalizeText = tempResult
End Function
