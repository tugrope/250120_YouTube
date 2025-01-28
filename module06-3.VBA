Function NormalizeText(targetCell As Range) As String
    Dim inputText As String
    Dim i As Long
    Dim resultText As String
    Dim currentChar As String
    Dim prevChar As String

    ' 入力セルからテキストを取得
    inputText = targetCell.Text

    ' まず半角カタカナを全角カタカナに変換
    resultText = ConvertToFullWidthKana(inputText)

    ' 一時変数を初期化
    Dim tempResult As String
    tempResult = ""
    prevChar = ""

    ' 文字列を1文字ずつ処理
    For i = 1 To Len(resultText)
        currentChar = Mid(resultText, i, 1)

        ' まず半角英数字を全角に変換し、ハイフン類似文字を処理
        Select Case AscW(currentChar)
            Case 48 To 57  ' 半角数字 (0-9)
                tempResult = tempResult & StrConv(currentChar, vbWide)
            Case 65 To 90  ' 半角英大文字 (A-Z)
                tempResult = tempResult & StrConv(currentChar, vbWide)
            Case 97 To 122 ' 半角英小文字 (a-z)
                tempResult = tempResult & StrConv(currentChar, vbWide)
            Case &H2D, &H2212, &HFF0D, &H2014, &H2015 ' ハイフン類似文字
                tempResult = tempResult & ChrW(&H30FC) ' 全角長音記号に変換
            Case Else      ' その他の文字はそのまま
                tempResult = tempResult & currentChar
        End Select

        ' 現在の文字を前の文字として保存
        prevChar = Right(tempResult, 1)
    Next i

    NormalizeText = tempResult
End Function
