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

    ' ' ハイフン類似文字の処理を別ループで実行
    ' Dim finalResult As String
    ' finalResult = ""
    ' For i = 1 To Len(tempResult)
    '     currentChar = Mid(tempResult, i, 1)

    '     ' ハイフン類似文字の判定と変換
    '     Select Case AscW(currentChar)
    '         Case &H2D, &H2212, &HFF0D, &H2014, &H2015
    '             finalResult = finalResult & ChrW(&H30FC)
    '         Case &H30FC
    '             ' 直前の文字が全角英数字の場合は半角ハイフンに変換
    '             If prevChar <> "" And Len(prevChar) > 0 Then
    '                 On Error Resume Next
    '                 Dim prevCharCode As Long
    '                 prevCharCode = AscW(prevChar)
    '                 If Err.Number = 0 Then
    '                     If (prevCharCode >= &HFF10 And prevCharCode <= &HFF19) Or _
    '                     (prevCharCode >= &HFF21 And prevCharCode <= &HFF3A) Or _
    '                     (prevCharCode >= &HFF41 And prevCharCode <= &HFF5A) Then
    '                         finalResult = finalResult & "-"
    '                     Else
    '                         finalResult = finalResult & "ー"
    '                     End If
    '                 Else
    '                     finalResult = finalResult & "ー"
    '                 End If
    '                 On Error GoTo 0
    '             Else
    '                 finalResult = finalResult & "ー"
    '             End If
    '         Case Else
    '             finalResult = finalResult & currentChar
    '     End Select

    '     ' 現在の文字を前の文字として保存
    '     prevChar = Right(finalResult, 1)
    ' Next i

    NormalizeText = tempResult
End Function
