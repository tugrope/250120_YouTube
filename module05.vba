Public Function NormalizeJapaneseText(ByVal inputText As String) As String
    Dim normalizedText As String
    Dim i As Long
    Dim currentChar As String
    Dim nextChar As String
    Dim unicodeVal As Long

    '全角スペースを半角スペースに変換
    normalizedText = Replace(inputText, "　", " ")

    '全角英字、全角数字、半角カタカナを変換
    For i = 1 To Len(normalizedText)
        currentChar = Mid(normalizedText, i, 1)
        unicodeVal = AscW(currentChar)

        If (unicodeVal >= &HFF21 And unicodeVal <= &HFF3A) Or (unicodeVal >= &HFF41 And unicodeVal <= &HFF5A) Then
            ' 全角英字を半角に変換
            Mid(normalizedText, i, 1) = ChrW(unicodeVal - &HFEE0)
        ElseIf unicodeVal >= &HFF10 And unicodeVal <= &HFF19 Then
            ' 全角数字を半角に変換
            Mid(normalizedText, i, 1) = ChrW(unicodeVal - &HFEE0)
        ElseIf unicodeVal >= &HFF65 And unicodeVal <= &HFF9F Then
            ' 半角カタカナを全角に変換
            Mid(normalizedText, i, 1) = StrConv(currentChar, vbWide)
        ElseIf unicodeVal = &HFF0D Or unicodeVal = &H30FC Then
            ' 長音記号を統一
            Mid(normalizedText, i, 1) = ChrW(&H30FC)
        End If
    Next i

    '連続する半角スペースを単一の半角スペースに変換
    Do While InStr(normalizedText, "  ") > 0
        normalizedText = Replace(normalizedText, "  ", " ")
    Loop

    NormalizeJapaneseText = normalizedText
End Function
