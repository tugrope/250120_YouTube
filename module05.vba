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
        ElseIf unicodeVal >= &HFF66 And unicodeVal <= &HFF9F Then
            ' 半角カタカナを全角に変換
            nextChar = Mid(normalizedText, i + 1, 1) ' 次の文字を取得
            If nextChar = ChrW(&HFF9E) Or nextChar = ChrW(&HFF9F) Then
                ' 濁点・半濁点が続く場合
                Select Case currentChar
                    Case "ｶ": Mid(normalizedText, i, 1) = "ガ"
                    Case "ｷ": Mid(normalizedText, i, 1) = "ギ"
                    Case "ｸ": Mid(normalizedText, i, 1) = "グ"
                    Case "ｹ": Mid(normalizedText, i, 1) = "ゲ"
                    Case "ｺ": Mid(normalizedText, i, 1) = "ゴ"
                    Case "ｻ": Mid(normalizedText, i, 1) = "ザ"
                    Case "ｼ": Mid(normalizedText, i, 1) = "ジ"
                    Case "ｽ": Mid(normalizedText, i, 1) = "ズ"
                    Case "ｾ": Mid(normalizedText, i, 1) = "ゼ"
                    Case "ｿ": Mid(normalizedText, i, 1) = "ゾ"
                    Case "ﾀ": Mid(normalizedText, i, 1) = "ダ"
                    Case "ﾁ": Mid(normalizedText, i, 1) = "ヂ"
                    Case "ﾂ": Mid(normalizedText, i, 1) = "ヅ"
                    Case "ﾃ": Mid(normalizedText, i, 1) = "デ"
                    Case "ﾄ": Mid(normalizedText, i, 1) = "ド"
                    Case "ﾊ": Mid(normalizedText, i, 1) = "バ"
                    Case "ﾋ": Mid(normalizedText, i, 1) = "ビ"
                    Case "ﾌ": Mid(normalizedText, i, 1) = "ブ"
                    Case "ﾍ": Mid(normalizedText, i, 1) = "ベ"
                    Case "ﾎ": Mid(normalizedText, i, 1) = "ボ"
                    Case "ﾊﾟ": Mid(normalizedText, i, 1) = "パ"
                    Case "ﾋﾟ": Mid(normalizedText, i, 1) = "ピ"
                    Case "ﾌﾟ": Mid(normalizedText, i, 1) = "プ"
                    Case "ﾍﾟ": Mid(normalizedText, i, 1) = "ペ"
                    Case "ﾎﾟ": Mid(normalizedText, i, 1) = "ポ"
                End Select
                ' 濁点・半濁点を削除し、次の文字をスキップする処理
                Mid(normalizedText, i + 1, 1) = "" ' 濁点・半濁点を空文字に置換
                i = i + 1  ' 次の文字（濁点・半濁点）を処理対象から除外
            Else
                ' 濁点・半濁点が続かない場合は1文字ずつ変換
                Mid(normalizedText, i, 1) = StrConv(currentChar, vbWide)
            End If
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
