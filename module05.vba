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

            ' 濁点・半濁点付きの文字を適切に変換
            Dim convertedChar As String
            convertedChar = ConvertHankakuKana(currentChar, nextChar)

            If Len(convertedChar) > 0 Then
                Mid(normalizedText, i, 1) = convertedChar
                If nextChar = ChrW(&HFF9E) Or nextChar = ChrW(&HFF9F) Then
                    ' 濁点・半濁点を削除し、既に変換後の文字に組み込み済み
                    Mid(normalizedText, i + 1, 1) = ""
                    i = i + 1
                End If
            Else
                ' 濁点・半濁点がない場合の通常の変換
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

' 半角カタカナを全角カタカナに変換する補助関数
Private Function ConvertHankakuKana(currentChar As String, nextChar As String) As String
    ' 濁点・半濁点のチェック
    Dim hasDakuten As Boolean
    Dim hasHandakuten As Boolean

    hasDakuten = (nextChar = ChrW(&HFF9E))    ' 濁点
    hasHandakuten = (nextChar = ChrW(&HFF9F))  ' 半濁点

    ' 文字変換テーブル
    Select Case currentChar
        ' カ行
        Case "ｶ": ConvertHankakuKana = IIf(hasDakuten, "ガ", "カ")
        Case "ｷ": ConvertHankakuKana = IIf(hasDakuten, "ギ", "キ")
        Case "ｸ": ConvertHankakuKana = IIf(hasDakuten, "グ", "ク")
        Case "ｹ": ConvertHankakuKana = IIf(hasDakuten, "ゲ", "ケ")
        Case "ｺ": ConvertHankakuKana = IIf(hasDakuten, "ゴ", "コ")

        ' サ行
        Case "ｻ": ConvertHankakuKana = IIf(hasDakuten, "ザ", "サ")
        Case "ｼ": ConvertHankakuKana = IIf(hasDakuten, "ジ", "シ")
        Case "ｽ": ConvertHankakuKana = IIf(hasDakuten, "ズ", "ス")
        Case "ｾ": ConvertHankakuKana = IIf(hasDakuten, "ゼ", "セ")
        Case "ｿ": ConvertHankakuKana = IIf(hasDakuten, "ゾ", "ソ")

        ' タ行
        Case "ﾀ": ConvertHankakuKana = IIf(hasDakuten, "ダ", "タ")
        Case "ﾁ": ConvertHankakuKana = IIf(hasDakuten, "ヂ", "チ")
        Case "ﾂ": ConvertHankakuKana = IIf(hasDakuten, "ヅ", "ツ")
        Case "ﾃ": ConvertHankakuKana = IIf(hasDakuten, "デ", "テ")
        Case "ﾄ": ConvertHankakuKana = IIf(hasDakuten, "ド", "ト")

        ' ハ行
        Case "ﾊ": ConvertHankakuKana = IIf(hasDakuten, "バ", IIf(hasHandakuten, "パ", "ハ"))
        Case "ﾋ": ConvertHankakuKana = IIf(hasDakuten, "ビ", IIf(hasHandakuten, "ピ", "ヒ"))
        Case "ﾌ": ConvertHankakuKana = IIf(hasDakuten, "ブ", IIf(hasHandakuten, "プ", "フ"))
        Case "ﾍ": ConvertHankakuKana = IIf(hasDakuten, "ベ", IIf(hasHandakuten, "ペ", "ヘ"))
        Case "ﾎ": ConvertHankakuKana = IIf(hasDakuten, "ボ", IIf(hasHandakuten, "ポ", "ホ"))

        Case Else: ConvertHankakuKana = ""
    End Select
End Function
