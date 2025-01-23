Public Function NormalizeJapaneseText(ByVal inputText As String) As String
    Dim normalizedText As String
    Dim i As Long
    Dim currentChar As String
    Dim nextChar As String

    '全角スペースを半角スペースに変換
    normalizedText = Replace(inputText, "　", " ")

    '全角英数字を半角英数字に変換
    For i = 1 To Len(normalizedText)
        currentChar = Mid(normalizedText, i, 1)
        Select Case currentChar
            '全角数字の変換（０-９ → 0-9）
            Case "０" To "９"
                Mid(normalizedText, i, 1) = ChrW(AscW(currentChar) - &HFEE0)
            '全角英字大文字の変換（Ａ-Ｚ → A-Z）
            Case "Ａ" To "Ｚ"
                Mid(normalizedText, i, 1) = ChrW(AscW(currentChar) - &HFEE0)
            '全角英字小文字の変換（ａ-ｚ → a-z）
            Case "ａ" To "ｚ"
                Mid(normalizedText, i, 1) = ChrW(AscW(currentChar) - &HFEE0)
        End Select
    Next i

    '半角カタカナを全角カタカナに変換
    For i = 1 To Len(normalizedText)
        currentChar = Mid(normalizedText, i, 1)

        '濁点・半濁点が続く場合の処理
        If i < Len(normalizedText) Then
            nextChar = Mid(normalizedText, i + 1, 1)
        Else
            nextChar = ""
        End If

        Select Case AscW(currentChar)
            '半角カタカナの範囲（ｦ-ﾟ）
            Case &HFF66 To &HFF9F
                '濁点（ﾞ）や半濁点（ﾟ）が続く場合
                If nextChar = ChrW(&HFF9E) Or nextChar = ChrW(&HFF9F) Then
                    '2文字を組み合わせて1つの全角カタカナに変換
                    Select Case currentChar & nextChar
                        Case "ｶﾞ": Mid(normalizedText, i, 2) = "ガ"
                        Case "ｷﾞ": Mid(normalizedText, i, 2) = "ギ"
                        Case "ｸﾞ": Mid(normalizedText, i, 2) = "グ"
                        Case "ｹﾞ": Mid(normalizedText, i, 2) = "ゲ"
                        Case "ｺﾞ": Mid(normalizedText, i, 2) = "ゴ"
                        Case "ｻﾞ": Mid(normalizedText, i, 2) = "ザ"
                        Case "ｼﾞ": Mid(normalizedText, i, 2) = "ジ"
                        Case "ｽﾞ": Mid(normalizedText, i, 2) = "ズ"
                        Case "ｾﾞ": Mid(normalizedText, i, 2) = "ゼ"
                        Case "ｿﾞ": Mid(normalizedText, i, 2) = "ゾ"
                        Case "ﾀﾞ": Mid(normalizedText, i, 2) = "ダ"
                        Case "ﾁﾞ": Mid(normalizedText, i, 2) = "ヂ"
                        Case "ﾂﾞ": Mid(normalizedText, i, 2) = "ヅ"
                        Case "ﾃﾞ": Mid(normalizedText, i, 2) = "デ"
                        Case "ﾄﾞ": Mid(normalizedText, i, 2) = "ド"
                        Case "ﾊﾞ": Mid(normalizedText, i, 2) = "バ"
                        Case "ﾋﾞ": Mid(normalizedText, i, 2) = "ビ"
                        Case "ﾌﾞ": Mid(normalizedText, i, 2) = "ブ"
                        Case "ﾍﾞ": Mid(normalizedText, i, 2) = "ベ"
                        Case "ﾎﾞ": Mid(normalizedText, i, 2) = "ボ"
                        Case "ﾊﾟ": Mid(normalizedText, i, 2) = "パ"
                        Case "ﾋﾟ": Mid(normalizedText, i, 2) = "ピ"
                        Case "ﾌﾟ": Mid(normalizedText, i, 2) = "プ"
                        Case "ﾍﾟ": Mid(normalizedText, i, 2) = "ペ"
                        Case "ﾎﾟ": Mid(normalizedText, i, 2) = "ポ"
                    End Select
                    i = i + 1  '濁点・半濁点分を飛ばす
                Else
                    '濁点・半濁点が続かない場合は1文字ずつ変換
                    Select Case currentChar
                        Case "ｦ": Mid(normalizedText, i, 1) = "ヲ"
                        Case "ｧ": Mid(normalizedText, i, 1) = "ァ"
                        Case "ｨ": Mid(normalizedText, i, 1) = "ィ"
                        Case "ｩ": Mid(normalizedText, i, 1) = "ゥ"
                        Case "ｪ": Mid(normalizedText, i, 1) = "ェ"
                        Case "ｫ": Mid(normalizedText, i, 1) = "ォ"
                        Case "ｬ": Mid(normalizedText, i, 1) = "ャ"
                        Case "ｭ": Mid(normalizedText, i, 1) = "ュ"
                        Case "ｮ": Mid(normalizedText, i, 1) = "ョ"
                        Case "ｯ": Mid(normalizedText, i, 1) = "ッ"
                        Case "ｰ": Mid(normalizedText, i, 1) = "ー"
                        Case "ｱ": Mid(normalizedText, i, 1) = "ア"
                        Case "ｲ": Mid(normalizedText, i, 1) = "イ"
                        Case "ｳ": Mid(normalizedText, i, 1) = "ウ"
                        Case "ｴ": Mid(normalizedText, i, 1) = "エ"
                        Case "ｵ": Mid(normalizedText, i, 1) = "オ"
                        Case "ｶ": Mid(normalizedText, i, 1) = "カ"
                        Case "ｷ": Mid(normalizedText, i, 1) = "キ"
                        Case "ｸ": Mid(normalizedText, i, 1) = "ク"
                        Case "ｹ": Mid(normalizedText, i, 1) = "ケ"
                        Case "ｺ": Mid(normalizedText, i, 1) = "コ"
                        Case "ｻ": Mid(normalizedText, i, 1) = "サ"
                        Case "ｼ": Mid(normalizedText, i, 1) = "シ"
                        Case "ｽ": Mid(normalizedText, i, 1) = "ス"
                        Case "ｾ": Mid(normalizedText, i, 1) = "セ"
                        Case "ｿ": Mid(normalizedText, i, 1) = "ソ"
                        Case "ﾀ": Mid(normalizedText, i, 1) = "タ"
                        Case "ﾁ": Mid(normalizedText, i, 1) = "チ"
                        Case "ﾂ": Mid(normalizedText, i, 1) = "ツ"
                        Case "ﾃ": Mid(normalizedText, i, 1) = "テ"
                        Case "ﾄ": Mid(normalizedText, i, 1) = "ト"
                        Case "ﾅ": Mid(normalizedText, i, 1) = "ナ"
                        Case "ﾆ": Mid(normalizedText, i, 1) = "ニ"
                        Case "ﾇ": Mid(normalizedText, i, 1) = "ヌ"
                        Case "ﾈ": Mid(normalizedText, i, 1) = "ネ"
                        Case "ﾉ": Mid(normalizedText, i, 1) = "ノ"
                        Case "ﾊ": Mid(normalizedText, i, 1) = "ハ"
                        Case "ﾋ": Mid(normalizedText, i, 1) = "ヒ"
                        Case "ﾌ": Mid(normalizedText, i, 1) = "フ"
                        Case "ﾍ": Mid(normalizedText, i, 1) = "ヘ"
                        Case "ﾎ": Mid(normalizedText, i, 1) = "ホ"
                        Case "ﾏ": Mid(normalizedText, i, 1) = "マ"
                        Case "ﾐ": Mid(normalizedText, i, 1) = "ミ"
                        Case "ﾑ": Mid(normalizedText, i, 1) = "ム"
                        Case "ﾒ": Mid(normalizedText, i, 1) = "メ"
                        Case "ﾓ": Mid(normalizedText, i, 1) = "モ"
                        Case "ﾔ": Mid(normalizedText, i, 1) = "ヤ"
                        Case "ﾕ": Mid(normalizedText, i, 1) = "ユ"
                        Case "ﾖ": Mid(normalizedText, i, 1) = "ヨ"
                        Case "ﾗ": Mid(normalizedText, i, 1) = "ラ"
                        Case "ﾘ": Mid(normalizedText, i, 1) = "リ"
                        Case "ﾙ": Mid(normalizedText, i, 1) = "ル"
                        Case "ﾚ": Mid(normalizedText, i, 1) = "レ"
                        Case "ﾛ": Mid(normalizedText, i, 1) = "ロ"
                        Case "ﾜ": Mid(normalizedText, i, 1) = "ワ"
                        Case "ﾝ": Mid(normalizedText, i, 1) = "ン"
                        Case "ﾞ": Mid(normalizedText, i, 1) = ""
                        Case "ﾟ": Mid(normalizedText, i, 1) = ""
                    End Select
                End If
        End Select
    Next i

    '連続する半角スペースを単一の半角スペースに変換
    Do While InStr(normalizedText, "  ") > 0
        normalizedText = Replace(normalizedText, "  ", " ")
    Loop

    NormalizeJapaneseText = normalizedText
End Function
