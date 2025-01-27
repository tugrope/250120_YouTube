Function ConvertToFullWidthKana(targetText As String) As String
    ' 半角カタカナを全角カタカナに変換する関数
    ' 引数：targetText - 変換対象の文字列
    ' 戻り値：変換後の文字列

    Dim halfWidthKana As String
    halfWidthKana = "ｶﾞ,ｷﾞ,ｸﾞ,ｹﾞ,ｺﾞ,ｻﾞ,ｼﾞ,ｽﾞ,ｾﾞ,ｿﾞ,ﾀﾞ,ﾁﾞ,ﾂﾞ,ﾃﾞ,ﾄﾞ,ﾊﾞ,ﾋﾞ,ﾌﾞ,ﾍﾞ,ﾎﾞ,ﾊﾟ,ﾋﾟ,ﾌﾟ,ﾍﾟ,ﾎﾟ" _
                  & ",ｱ,ｲ,ｳ,ｴ,ｵ,ｶ,ｷ,ｸ,ｹ,ｺ,ｻ,ｼ,ｽ,ｾ,ｿ,ﾀ,ﾁ,ﾂ,ﾃ,ﾄ,ﾅ,ﾆ,ﾇ,ﾈ,ﾉ" _
                  & ",ﾊ,ﾋ,ﾌ,ﾍ,ﾎ,ﾏ,ﾐ,ﾑ,ﾒ,ﾓ,ﾔ,ﾕ,ﾖ,ﾗ,ﾘ,ﾙ,ﾚ,ﾛ,ﾜ,ｦ,ﾝ" _
                  & ",ｧ,ｨ,ｩ,ｪ,ｫ,ｬ,ｭ,ｮ,ｯ,ｰ,｡,｢,｣,､,･" ' 半角カタカナリスト

    Dim fullWidthKana As String
    fullWidthKana = "ガ,ギ,グ,ゲ,ゴ,ザ,ジ,ズ,ゼ,ゾ,ダ,ヂ,ヅ,デ,ド,バ,ビ,ブ,ベ,ボ,パ,ピ,プ,ペ,ポ" _
                  & ",ア,イ,ウ,エ,オ,カ,キ,ク,ケ,コ,サ,シ,ス,セ,ソ,タ,チ,ツ,テ,ト,ナ,ニ,ヌ,ネ,ノ" _
                  & ",ハ,ヒ,フ,ヘ,ホ,マ,ミ,ム,メ,モ,ヤ,ユ,ヨ,ラ,リ,ル,レ,ロ,ワ,ヲ,ン" _
                  & ",ァ,ィ,ゥ,ェ,ォ,ャ,ュ,ョ,ッ,ー,。,「,」,、,・" ' 全角カタカナリスト

    Dim halfWidthArray() As String
    Dim fullWidthArray() As String
    halfWidthArray = Split(halfWidthKana, ",")
    fullWidthArray = Split(fullWidthKana, ",")

    Dim i As Integer
    For i = 0 To UBound(halfWidthArray)
        targetText = Replace(targetText, halfWidthArray(i), fullWidthArray(i))
    Next i

    ConvertToFullWidthKana = targetText ' 結果を返す

End Function