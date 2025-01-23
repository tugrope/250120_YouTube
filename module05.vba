Function NormalizeJapaneseText(targetCell As Range) As String
    Dim inputText As String
    Dim normalizedText As String
    Dim i As Long
    Dim charCode As Integer

    ' セルの内容を取得
    inputText = targetCell.Value

    ' 1. 全角スペースを半角スペースに変換
    normalizedText = Replace(inputText, "　", " ")

    ' 2. 連続する半角スペースを単一の半角スペースに
    Do While InStr(normalizedText, "  ") > 0
        normalizedText = Replace(normalizedText, "  ", " ")
    Loop

    ' 3. 全角英字を半角英字に、全角数字を半角数字に変換
    For i = 1 To Len(normalizedText)
        charCode = AscW(Mid(normalizedText, i, 1))
        ' 全角数字（０～９）
        If charCode >= &HFF10 And charCode <= &HFF19 Then
            Mid(normalizedText, i, 1) = ChrW(charCode - &HFEE0)
        ' 全角英字（A-Z, a-z）
        ElseIf charCode >= &HFF21 And charCode <= &HFF3A Then
            Mid(normalizedText, i, 1) = ChrW(charCode - &HFEE0)
        ElseIf charCode >= &HFF41 And charCode <= &HFF5A Then
            Mid(normalizedText, i, 1) = ChrW(charCode - &HFEE0)
        End If
    Next i

    ' 4. 半角カタカナを全角カタカナに変換
    Dim halfKana As String
    Dim fullKana As String
    Dim pos As Integer

    halfKana = "ｱｲｳｴｵｶｷｸｹｺｻｼｽｾｿﾀﾁﾂﾃﾄﾅﾆﾇﾈﾉﾊﾋﾌﾍﾎﾏﾐﾑﾒﾓﾔﾕﾖﾗﾘﾙﾚﾛﾜｦﾝｯｧｨｩｪｫｬｭｮｰ"
    fullKana = "アイウエオカキクケコサシスセソタチツテトナニヌネノハヒフヘホマミムメモヤユヨラリルレロワヲンッァィゥェォャュョー"

    For i = 1 To Len(normalizedText)
        pos = InStr(halfKana, Mid(normalizedText, i, 1))
        If pos > 0 Then
            Mid(normalizedText, i, 1) = Mid(fullKana, pos, 1)
        End If
    Next i

    ' 5. ハイフンに類似する文字を全角長音記号（U+30FC）に統一
    Dim hyphens As Variant
    hyphens = Array(ChrW(&H002D), ChrW(&H2212), ChrW(&H2010), ChrW(&H2015), ChrW(&H2212), ChrW(&HFF0D))

    For i = LBound(hyphens) To UBound(hyphens)
        normalizedText = Replace(normalizedText, hyphens(i), "ー")
    Next i

    ' 6. 英文字および数字の後に続くU+30FCをU+002D（ハイフンマイナス）に置換
    For i = 1 To Len(normalizedText) - 1
        If ((Mid(normalizedText, i, 1) Like "[A-Za-z0-9]") And Mid(normalizedText, i + 1, 1) = "ー") Then
            Mid(normalizedText, i + 1, 1) = "-"
        End If
    Next i

    ' 正規化されたテキストを返す
    NormalizeJapaneseText = normalizedText
End Function
