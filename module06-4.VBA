' 関数名: NormalizeAndProcessText
'
' 機能: 入力テキストを正規化し、以下の処理を順番に実行する
'   module06-3.vbaとmodule6-44.vbaをマージして作成した。
'
' 処理フロー:
'   1. 入力セルからテキストを取得
'   2. 半角カタカナを全角カタカナに変換
'   3. 半角英数字を全角英数字に変換
'   4. ハイフン類似文字（全角/半角ハイフン、ダッシュなど）を全角長音記号「ー」に統一
'   5. 全角英数字の後に続く長音記号「ー」をハイフン「-」に変換
'
' 注意点:
'   - 各処理は順番に実行され、前の処理結果が次の処理に引き継がれる
'   - エラー発生時はErrorHandlerにジャンプ
'   - 全角英数字の判定にはIsFullWidthAlphanumeric関数を使用


Option Explicit

' メイン関数: NormalizeAndProcessText
Function NormalizeAndProcessText(targetCell As Range) As String
    On Error GoTo ErrorHandler

    Dim inputText As String
    Dim i As Long
    Dim intermediateText As String
    Dim finalText As String
    Dim currentChar As String
    Dim prevChar As String
    Dim tempResult As String

    ' 入力セルからテキストを取得
    inputText = targetCell.Text

    ' 半角カタカナを全角カタカナに変換
    intermediateText = ConvertToFullWidthKana(inputText)

    ' 一時変数を初期化
    tempResult = ""
    prevChar = ""

    ' 半角英数字の全角変換およびハイフン類似文字の処理
    For i = 1 To Len(intermediateText)
        currentChar = Mid(intermediateText, i, 1)

        Select Case AscW(currentChar)
            Case 48 To 57, 65 To 90, 97 To 122 ' 半角数字および英字
                tempResult = tempResult & StrConv(currentChar, vbWide)
            Case &H2D, &H2212, &HFF0D, &H2014, &H2015 ' ハイフン類似文字
                tempResult = tempResult & ChrW(&H30FC) ' 全角長音記号に変換
            Case Else ' その他の文字はそのまま
                tempResult = tempResult & currentChar
        End Select

        ' 現在の文字を前の文字として保存
        prevChar = Right(tempResult, 1)
    Next i

    ' 全角英数字後の長音記号をハイフンに変換
    finalText = ""
    prevChar = ""

    For i = 1 To Len(tempResult)
        currentChar = Mid(tempResult, i, 1)

        If prevChar <> "" Then
            If IsFullWidthAlphanumeric(prevChar) Then
                If currentChar = ChrW(&H30FC) Then
                    currentChar = "-"
                End If
            End If
        End If

        finalText = finalText & currentChar
        prevChar = currentChar
    Next i

    NormalizeAndProcessText = finalText
    Exit Function

ErrorHandler:
    MsgBox "エラーが発生しました: " & Err.Description, vbExclamation
    NormalizeAndProcessText = ""
End Function

' 補助関数: 文字が全角英数字かどうかを判定
Function IsFullWidthAlphanumeric(character As String) As Boolean
    Select Case character
        Case "０" To "９", "Ａ" To "Ｚ", "ａ" To "ｚ"
            IsFullWidthAlphanumeric = True
        Case Else
            IsFullWidthAlphanumeric = False
    End Select
End Function

' 補助関数: 半角カタカナを全角カタカナに変換
Private Function ConvertToFullWidthKana(targetText As String) As String
    ' 半角カタカナを全角カタカナに変換する関数
    ' 引数：targetText - 変換対象の文字列
    ' 戻り値：変換後の文字列

    Dim halfWidthKana As String
    halfWidthKana = "ｶﾞ,ｷﾞ,ｸﾞ,ｹﾞ,ｺﾞ,ｻﾞ,ｼﾞ,ｽﾞ,ｾﾞ,ｿﾞ,ﾀﾞ,ﾁﾞ,ﾂﾞ,ﾃﾞ,ﾄﾞ,ﾊﾞ,ﾋﾞ,ﾌﾞ,ﾍﾞ,ﾎﾞ,ﾊﾟ,ﾋﾟ,ﾌﾟ,ﾍﾟ,ﾎﾟ" _
                  & ",ｱ,ｲ,ｳ,ｴ,ｵ,ｶ,ｷ,ｸ,ｹ,ｺ,ｻ,ｼ,ｽ,ｾ,ｿ,ﾀ,ﾁ,ﾂ,ﾃ,ﾄ,ﾅ,ﾆ,ﾇ,ﾈ,ﾉ" _
                  & ",ﾊ,ﾋ,ﾌ,ﾍ,ﾎ,ﾏ,ﾐ,ﾑ,ﾒ,ﾓ,ﾔ,ﾕ,ﾖ,ﾗ,ﾘ,ﾙ,ﾚ,ﾛ,ﾜ,ｦ,ﾝ" _
                  & ",ｧ,ｨ,ｩ,ｪ,ｫ,ｬ,ｭ,ｮ,ｯ,ｰ,｡,｢,｣,､,･"

    Dim fullWidthKana As String
    fullWidthKana = "ガ,ギ,グ,ゲ,ゴ,ザ,ジ,ズ,ゼ,ゾ,ダ,ヂ,ヅ,デ,ド,バ,ビ,ブ,ベ,ボ,パ,ピ,プ,ペ,ポ" _
                  & ",ア,イ,ウ,エ,オ,カ,キ,ク,ケ,コ,サ,シ,ス,セ,ソ,タ,チ,ツ,テ,ト,ナ,ニ,ヌ,ネ,ノ" _
                  & ",ハ,ヒ,フ,ヘ,ホ,マ,ミ,ム,メ,モ,ヤ,ユ,ヨ,ラ,リ,ル,レ,ロ,ワ,ヲ,ン" _
                  & ",ァ,ィ,ゥ,ェ,ォ,ャ,ュ,ョ,ッ,ー,。,「,」,、,・"

    Dim halfWidthArray() As String
    Dim fullWidthArray() As String
    halfWidthArray = Split(halfWidthKana, ",")
    fullWidthArray = Split(fullWidthKana, ",")

    Dim i As Integer
    For i = 0 To UBound(halfWidthArray)
        targetText = Replace(targetText, halfWidthArray(i), fullWidthArray(i))
    Next i

    ConvertToFullWidthKana = targetText
End Function
