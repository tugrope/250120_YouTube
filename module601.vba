' 関数名: NormalizedZenHaiText

' module06-4.vba　では、全角英数字の後に続く長音記号「ー」をハイフン「-」に変換する処理コードを作れなかったため
' このモジュールに単独コードで作成した。
' このコードをmodule06-3.vbaとマージしてmodule06-4.vbaを作成した。
'
' 機能: 全角英数字の後に続く長音記号「ー」をハイフン「-」に変換する
' 処理フロー:
'   1. 入力セルからテキストを取得
'   2. 結果文字列を初期化
'   3. 文字列を1文字ずつ処理:
'      a. 現在の文字を取得
'      b. 直前の文字が全角英数字かつ現在の文字が長音記号「ー」の場合、ハイフン「-」に変換
'      c. 結果文字列に現在の文字を追加
'      d. 現在の文字を前の文字として保存
'   4. 変換後の文字列を返す

Function NormalizedZenHaiText(targetCell As Range) As String
    Dim inputText As String
    Dim i As Long
    Dim resultText As String
    Dim currentChar As String
    Dim prevChar As String

    ' 入力セルからテキストを取得
    inputText = targetCell.Text

    ' 結果文字列を初期化
    resultText = ""
    prevChar = ""

    ' 文字列を1文字ずつ処理
    For i = 1 To Len(inputText)
        currentChar = Mid(inputText, i, 1)

        ' 全角英数字の次に続く長音記号をハイフンに変換
        If prevChar <> "" Then
            ' 直前の文字が全角英数字かどうかを判定
            If IsFullWidthAlphanumeric(prevChar) Then
                ' 現在の文字が長音記号「ー」かどうかを判定
                If currentChar = ChrW(&H30FC) Then
                    currentChar = "-"
                End If
            End If
        End If

        ' 結果文字列に追加
        resultText = resultText & currentChar

        ' 現在の文字を前の文字として保存
        prevChar = currentChar
    Next i

    NormalizedZenHaiText = resultText
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
