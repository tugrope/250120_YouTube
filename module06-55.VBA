' 6-5. 全角英数字を半角英数字に変換
' 6-6. 全角スペースを半角スペースに変換
' 6-7. 連続する半角スペースを単一の半角スペースにする

Function NormalizeAndProcessText3(targetCell As Range) As String
    On Error GoTo ErrorHandler

    Dim inputText As String
    Dim intermediateText As String
    Dim finalText As String

    ' 入力セルからテキストを取得
    inputText = targetCell.Text

    ' 6-5. 全角英数字を半角英数字に変換（カタカナを除く）
    ' 6-6. 全角スペースを半角スペースに変換
    intermediateText = ConvertToNarrowExceptKatakana(inputText)

    ' 6-7. 連続する半角スペースを単一の半角スペースにする
    finalText = ReplaceMultipleSpaces2(intermediateText)

    ' 結果を返す
    NormalizeAndProcessText3 = finalText
    Exit Function

ErrorHandler:
    MsgBox "テキストの正規化中にエラーが発生しました: " & Err.Description, vbExclamation
    NormalizeAndProcessText3 = inputText
End Function

' 補助関数: 連続する半角スペースを単一の半角スペースに置換
Private Function ReplaceMultipleSpaces2(text As String) As String
    Do While InStr(text, "  ") > 0
        text = Replace(text, "  ", " ")
    Loop
    ReplaceMultipleSpaces2 = text
End Function

' 補助関数: 英数字とスペースのみを半角に変換
Private Function ConvertToNarrowExceptKatakana(text As String) As String
    Dim i As Long
    Dim result As String
    Dim currentChar As String

    For i = 1 To Len(text)
        currentChar = Mid(text, i, 1)
        ' 全角の英数字またはスペースかどうかをチェック
        If IsWideAlphaNumericOrSpace(currentChar) Then
            ' 英数字とスペースのみ半角に変換
            result = result & StrConv(currentChar, vbNarrow)
        Else
            ' その他の文字（カタカナ含む）はそのまま
            result = result & currentChar
        End If
    Next i

    ConvertToNarrowExceptKatakana = result
End Function

' 全角の英数字またはスペースかどうかを判定
Private Function IsWideAlphaNumericOrSpace(char As String) As Boolean
    Dim charCode As Long
    charCode = AscW(char)

    ' 全角スペース
    If charCode = &H3000 Then
        IsWideAlphaNumericOrSpace = True
        Exit Function
    End If

    ' 全角英数字（０-９、Ａ-Ｚ、ａ-ｚ）
    If (charCode >= &HFF10 And charCode <= &HFF19) Or _  '全角数字
       (charCode >= &HFF21 And charCode <= &HFF3A) Or _  '全角英大文字
       (charCode >= &HFF41 And charCode <= &HFF5A) Then  '全角英小文字
        IsWideAlphaNumericOrSpace = True
        Exit Function
    End If

    IsWideAlphaNumericOrSpace = False
End Function