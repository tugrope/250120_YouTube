Sub RemovePrefecture()
    Dim ws As Worksheet
    Dim lastRow As Long
    Dim i As Long
    Dim addressText As String
    Dim prefectures As Variant
    Dim prefecture As Variant
    Dim foundPrefecture As String

    ' 都道府県名のリスト
    prefectures = Array("北海道", "青森県", "岩手県", "宮城県", "秋田県", "山形県", "福島県", _
                        "茨城県", "栃木県", "群馬県", "埼玉県", "千葉県", "東京都", "神奈川県", _
                        "新潟県", "富山県", "石川県", "福井県", "山梨県", "長野県", "岐阜県", _
                        "静岡県", "愛知県", "三重県", "滋賀県", "京都府", "大阪府", "兵庫県", _
                        "奈良県", "和歌山県", "鳥取県", "島根県", "岡山県", "広島県", "山口県", _
                        "徳島県", "香川県", "愛媛県", "高知県", "福岡県", "佐賀県", "長崎県", _
                        "熊本県", "大分県", "宮崎県", "鹿児島県", "沖縄県")

    ' アクティブなワークシートを取得
    Set ws = ActiveSheet

    ' セルC2から最終行を取得
    lastRow = ws.Cells(ws.Rows.Count, "C").End(xlUp).Row

    ' C列のデータを処理
    For i = 2 To lastRow
        addressText = ws.Cells(i, "C").Value
        foundPrefecture = ""

        ' 都道府県名を検索
        For Each prefecture In prefectures
            If InStr(addressText, prefecture) = 1 Then
                foundPrefecture = prefecture
                Exit For
            End If
        Next prefecture

        ' 都道府県名を削除してA列に出力
        If foundPrefecture <> "" Then
            ws.Cells(i, "A").Value = Trim(Replace(addressText, foundPrefecture, ""))
        Else
            ws.Cells(i, "A").Value = addressText
        End If
    Next i

    MsgBox "処理が完了しました。", vbInformation
End Sub