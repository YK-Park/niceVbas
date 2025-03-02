Public ResultLines As Variant ' 結果を保存するための配列

Sub ExtractLinesFromWordDocument(startKeyword As String, endKeyword As String)
    ' Wordドキュメントからキーワード間のテキストを抽出して処理
    Dim docText As String
    Dim extractedText As String
    Dim startPos As Long, endPos As Long
    Dim Lines As Variant
    
    ' 現在のWordドキュメントのテキストを取得
    docText = ActiveDocument.Range.Text
    
    ' デバッグ情報
    Debug.Print "ドキュメントの長さ: " & Len(docText) & " 文字"
    Debug.Print "検索キーワード: " & startKeyword & " から " & endKeyword
    
    ' キーワード間のテキストを抽出
    startPos = InStr(1, docText, startKeyword)
    If startPos > 0 Then
        Debug.Print "開始キーワード '" & startKeyword & "' が見つかりました。位置: " & startPos
        
        startPos = startPos + Len(startKeyword)
        endPos = InStr(startPos, docText, endKeyword)
        
        If endPos > 0 Then
            Debug.Print "終了キーワード '" & endKeyword & "' が見つかりました。位置: " & endPos
            extractedText = Mid(docText, startPos, endPos - startPos)
        Else
            Debug.Print "終了キーワード '" & endKeyword & "' が見つかりませんでした。残りのすべてのテキストを抽出します。"
            extractedText = Mid(docText, startPos)
        End If
        
        ' テキストを行ごとに分割
        Lines = Split(extractedText, vbCr)  ' Wordでは通常vbCrが行区切り
        If UBound(Lines) = 0 Then 
            Lines = Split(extractedText, vbCrLf)  ' CRLFの場合
            If UBound(Lines) = 0 Then
                Lines = Split(extractedText, vbLf)  ' LFの場合
            End If
        End If
        
        Debug.Print "抽出されたテキストの行数: " & UBound(Lines) + 1 & " 行"
        
        ' -で始まる行を検索し、2行下の内容を取得
        FindLineStartingWithDashAndGet2LinesBelow Lines
    Else
        Debug.Print "開始キーワード '" & startKeyword & "' が見つかりませんでした。"
        ReDim ResultLines(0 To 0)
        ResultLines(0) = ""
    End If
End Sub

Sub FindLineStartingWithDashAndGet2LinesBelow(Lines As Variant)
    ' -で始まる行を検索し、その2行下の内容を取得する
    Dim i As Long
    Dim foundCount As Long
    Dim tempArray() As String
    
    ' 結果を格納する一時配列を初期化
    ReDim tempArray(0 To 100)  ' 十分な大きさで初期化（必要に応じて調整）
    foundCount = 0
    
    ' -で始まる行を検索
    For i = LBound(Lines) To UBound(Lines) - 2  ' 最後の2行は処理しない（2行下がないため）
        If Left(Trim(Lines(i)), 1) = "-" Then
            Debug.Print "-で始まる行が見つかりました: " & Lines(i)
            
            ' 2行下の内容を取得
            Debug.Print "2行下の内容: " & Lines(i + 2)
            
            ' 結果を配列に追加
            tempArray(foundCount) = Lines(i + 2)
            foundCount = foundCount + 1
        End If
    Next i
    
    ' 実際に見つかった数に配列をリサイズ
    If foundCount > 0 Then
        ReDim Preserve tempArray(0 To foundCount - 1)
        ResultLines = tempArray
        Debug.Print "合計 " & foundCount & " 個の結果が見つかりました"
    Else
        Debug.Print "-で始まる行が見つかりませんでした"
        ReDim ResultLines(0 To 0)
        ResultLines(0) = ""
    End If
End Sub

Sub DisplayResults()
    ' 結果を表示
    Dim i As Long
    
    If Not IsArray(ResultLines) Then
        MsgBox "有効な結果がありません", vbInformation
        Exit Sub
    End If
    
    ' 結果を表示またはWordドキュメントに挿入
    If UBound(ResultLines) >= LBound(ResultLines) Then
        ' 新しい段落を挿入
        Selection.TypeParagraph
        Selection.TypeText "---- 抽出結果 ----"
        Selection.TypeParagraph
        
        ' 結果を挿入
        For i = LBound(ResultLines) To UBound(ResultLines)
            If Len(ResultLines(i)) > 0 Then
                Selection.TypeText ResultLines(i)
                Selection.TypeParagraph
            End If
        Next i
        
        Selection.TypeText "----------------"
        Selection.TypeParagraph
    Else
        MsgBox "抽出結果はありません", vbInformation
    End If
End Sub

Sub RunExtraction()
    ' マクロの実行エントリーポイント
    Dim startKeyword As String
    Dim endKeyword As String
    
    ' キーワードを設定（必要に応じて変更）
    startKeyword = InputBox("開始キーワードを入力してください:", "抽出設定", "START")
    If startKeyword = "" Then Exit Sub ' キャンセルされた場合
    
    endKeyword = InputBox("終了キーワードを入力してください:", "抽出設定", "END")
    If endKeyword = "" Then Exit Sub ' キャンセルされた場合
    
    ' 抽出処理実行
    ExtractLinesFromWordDocument startKeyword, endKeyword
    
    ' 結果を表示するか確認
    If MsgBox("抽出結果をドキュメントに挿入しますか？", vbYesNo + vbQuestion, "確認") = vbYes Then
        DisplayResults
    End If
End Sub