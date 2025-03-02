첫문장 중앙정렬 서브루틴
Sub CenterFirstSentence(doc As Document, insertPoint As Long, textToInsert As String)
    ' テキスト内の最初の文を中央揃えにする関数
    ' doc: アクティブドキュメント
    ' insertPoint: テキストの挿入位置
    ' textToInsert: 挿入するテキスト
    
    Dim firstLineEnd As Long
    Dim firstLine As String
    Dim remainingText As String
    Dim rng As Range
    
    ' 最初の改行文字を探す
    firstLineEnd = InStr(1, textToInsert, vbCr)
    If firstLineEnd = 0 Then firstLineEnd = InStr(1, textToInsert, vbLf)
    
    If firstLineEnd > 0 Then
        ' 最初の行と残りのテキストを分ける
        firstLine = Left(textToInsert, firstLineEnd - 1)
        remainingText = Mid(textToInsert, firstLineEnd)
        
        ' テキストを挿入する（最初の行と残りのテキストを別々に）
        Set rng = doc.Range(insertPoint, insertPoint)
        rng.InsertAfter firstLine
        rng.ParagraphFormat.Alignment = wdAlignParagraphCenter
        
        ' 挿入位置を更新
        insertPoint = rng.End
        
        ' 残りのテキストを挿入
        Set rng = doc.Range(insertPoint, insertPoint)
        rng.InsertAfter remainingText
    Else
        ' 改行がない場合は全テキストを中央揃えに
        Set rng = doc.Range(insertPoint, insertPoint)
        rng.InsertAfter textToInsert
        rng.ParagraphFormat.Alignment = wdAlignParagraphCenter
    End If
End Sub

Sub InsertFinalText()
    ' 最終テキストをワード文書に挿入し、書式設定する
    
    If Len(insertText) = 0 Then
        MsgBox "挿入するテキストがありません。", vbInformation
        Exit Sub
    End If
    
    ' 例：9番目の改ページ位置を探す
    Dim foundPageBreaks As Integer
    foundPageBreaks = 0
    Dim rng As Range
    Dim insertLocation As Long
    
    Set rng = doc.Range(0, 0)
    
    With rng.Find
        .Text = Chr(12)  ' 改ページコード
        .Forward = True
        .Wrap = wdFindStop
        .Execute
        
        Do While .Found
            foundPageBreaks = foundPageBreaks + 1
            If foundPageBreaks = 9 Then
                Exit Do
            End If
            rng.Collapse wdCollapseEnd
            .Execute
        Loop
    End With
    
    ' 9つ目の改ページの前に中央揃えされた最初の文とその他のテキストを挿入
    If foundPageBreaks = 9 Then
        rng.Collapse Direction:=wdCollapseStart
        insertLocation = rng.Start ' 挿入位置を記録
        
        ' 修正された関数を使用して挿入と書式設定を行う
        CenterFirstSentence doc, insertLocation, insertText
        
        MsgBox "テキストが正常に挿入されました。", vbInformation
    Else
        MsgBox "文書に" & foundPageBreaks & "つの改ページが見つかりました（9つ必要）。", vbExclamation
    End If
End Sub

' ProcessSection1関数を少し修正して一般的な文を取得する例
Sub ProcessSection1()
    ' 第1セクションから特定の文章を抽出する
    Dim textLine As Long
    Dim finalText As String
    
    finalText = ""
    count = 0
    
    ' もし配列が空でなければ処理
    If IsArray(lines1) Then
        If UBound(lines1) >= LBound(lines1) Then
            ' 例：最初の有効な行を抽出（空白行をスキップ）
            For i = LBound(lines1) To UBound(lines1)
                If Trim(lines1(i)) <> "" Then
                    finalText = lines1(i) & vbCrLf
                    Exit For
                End If
            Next i
            
            ' 最初の有効な行の後の2行を追加で取得
            Dim linesAdded As Integer
            linesAdded = 0
            
            For j = i + 1 To UBound(lines1)
                If linesAdded < 2 Then  ' 2行まで追加
                    If Trim(lines1(j)) <> "" Then
                        finalText = finalText & lines1(j) & vbCrLf
                        linesAdded = linesAdded + 1
                    End If
                Else
                    Exit For
                End If
            Next j
        End If
    End If
    
    ' 抽出したテキストをグローバル変数に保存
    If Len(finalText) > 0 Then
        insertText = finalText
    End If
End Sub