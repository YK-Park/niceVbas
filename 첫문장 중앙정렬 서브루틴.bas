첫문장 중앙정렬 서브루틴
Sub ApplyCenterToFirstLine(ByRef textToFormat As String)
    ' テキスト変数の最初の行に中央揃えマークアップを適用する
    ' 引数のテキストを直接変更する
    
    Dim firstLineEnd As Long
    Dim firstLine As String
    Dim remainingText As String
    
    ' 入力が空の場合は処理しない
    If Len(textToFormat) = 0 Then Exit Sub
    
    ' 最初の改行を探す
    firstLineEnd = InStr(1, textToFormat, vbCr)
    If firstLineEnd = 0 Then
        firstLineEnd = InStr(1, textToFormat, vbLf)
    End If
    
    If firstLineEnd > 0 Then
        ' 改行がある場合、最初の行と残りのテキストに分ける
        firstLine = Left(textToFormat, firstLineEnd - 1)
        remainingText = Mid(textToFormat, firstLineEnd)
        
        ' 変数を書式付きテキストで更新
        textToFormat = "<center>" & firstLine & "</center>" & remainingText
    Else
        ' 改行がない場合はすべてのテキストを中央揃えに
        textToFormat = "<center>" & textToFormat & "</center>"
    End If
End Sub

Sub InsertFinalText()
    ' 最終テキストをワード文書に挿入
    
    If Len(insertText) = 0 Then
        MsgBox "挿入するテキストがありません。", vbInformation
        Exit Sub
    End If
    
    ' 9番目の改ページ位置を探す
    Dim foundPageBreaks As Integer
    foundPageBreaks = 0
    Dim rng As Range
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
    
    ' 9つ目の改ページの前にテキストを挿入
    If foundPageBreaks = 9 Then
        rng.Collapse Direction:=wdCollapseStart
        
        ' テキストを挿入
        rng.InsertAfter insertText
        
        ' 挿入された範囲を取得
        Dim insertedRange As Range
        Set insertedRange = doc.Range(rng.Start, rng.Start + Len(insertText))
        
        ' 中央揃えタグを実際の書式に変換し、タグを削除
        ApplyFormattingAndRemoveTags insertedRange
        
        MsgBox "テキストが正常に挿入されました。", vbInformation
    Else
        MsgBox "文書に" & foundPageBreaks & "つの改ページが見つかりました（9つ必要）。", vbExclamation
    End If
End Sub

Sub ApplyFormattingAndRemoveTags(docRange As Range)
    ' 文書内のタグを検索し、書式を適用してからタグを削除する
    
    Dim startTagText As String, endTagText As String
    startTagText = "<center>"
    endTagText = "</center>"
    
    Dim rngFind As Range
    Set rngFind = docRange.Duplicate
    
    ' 開始タグを検索
    With rngFind.Find
        .Text = startTagText
        .Replacement.Text = ""
        .Forward = True
        .Wrap = wdFindStop
        .Format = False
        .MatchCase = True
        .MatchWholeWord = False
        .Execute
    End With
    
    ' タグが見つかった場合
    If rngFind.Find.Found Then
        ' タグの開始位置を記録
        Dim tagStart As Long
        tagStart = rngFind.Start
        
        ' タグを削除
        rngFind.Text = ""
        
        ' 終了タグを検索
        With rngFind.Find
            .Text = endTagText
            .Replacement.Text = ""
            .Forward = True
            .Wrap = wdFindStop
            .Format = False
            .MatchCase = True
            .MatchWholeWord = False
            .Execute
        End With
        
        ' 終了タグが見つかった場合
        If rngFind.Find.Found Then
            ' タグの終了位置を記録
            Dim tagEnd As Long
            tagEnd = rngFind.Start
            
            ' タグを削除
            rngFind.Text = ""
            
            ' タグの間のテキスト範囲を取得して中央揃えに設定
            Dim contentRange As Range
            Set contentRange = doc.Range(tagStart, tagEnd - Len(startTagText))
            contentRange.ParagraphFormat.Alignment = wdAlignParagraphCenter
        End If
    End If
End Sub

' 使用例
Sub ProcessSection1()
    ' 第1セクションから特定の文章を抽出する
    Dim textLine As Long
    Dim finalText As String
    
    finalText = ""
    count = 0
    
    ' テキストを取得するロジック
    ' ...（既存のロジック）
    
    ' 抽出したテキストをグローバル変数に保存
    If Len(finalText) > 0 Then
        ' テキストを取得したら、中央揃えマークアップを適用
        ApplyCenterToFirstLine finalText
        insertText = finalText
    End If
End Sub