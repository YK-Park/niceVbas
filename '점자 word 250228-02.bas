'점자 word 250228-02
Option Explicit

' グローバル変数の宣言

Public lines() As String  ' キーワード間で抽出されたテキスト全体

' テキストファイルへのパスを受け取り、テキストを抽出して処理するマクロ
Sub ProcessWordFile1WithPath(wordFilePath As String, textFilePath As String)
    ' 現在のドキュメントを使用
    Dim doc As Document
    Set doc = ActiveDocument
    
    ' テキストファイルから内容を読み込む
    Dim textContent As String
    textContent = ReadTextFileContent(textFilePath)
    
    ' 1から5までの数字で始まる行を抽出
    Dim extractedText As String
    extractedText = ExtractNumberedText(textContent)
    
    ' テキストを挿入
    InsertBeforePageBreak doc, extractedText
    
    ' 保存
    doc.Save
End Sub

' テキストファイルへのパスとキーワードを受け取り、テキストを抽出して処理するマクロ
Sub ProcessWordFile2WithPath(wordFilePath As String, textFilePath As String, startKeyword As String, endKeyword As String)
    ' 現在のドキュメントを使用
    Dim doc As Document
    Set doc = ActiveDocument
    
    ' テキストファイルから内容を読み込む
    Dim textContent As String
    textContent = ReadTextFileContent(textFilePath)
    
    ' グローバル変数を使用してキーワード間のテキストを抽出（行配列も設定）
    Call ExtractTextBetweenKeywords(textContent, startKeyword, endKeyword)
    
    ' "-"で始まる行を処理して挿入
    Call ProcessAndInsertDashLine(doc)
    
    ' 保存
    doc.Save
End Sub

' テキストファイルから1-5の数字で始まる行を抽出
Function ExtractNumberedText(textContent As String) As String
    Dim Lines As Variant
    Dim Line As Variant
    Dim ResultText As String
    
    ' テキストを行に分割
    Lines = Split(textContent, vbCrLf)
    If UBound(Lines) = 0 Then Lines = Split(textContent, vbLf)  ' CRのみの場合に対応
    
    ' 1から5までの数字で始まる行を抽出
    For Each Line In Lines
        If Len(Line) > 0 Then
            If IsNumeric(Left(Line, 1)) Then
                If CInt(Left(Line, 1)) >= 1 And CInt(Left(Line, 1)) <= 5 Then
                    ResultText = ResultText & Line & vbCrLf
                End If
            End If
        End If
    Next Line
    
    ExtractNumberedText = ResultText
End Function

Public lines() As String  ' 抽出された行を保存するString配列

Sub ReadAndExtractLines(filePath As String, startKeyword As String, endKeyword As String)
    On Error GoTo ErrorHandler
    
    ' 変数の初期化
    Dim fileNum As Integer
    Dim byteData() As Byte
    Dim fileContent As String
    Dim allLines As Variant
    Dim i As Long, startIndex As Long, endIndex As Long
    Dim keywordFound As Boolean
    
    ' テキストファイルをバイナリで開いて内容を読み込む
    fileNum = FreeFile
    Open filePath For Binary As #fileNum
    ReDim byteData(LOF(fileNum) - 1)
    Get #fileNum, , byteData
    Close #fileNum
    
    ' バイナリデータをテキストに変換
    fileContent = StrConv(byteData, vbUnicode)
    
    ' テキストを行ごとに分割
    allLines = Split(fileContent, vbCrLf)
    If UBound(allLines) = 0 Then allLines = Split(fileContent, vbLf)  ' CRのみの場合に対応
    
    Debug.Print "ファイルから" & UBound(allLines) + 1 & "行を読み込みました"
    
    ' キーワードを含む行の範囲を探す
    keywordFound = False
    startIndex = -1
    endIndex = -1
    
    For i = LBound(allLines) To UBound(allLines)
        ' 開始キーワードを検索
        If startIndex = -1 And InStr(allLines(i), startKeyword) > 0 Then
            startIndex = i + 1  ' キーワードの次の行から開始
            Debug.Print "開始キーワード '" & startKeyword & "' が " & i & " 行目で見つかりました"
            keywordFound = True
        ' 終了キーワードを検索（開始キーワードが見つかった後）
        ElseIf startIndex > -1 And InStr(allLines(i), endKeyword) > 0 Then
            endIndex = i - 1  ' キーワードの前の行まで
            Debug.Print "終了キーワード '" & endKeyword & "' が " & i & " 行目で見つかりました"
            Exit For
        End If
    Next i
    
    ' キーワードが見つからなかった場合
    If Not keywordFound Then
        Debug.Print "開始キーワード '" & startKeyword & "' が見つかりませんでした"
        ReDim lines(0 To 0)
        lines(0) = ""
        Exit Sub
    End If
    
    ' 終了キーワードが見つからなかった場合、ファイルの最後まで抽出
    If endIndex = -1 Then
        endIndex = UBound(allLines)
        Debug.Print "終了キーワード '" & endKeyword & "' が見つかりませんでした。ファイルの最後まで抽出します"
    End If
    
    ' 抽出した行の範囲をString配列に設定
    Dim lineCount As Long
    lineCount = endIndex - startIndex + 1
    
    If lineCount > 0 Then
        ReDim lines(0 To lineCount - 1)
        
        For i = 0 To lineCount - 1
            lines(i) = allLines(startIndex + i)
        Next i
        
        Debug.Print "抽出された行数: " & lineCount & " 行"
    Else
        ' キーワード間に行がない場合
        ReDim lines(0 To 0)
        lines(0) = ""
        Debug.Print "キーワード間に行がありませんでした"
    End If
    
    Exit Sub
    
ErrorHandler:
    MsgBox "エラーが発生しました: " & Err.Description
    On Error Resume Next
    If fileNum > 0 Then Close #fileNum
    ReDim lines(0 To 0)
    lines(0) = ""
End Sub

Sub ProcessLines()
    ' 抽出された行を処理する例
    Dim i As Long
    
    For i = LBound(lines) To UBound(lines)
        ' ここでString配列の各行を処理
        ' 例: 特定の文字列を含む行を検索
        If InStr(lines(i), "検索キーワード") > 0 Then
            Debug.Print "一致する行: " & lines(i)
        End If
    Next i
End Sub



' ページ区切りの前にテキストを挿入
Sub InsertBeforePageBreak(doc As Document, textContent As String)
    If Len(textContent) = 0 Then Exit Sub
    
    Dim BreakPoint As Object
    
    ' ページ区切りを検索
    Set BreakPoint = doc.Content.Find
    With BreakPoint
        .Text = "^m"  ' ページ区切り記号
        .Forward = True
        .Execute
    End With
    
    ' ページ区切りが見つかった場合はその前に挿入、見つからない場合は文書の最初に挿入
    If BreakPoint.Found Then
        doc.Range(0, BreakPoint.Start).InsertAfter textContent
    Else
        doc.Range(0, 0).InsertAfter textContent
    End If
End Sub

' "-"で始まる行を処理してワードに挿入（グローバル変数を使用）
Sub ProcessAndInsertDashLine(doc As Document)
    Dim Line As Variant
    Dim FirstDashLine As String
    Dim ProcessedText As String
    
    ' グローバル変数に保存された行配列から"-"で始まる最初の行を検索
    For Each Line In g_ExtractedLines
        If Len(Line) > 0 Then
            If Left(Trim(Line), 1) = "-" Then
                FirstDashLine = Line
                Debug.Print "「-」で始まる行を見つけました: " & FirstDashLine
                Exit For
            End If
        End If
    Next Line
    
    ' 行が見つかった場合、処理
    If Len(FirstDashLine) > 0 Then
        ProcessedText = ProcessDashLine(FirstDashLine)
        Debug.Print "処理後のテキスト: " & ProcessedText
    End If
    
    ' ページ区切りの前に挿入
    If Len(ProcessedText) > 0 Then
        InsertBeforePageBreak doc, ProcessedText
    End If
End Sub

' "-"で始まる行を特定の形式に変換
Function ProcessDashLine(dashLine As String) As String
    Dim FirstDashLine As String
    Dim StartPos As Long, EndPos As Long
    Dim EnglishPart As String, JapanesePart As String
    Dim ProcessedText As String
    
    FirstDashLine = Trim(Mid(dashLine, 2))  ' "-"を削除
    
    ' "と"の間にある英語部分を抽出
    StartPos = InStr(1, FirstDashLine, """")
    If StartPos > 0 Then
        EndPos = InStr(StartPos + 1, FirstDashLine, """")
        If EndPos > 0 Then
            EnglishPart = Mid(FirstDashLine, StartPos + 1, EndPos - StartPos - 1)
        End If
    End If
    
    ' かっこ内にある日本語部分を抽出
    StartPos = InStr(1, FirstDashLine, "(")
    If StartPos > 0 Then
        EndPos = InStr(StartPos + 1, FirstDashLine, ")")
        If EndPos > 0 Then
            JapanesePart = Mid(FirstDashLine, StartPos + 1, EndPos - StartPos - 1)
        End If
    End If
    
    ' 日本語部分と英語部分を「●日本語部分"英語部分"と言った。」の形式に変換
    If Len(JapanesePart) > 0 And Len(EnglishPart) > 0 Then
        ProcessedText = "●" & JapanesePart & """" & EnglishPart & """と言った。"
    End If
    
    ProcessDashLine = ProcessedText
End Function

' 行配列から特定の条件に一致する行のみを返す関数の例
Function FilterLines(condition As String) As String
    Dim Line As Variant
    Dim ResultText As String
    
    ' 条件に一致する行を抽出（例: 特定の文字列を含む行）
    For Each Line In g_ExtractedLines
        If InStr(1, Line, condition) > 0 Then
            ResultText = ResultText & Line & vbCrLf
        End If
    Next Line
    
    FilterLines = ResultText
End Function