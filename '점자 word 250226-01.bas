'점자 word 250226-01

Option Explicit

' エクセルから呼び出されるメインマクロ、または単独で実行可能
Sub ProcessTextFiles()
    Dim WordFolder As String
    Dim TextFolder As String
    Dim TextContent1 As String, TextContent2 As String
    Dim ExtractedText1 As String, ExtractedText2 As String, ExtractedText3 As String, ExtractedText4 As String
    
    ' ワードファイルのあるフォルダを選択
    With Application.FileDialog(msoFileDialogFolderPicker)
        .Title = "ワードファイルがあるフォルダを選択"
        .AllowMultiSelect = False
        If .Show = False Then Exit Sub
        WordFolder = .SelectedItems(1)
    End With
    
    ' テキストファイルのあるフォルダを選択
    With Application.FileDialog(msoFileDialogFolderPicker)
        .Title = "テキストファイルがあるフォルダを選択"
        .AllowMultiSelect = False
        If .Show = False Then Exit Sub
        TextFolder = .SelectedItems(1)
    End With
    
    ' テキストファイルを選択
    Dim TextFile1Path As String, TextFile2Path As String
    
    With Application.FileDialog(msoFileDialogFilePicker)
        .Title = "最初のテキストファイルを選択"
        .InitialFileName = TextFolder & "\"
        .AllowMultiSelect = False
        .Filters.Add "テキストファイル", "*.txt", 1
        If .Show = False Then Exit Sub
        TextFile1Path = .SelectedItems(1)
    End With
    
    With Application.FileDialog(msoFileDialogFilePicker)
        .Title = "2番目のテキストファイルを選択"
        .InitialFileName = TextFolder & "\"
        .AllowMultiSelect = False
        .Filters.Add "テキストファイル", "*.txt", 1
        If .Show = False Then Exit Sub
        TextFile2Path = .SelectedItems(1)
    End With
    
    ' テキストファイルの内容を読み込む
    TextContent1 = ReadTextFile(TextFile1Path)
    TextContent2 = ReadTextFile(TextFile2Path)
    
    ' 内容を抽出
    ExtractedText1 = ExtractNumberedText(TextContent1)
    ExtractedText2 = ExtractTextBetweenKeywords(TextContent2, "키워드1", "키워드2")
    ExtractedText3 = ExtractTextBetweenKeywords(TextContent2, "키워드3", "키워드4")
    ExtractedText4 = ExtractTextBetweenKeywords(TextContent2, "키워드5", "키워드6")
    
    ' ワードフォルダ内のファイルを処理
    ProcessWordFolder WordFolder, ExtractedText1, ExtractedText2, ExtractedText3, ExtractedText4
    
    MsgBox "すべてのファイルが処理されました。", vbInformation
End Sub

' エクセルから渡された内容を使ってワードフォルダを処理
Sub ProcessWordFolder(WordFolder As String, ExtractedText1 As String, ExtractedText2 As String, ExtractedText3 As String, ExtractedText4 As String)
    Dim WordFile1 As String, WordFile2 As String, WordFile3 As String, WordFile4 As String
    Dim AllFiles As Object
    Dim File As Object
    Dim FSO As Object
    
    ' FileSystemObjectの作成
    Set FSO = CreateObject("Scripting.FileSystemObject")
    
    ' フォルダ内のすべてのファイルを取得
    Set AllFiles = FSO.GetFolder(WordFolder).Files
    
    ' ワードファイル名を設定（実際のファイル名に合わせて変更）
    WordFile1 = "Word1.docx"
    WordFile2 = "Word2.docx"
    WordFile3 = "Word3.docx"
    WordFile4 = "Word4.docx"
    
    ' 各ファイルを処理
    For Each File In AllFiles
        If FSO.GetFileName(File) = WordFile1 Then
            ProcessWordFile1 File.Path, ExtractedText1
        ElseIf FSO.GetFileName(File) = WordFile2 Then
            ProcessWordFile2 File.Path, ExtractedText2
        ElseIf FSO.GetFileName(File) = WordFile3 Then
            ProcessWordFile2 File.Path, ExtractedText3  ' ProcessWordFile2と同じ処理を使用
        ElseIf FSO.GetFileName(File) = WordFile4 Then
            ProcessWordFile2 File.Path, ExtractedText4  ' ProcessWordFile2と同じ処理を使用
        End If
    Next File
    
    ' オブジェクトの解放
    Set AllFiles = Nothing
    Set FSO = Nothing
End Sub

' エクセルから呼び出されるマクロ（互換性のため残す）
Sub ProcessExtractedTextWithFolder(WordFolder As String, WordFileName1 As String, WordFileName2 As String, WordFileName3 As String, WordFileName4 As String, ExtractedText1 As String, ExtractedText2 As String, ExtractedText3 As String, ExtractedText4 As String)
    ' 各ファイルのパスを作成して処理
    Dim WordFile1Path As String, WordFile2Path As String, WordFile3Path As String, WordFile4Path As String
    
    WordFile1Path = WordFolder & "\" & WordFileName1
    WordFile2Path = WordFolder & "\" & WordFileName2
    WordFile3Path = WordFolder & "\" & WordFileName3
    WordFile4Path = WordFolder & "\" & WordFileName4
    
    ' 各ファイルを処理
    ProcessWordFile1 WordFile1Path, ExtractedText1
    ProcessWordFile2 WordFile2Path, ExtractedText2
    ProcessWordFile2 WordFile3Path, ExtractedText3  ' ProcessWordFile2と同じ処理を使用
    ProcessWordFile2 WordFile4Path, ExtractedText4  ' ProcessWordFile2と同じ処理を使用
    
    MsgBox "すべてのワードファイルが処理されました。", vbInformation
End Sub

' Word1用の処理関数
Sub ProcessWordFile1(filePath As String, textContent As String)
    ' ファイルを開いて処理
    On Error Resume Next
    Dim doc As Document
    Set doc = Documents.Open(filePath)
    
    If Err.Number <> 0 Then
        MsgBox "ファイルを開けません: " & filePath & vbCrLf & "エラー: " & Err.Description, vbExclamation
        Exit Sub
    End If
    On Error GoTo 0
    
    ' テキストを挿入
    InsertBeforePageBreak doc, textContent
    
    ' 保存して閉じる
    doc.Save
    doc.Close
End Sub

' Word2, Word3, Word4用の処理関数
Sub ProcessWordFile2(filePath As String, textContent As String)
    ' ファイルを開いて処理
    On Error Resume Next
    Dim doc As Document
    Set doc = Documents.Open(filePath)
    
    If Err.Number <> 0 Then
        MsgBox "ファイルを開けません: " & filePath & vbCrLf & "エラー: " & Err.Description, vbExclamation
        Exit Sub
    End If
    On Error GoTo 0
    
    ' "-"で始まる行を処理して挿入
    ProcessAndInsertDashLine doc, textContent
    
    ' 保存して閉じる
    doc.Save
    doc.Close
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

' "-"で始まる行を処理してワードに挿入
Sub ProcessAndInsertDashLine(doc As Document, textContent As String)
    Dim Lines As Variant
    Dim Line As Variant
    Dim FirstDashLine As String
    Dim ProcessedText As String
    
    ' テキストを行に分割
    Lines = Split(textContent, vbCrLf)
    If UBound(Lines) = 0 Then Lines = Split(textContent, vbLf)  ' CRのみの場合に対応
    
    ' "-"で始まる最初の行を検索
    For Each Line In Lines
        If Len(Line) > 0 Then
            If Left(Trim(Line), 1) = "-" Then
                FirstDashLine = Line
                Exit For
            End If
        End If
    Next Line
    
    ' 行が見つかった場合、処理
    If Len(FirstDashLine) > 0 Then
        ProcessedText = ProcessDashLine(FirstDashLine)
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

' テキストファイルを読み込む関数
Function ReadTextFile(filePath As String) As String
    Dim FSO As Object
    Dim TextStream As Object
    Dim Content As String
    Dim FileBytes() As Byte
    Dim FileNum As Integer
    
    ' まずはバイナリとして読み込んでUTF-8かShift-JISかを判断
    FileNum = FreeFile
    Open filePath For Binary As #FileNum
    ReDim FileBytes(LOF(FileNum) - 1)
    Get #FileNum, , FileBytes
    Close #FileNum
    
    ' FileSystemObjectの作成
    Set FSO = CreateObject("Scripting.FileSystemObject")
    
    ' UTF-8のBOMをチェック
    If UBound(FileBytes) >= 2 Then
        If FileBytes(0) = 239 And FileBytes(1) = 187 And FileBytes(2) = 191 Then
            ' UTF-8 with BOM
            Set TextStream = FSO.OpenTextFile(filePath, 1, False, -1) ' -1 = TristateMixed
        Else
            ' エンコーディングを推測
            If IsUTF8(FileBytes) Then
                ' UTF-8 without BOM（ここではADODBを使用）
                Dim ADOStream As Object
                Set ADOStream = CreateObject("ADODB.Stream")
                ADOStream.Open
                ADOStream.Type = 1 ' Binary
                ADOStream.LoadFromFile filePath
                ADOStream.Position = 0
                ADOStream.Type = 2 ' Text
                ADOStream.Charset = "UTF-8"
                Content = ADOStream.ReadText
                ADOStream.Close
                Set ADOStream = Nothing
                ReadTextFile = Content
                Exit Function
            Else
                ' Shift-JIS想定
                Set TextStream = FSO.OpenTextFile(filePath, 1, False, 0) ' 0 = TristateUseDefault
            End If
        End If
    Else
        ' 小さいファイルはデフォルトエンコーディングで
        Set TextStream = FSO.OpenTextFile(filePath, 1, False, 0)
    End If
    
    ' ファイル内容を読み込む
    Content = TextStream.ReadAll
    TextStream.Close
    
    ' オブジェクトの解放
    Set TextStream = Nothing
    Set FSO = Nothing
    
    ReadTextFile = Content
End Function

' UTF-8かどうかを判定する関数
Function IsUTF8(Bytes() As Byte) As Boolean
    Dim i As Long
    Dim UTF8Count As Long, SJISCount As Long
    Dim ByteCount As Long
    
    i = 0
    While i <= UBound(Bytes)
        ' UTF-8の特徴的なパターンをチェック
        If (Bytes(i) And &H80) = 0 Then
            ' 1バイト文字
            i = i + 1
        ElseIf (Bytes(i) And &HE0) = &HC0 Then
            ' 2バイト文字の先頭
            If i + 1 <= UBound(Bytes) Then
                If (Bytes(i + 1) And &HC0) = &H80 Then
                    UTF8Count = UTF8Count + 1
                End If
            End If
            i = i + 2
        ElseIf (Bytes(i) And &HF0) = &HE0 Then
            ' 3バイト文字の先頭
            If i + 2 <= UBound(Bytes) Then
                If (Bytes(i + 1) And &HC0) = &H80 And (Bytes(i + 2) And &HC0) = &H80 Then
                    UTF8Count = UTF8Count + 1
                End If
            End If
            i = i + 3
        ElseIf (Bytes(i) And &HF8) = &HF0 Then
            ' 4バイト文字の先頭
            If i + 3 <= UBound(Bytes) Then
                If (Bytes(i + 1) And &HC0) = &H80 And (Bytes(i + 2) And &HC0) = &H80 And (Bytes(i + 3) And &HC0) = &H80 Then
                    UTF8Count = UTF8Count + 1
                End If
            End If
            i = i + 4
        Else
            ' SJIS文字の可能性
            If i + 1 <= UBound(Bytes) Then
                If (Bytes(i) >= &H81 And Bytes(i) <= &H9F) Or (Bytes(i) >= &HE0 And Bytes(i) <= &HFC) Then
                    If (Bytes(i + 1) >= &H40 And Bytes(i + 1) <= &HFC) And Bytes(i + 1) <> &H7F Then
                        SJISCount = SJISCount + 1
                    End If
                End If
            End If
            i = i + 1
        End If
        
        ByteCount = ByteCount + 1
        If ByteCount > 1000 Then Exit While ' 最初の1000バイトまでチェック
    Wend
    
    ' UTF-8の特徴が多ければUTF-8と判断
    IsUTF8 = (UTF8Count > SJISCount) And UTF8Count > 0
End Function

' 1から5までの数字で始まる文章を抽出
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

' キーワード間のテキストを抽出
Function ExtractTextBetweenKeywords(textContent As String, startKeyword As String, endKeyword As String) As String
    Dim StartPos As Long, EndPos As Long
    Dim ExtractedText As String
    
    ' キーワード間のテキストを抽出
    StartPos = InStr(1, textContent, startKeyword)
    If StartPos > 0 Then
        StartPos = StartPos + Len(startKeyword)
        EndPos = InStr(StartPos, textContent, endKeyword)
        If EndPos > 0 Then
            ExtractedText = Mid(textContent, StartPos, EndPos - StartPos)
        Else
            ExtractedText = Mid(textContent, StartPos)
        End If
    End If
    
    ExtractTextBetweenKeywords = ExtractedText
End Function