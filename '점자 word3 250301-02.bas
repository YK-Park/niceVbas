'점자 word3 250301-02

Option Explicit

' 키워드 상수 정의 (3쌍)
Const keyword1_1 As String = "A1"
Const keyword1_2 As String = "A2"
Const keyword2_1 As String = "B1"
Const keyword2_2 As String = "B2" 
Const keyword3_1 As String = "C1"
Const keyword3_2 As String = "C2"

' 전역 변수 선언
Public allText As String
Public insertText As String
Public fileNum As Integer
Public byteData() As Byte
Public lines1 As Variant  ' 첫 번째 섹션용 라인 배열
Public lines2 As Variant  ' 두 번째 섹션용 라인 배열
Public lines3 As Variant  ' 세 번째 섹션용 라인 배열
Public allLines As Variant
Public fileContent As String
Public doc As Document
Public startLines As Long
Public endLines As Long
Public textLines As Long
Public i As Long
Public j As Long
Public k As Long
Public count As Long
Public extractedLines() As String
Public extractedText1 As String  ' 첫 번째 섹션 추출 텍스트
Public extractedText2 As String  ' 두 번째 섹션 추출 텍스트
Public extractedText3 As String  ' 세 번째 섹션 추출 텍스트
Public ResultText As String
Public textContent As String

Sub ProcessWordFile2WithPath(wordFilePath As String, textFilePath As String)
    ' 에러 처리 추가
    On Error Resume Next
    
    Set doc = ActiveDocument
    
    ' 텍스트 파일 읽기
    textContent = ReadTextFile(textFilePath)
    
    ' 각 키워드 쌍으로 텍스트 추출
    extractedText1 = ExtractTextBetweenKeywords(textContent, keyword1_1, keyword1_2, lines1)
    extractedText2 = ExtractTextBetweenKeywords(textContent, keyword2_1, keyword2_2, lines2)
    extractedText3 = ExtractTextBetweenKeywords(textContent, keyword3_1, keyword3_2, lines3)
    
    ' 추출된 텍스트에서 특정 문장 처리
    Call ProcessSection1  ' 첫 번째 섹션 처리
    Call ProcessSection2  ' 두 번째 섹션 처리
    Call ProcessSection3  ' 세 번째 섹션 처리
    
    ' 최종 결과 삽입
    Call InsertFinalText
    
    If Err.Number <> 0 And Err.Number <> 0 Then
        MsgBox "エラーが発生しました: " & Err.Description, vbExclamation
    End If
    
    On Error GoTo 0
End Sub

Function ReadTextFile(filePath As String) As String
    ' テキストファイルを読み込む関数
    Dim stream As Object
    Set stream = CreateObject("ADODB.Stream")

    On Error Resume Next
    stream.Charset = "Shift-JIS"  ' 日本語環境ではShift-JISが一般的
    stream.Open
    stream.LoadFromFile filePath
    
    If Err.Number <> 0 Then
        stream.Close
        Set stream = Nothing
        
        ' UTF-8でも試す
        Set stream = CreateObject("ADODB.Stream")
        stream.Charset = "UTF-8"
        stream.Open
        stream.LoadFromFile filePath
    End If

    ReadTextFile = stream.ReadText
    stream.Close
    Set stream = Nothing
    On Error GoTo 0
End Function

Function ExtractTextBetweenKeywords(textContent As String, startKeyword As String, endKeyword As String, ByRef linesArray As Variant) As String
    ' キーワード間のテキストを抽出して行配列も設定する関数
    Dim StartPos As Long, EndPos As Long
    Dim extractedContent As String
    Dim tempLines As Variant
    
    ' デバッグ情報
    Debug.Print "検索キーワード: " & startKeyword & " から " & endKeyword

    ' キーワード間のテキストを抽出
    StartPos = InStr(1, textContent, startKeyword)
    If StartPos > 0 Then
        Debug.Print "開始キーワード '" & startKeyword & "' が見つかりました。位置: " & StartPos

        StartPos = StartPos + Len(startKeyword)
        EndPos = InStr(StartPos, textContent, endKeyword)

        If EndPos > 0 Then
            Debug.Print "終了キーワード '" & endKeyword & "' が見つかりました。位置: " & EndPos
            extractedContent = Mid(textContent, StartPos, EndPos - StartPos)
        Else
            Debug.Print "終了キーワード '" & endKeyword & "' が見つかりませんでした。"
            extractedContent = Mid(textContent, StartPos)
        End If

        ' 抽出したテキストを行に分割
        tempLines = Split(extractedContent, vbCrLf)
        If UBound(tempLines) = 0 Then tempLines = Split(extractedContent, vbLf)
        
        ' 配列をByRefで返す
        linesArray = tempLines
    Else
        Debug.Print "開始キーワード '" & startKeyword & "' が見つかりませんでした。"
        extractedContent = ""
        linesArray = Array()  ' 空の配列
    End If

    ExtractTextBetweenKeywords = extractedContent
End Function

Sub ProcessSection1()
    ' 第1セクションから特定の文章を抽出する
    Dim textLine As Long
    Dim finalText As String
    
    finalText = ""
    count = 0
    
    ' もし配列が空でなければ処理
    If IsArray(lines1) Then
        If UBound(lines1) >= LBound(lines1) Then
            ' 例：「●」マークを含む行を探す
            For i = LBound(lines1) To UBound(lines1)
                If InStr(lines1(i), "●") > 0 Then
                    count = count + 1
                    If count = 1 Then  ' 最初の「●」マークがある行の2行後
                        textLine = i + 2
                        If textLine <= UBound(lines1) Then
                            finalText = finalText & lines1(textLine) & vbCrLf
                        End If
                    End If
                End If
            Next i
        End If
    End If
    
    ' 抽出したテキストをグローバル変数に保存
    If Len(finalText) > 0 Then
        insertText = insertText & "==Section 1==" & vbCrLf & finalText & vbCrLf
    End If
End Sub

Sub ProcessSection2()
    ' 第2セクションから特定の文章を抽出する
    Dim textLine As Long
    Dim finalText As String
    
    finalText = ""
    
    ' もし配列が空でなければ処理
    If IsArray(lines2) Then
        If UBound(lines2) >= LBound(lines2) Then
            ' 例：最初の3行を取得
            For i = LBound(lines2) To LBound(lines2) + 2
                If i <= UBound(lines2) Then
                    finalText = finalText & lines2(i) & vbCrLf
                End If
            Next i
        End If
    End If
    
    ' 抽出したテキストをグローバル変数に追加
    If Len(finalText) > 0 Then
        insertText = insertText & "==Section 2==" & vbCrLf & finalText & vbCrLf
    End If
End Sub

Sub ProcessSection3()
    ' 第3セクションから特定の文章を抽出する
    Dim finalText As String
    Dim keyPhrase As String
    
    finalText = ""
    keyPhrase = "重要"  ' 例：「重要」を含む行を探す
    
    ' もし配列が空でなければ処理
    If IsArray(lines3) Then
        If UBound(lines3) >= LBound(lines3) Then
            For i = LBound(lines3) To UBound(lines3)
                If InStr(lines3(i), keyPhrase) > 0 Then
                    finalText = finalText & lines3(i) & vbCrLf
                    ' その行の次の行も取得
                    If i + 1 <= UBound(lines3) Then
                        finalText = finalText & lines3(i + 1) & vbCrLf
                    End If
                End If
            Next i
        End If
    End If
    
    ' 抽出したテキストをグローバル変数に追加
    If Len(finalText) > 0 Then
        insertText = insertText & "==Section 3==" & vbCrLf & finalText
    End If
End Sub

Sub InsertFinalText()
    ' 最終テキストをワード文書に挿入
    
    If Len(insertText) = 0 Then
        MsgBox "挿入するテキストがありません。", vbInformation
        Exit Sub
    End If
    
    ' 例：9番目の改ページ位置に挿入
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
        rng.InsertBefore insertText
        MsgBox "テキストが正常に挿入されました。", vbInformation
    Else
        MsgBox "文書に" & foundPageBreaks & "つの改ページが見つかりました（9つ必要）。", vbExclamation
    End If
End Sub

Sub DeleteManualPageBreaks()
    ' 手動ページ区切りを削除するサブルーチン
    On Error Resume Next
    
    Dim myRange As Range
    Set myRange = ActiveDocument.Range
    
    Application.ScreenUpdating = False
    Application.StatusBar = "手動で挿入されたページ区切りを削除中..."
    
    With myRange.Find
        .ClearFormatting
        .Text = "^m"
        .Replacement.Text = ""
        .Forward = True
        .Wrap = wdFindContinue
        .Format = False
        .MatchCase = False
        .MatchWholeWord = False
        .MatchWildcards = False
        .MatchSoundsLike = False
        .MatchAllWordForms = False
        .Execute Replace:=wdReplaceAll
    End With
    
    Application.StatusBar = False
    Application.ScreenUpdating = True
    
    If Err.Number = 0 Then
        MsgBox "すべての手動で挿入されたページ区切りが削除されました。", vbInformation
    End If
    
    On Error GoTo 0
End Sub