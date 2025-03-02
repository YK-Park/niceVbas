'2차 점자 묵자


내용은 변하고, 형태가 같은 경우, 이런 형식의 문장을 취득하여 변경한 다음 삽입하기 위해

-로 시작하는 문장을 찾는다.
"" () 「」를 마커로 사용한다
●다음에 ()의 문장을 취득한다음, 「」안의 내용대신 "" 내용을 넣는 것으로 할 수 있지 않을까?


2개의 텍스트 파일에서 각각 부분적인 텍스트를 취득하여 4개의 워드파일에 삽입하는 vba를 만들고 싶다.
먼저 엑셀파일에서
1.워드파일폴더와 텍스트파일들을 선택한다.
2.텍스트파일의 내용 중에 조건에 맞는 부분을 키워드로 넓게 취득한다.

워드 1에 삽입할 내용은 첫번째 텍스트파일에서 1부터 5까지의 숫자로 시작하는 문장들이다.
워드 2에 삽입할 내용은 두번째 텍스트파일에서 키워드1부터 키워드2까지의 내용 중에 있다.
워드 3에 삽입할 내용은 두번째 텍스트파일에서 키워드3부터 키워드4까지의 내용 중에 있다.
워드 4에 삽입할 내용은 두번째 텍스트파일에서 키워드5부터 키워드6까지의 내용 중에 있다. 

워드 1에서는 엑셀에서 취득한 내용을 페이지나눔기호 앞에 삽입하고 서식을 설정한다.
워드 2에서는 텍스트 파일에서 키워드1부터 키워드2까지의 내용을 건네받은 내용 중에 다음과 같은 형태의 문장이 있다

-Jane said to James "Thank you for coming" (ジェインはジェームズに「来てくださってありがとうございます」と言った。)

이 문장을 취득하여 다음과 같이 만들어서 페이지나눔기호 앞에 삽입한다.
●ジェインはジェームズに"Thank you for coming"と言った。

기존 워드 파일을 열어서 작업한다 
텍스트 파일에서 "-"로 시작하는 문장은 여러개 있는데 그 중 첫 번째 라인만 처리하면 된다

이런 식으로 워드 3과 워드 4도 처리하고 싶다.



아래 내용을 적절히 활용하여 코드를 완성시켜줘

Sub ProcessTextFile()
    ' テキストファイルを処理してWordに挿入するマクロ
    
    Dim FSO As Object
    Dim TextStream As Object
    Dim WordApp As Object
    Dim WordDoc As Object
    Dim TextFilePath As String
    Dim WordFilePath As String
    Dim TextLine As String
    Dim ProcessedText As String
    
    On Error GoTo ErrorHandler
    
    ' ワードアプリケーションオブジェクトを作成
    Set WordApp = CreateObject("Word.Application")
    
    ' ワードファイルを選択
    WordFilePath = Application.GetOpenFilename("Word ファイル (*.docx),*.docx")
    If WordFilePath = "False" Then
        MsgBox "ワードファイルが選択されていません。"
        Exit Sub
    End If
    
    ' テキストファイルを選択
    TextFilePath = Application.GetOpenFilename("テキストファイル (*.txt),*.txt")
    If TextFilePath = "False" Then
        MsgBox "テキストファイルが選択されていません。"
        Exit Sub
    End If
    
    ' ファイルシステムオブジェクトを作成
    Set FSO = CreateObject("Scripting.FileSystemObject")
    
    ' テキストファイルを開く（SJIS形式）
    Set TextStream = FSO.OpenTextFile(TextFilePath, 1, False, -2)
    
    ' ワードファイルを開く
    Set WordDoc = WordApp.Documents.Open(WordFilePath)
    WordApp.Visible = True
    
    ' 最初の「-」で始まる行を探して処理
    Do While Not TextStream.AtEndOfStream
        TextLine = TextStream.ReadLine
        If Left(TextLine, 1) = "-" Then
            ProcessedText = ProcessLine(TextLine)
            Exit Do
        End If
    Loop
    
    ' テキストファイルを閉じる
    TextStream.Close
    
    ' 処理した文章をワードに挿入
    If ProcessedText <> "" Then
        ' ページ区切りを検索
        With WordDoc.Range.Find
            .Text = "^m"  ' ページ区切りの検索
            .Forward = True
            .Execute
            
            If .Found Then
                ' ページ区切りの前に文章を挿入
                WordDoc.Range(.Range.Start, .Range.Start).InsertBefore "●" & ProcessedText & vbCrLf
            Else
                ' ページ区切りが見つからない場合は文書の最後に追加
                WordDoc.Range(WordDoc.Content.End - 1).InsertAfter "●" & ProcessedText & vbCrLf
            End If
        End With
    End If
    
    ' 変更を保存
    WordDoc.Save
    
CleanUp:
    ' オブジェクトの解放
    If Not TextStream Is Nothing Then TextStream.Close
    If Not WordDoc Is Nothing Then WordDoc.Close
    If Not WordApp Is Nothing Then WordApp.Quit
    Set TextStream = Nothing
    Set WordDoc = Nothing
    Set WordApp = Nothing
    Set FSO = Nothing
    Exit Sub
    
ErrorHandler:
    MsgBox "エラーが発生しました: " & Err.Description
    Resume CleanUp
End Sub

Function ProcessLine(ByVal TextLine As String) As String
    ' 文章を処理する関数
    
    Dim EnglishText As String
    Dim JapaneseText As String
    Dim StartQuote As Long
    Dim EndQuote As Long
    Dim StartParen As Long
    Dim EndParen As Long
    Dim StartJapQuote As Long
    Dim EndJapQuote As Long
    
    ' 英語の引用部分を取得
    StartQuote = InStr(TextLine, """")
    EndQuote = InStrRev(TextLine, """")
    If StartQuote > 0 And EndQuote > StartQuote Then
        EnglishText = Mid(TextLine, StartQuote, EndQuote - StartQuote + 1)
    End If
    
    ' 日本語の部分を取得
    StartParen = InStr(TextLine, "(")
    EndParen = InStr(TextLine, ")")
    If StartParen > 0 And EndParen > StartParen Then
        JapaneseText = Mid(TextLine, StartParen + 1, EndParen - StartParen - 1)
        
        ' 「」の中身を""の中身に置き換える
        StartJapQuote = InStr(JapaneseText, "「")
        EndJapQuote = InStr(JapaneseText, "」")
        If StartJapQuote > 0 And EndJapQuote > StartJapQuote Then
            ProcessLine = Left(JapaneseText, StartJapQuote - 1) & _
                         EnglishText & _
                         Mid(JapaneseText, EndJapQuote + 1)
        End If
    End If
End Function