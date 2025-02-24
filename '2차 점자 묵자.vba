'2차 점자 묵자
텍스트 파일에서 일부 텍스트를 취득하여 워드파일에 삽입하는 vba를 만들고 싶다.
그런데 텍스트 파일에서 다음과 같은 형태의 문장이 있다고 하자

-Jane said to James "Thank you for coming" (ジェインはジェームズに「来てくださってありがとうございます」と言った。)

이 문장을 취득하여 다음과 같이 만들어서 워드에 삽입하고 싶다.
●ジェインはジェームズに"Thank you for coming"と言った。

내용은 변하고, 형태가 같은 경우, 이런 형식의 문장을 취득하여 변경한 다음 삽입하기 위해

-로 시작하는 문장을 찾는다.
"" () 「」를 마커로 사용한다
●다음에 ()의 문장을 취득한다음, 「」안의 내용대신 "" 내용을 넣는 것으로 할 수 있지 않을까?


Sub ProcessTextFile()
    ' テキストファイルを処理してWordに挿入するマクロ
    
    Dim FSO As Object
    Dim TextStream As Object
    Dim FilePath As String
    Dim TextLine As String
    Dim ProcessedText As String
    
    ' ファイルシステムオブジェクトを作成
    Set FSO = CreateObject("Scripting.FileSystemObject")
    
    ' テキストファイルのパスを設定
    FilePath = Application.GetOpenFilename("テキストファイル (*.txt),*.txt")
    If FilePath = "False" Then Exit Sub
    
    ' テキストファイルを開く（SJIS形式）
    Set TextStream = FSO.OpenTextFile(FilePath, 1, False, -2)
    
    ' 最初の「-」で始まる行を探す
    Do While Not TextStream.AtEndOfStream
        TextLine = TextStream.ReadLine
        If Left(TextLine, 1) = "-" Then
            ProcessedText = ProcessLine(TextLine)
            Exit Do
        End If
    Loop
    
    ' テキストファイルを閉じる
    TextStream.Close
    
    ' 処理した文章をWordに挿入
    If ProcessedText <> "" Then
        Selection.TypeText "●" & ProcessedText
    End If
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