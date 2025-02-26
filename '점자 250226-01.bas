'점자 excel250226-01
Option Explicit

' ファイル処理のための変数宣言
Dim FSO As Object

' キーワード定義 - 実際に使用するキーワードに変更してください
Const Keyword1 As String = "키워드1"
Const Keyword2 As String = "키워드2"
Const Keyword3 As String = "키워드3"
Const Keyword4 As String = "키워드4"
Const Keyword5 As String = "키워드5"
Const Keyword6 As String = "키워드6"

Sub ProcessFolderContents()
    ' 変数宣言
    Dim FolderPath As String
    Dim TextFiles As Collection
    Dim TextFile1Path As String, TextFile2Path As String
    Dim TextContent1 As String, TextContent2 As String
    Dim WordApp As Object
    Dim ExtractedText1 As String
    Dim ExtractedText2 As String, ExtractedText3 As String, ExtractedText4 As String
    Dim TextFile1Size As Long, TextFile2Size As Long
    
    ' FSO（FileSystemObject）の初期化
    Set FSO = CreateObject("Scripting.FileSystemObject")
    
    ' すべてのファイルが入っているフォルダを選択
    With Application.FileDialog(msoFileDialogFolderPicker)
        .Title = "ワードとテキストファイルがあるフォルダを選択"
        .AllowMultiSelect = False
        If .Show = False Then Exit Sub
        FolderPath = .SelectedItems(1)
    End With
    
    ' フォルダ内のテキストファイルを収集
    Set TextFiles = New Collection
    CollectTextFiles FolderPath, TextFiles
    
    ' テキストファイルがない場合
    If TextFiles.Count < 2 Then
        MsgBox "フォルダには2つ以上のテキストファイルがありません。処理できません。", vbExclamation
        Exit Sub
    End If
    
    ' ファイルサイズに基づいてテキストファイルを選択
    DetermineTextFilesBySize TextFiles, TextFile1Path, TextFile2Path
    
    ' ファイルのエンコーディングを確認して内容を読み込む
    TextContent1 = ReadTextFile(TextFile1Path)
    TextContent2 = ReadTextFile(TextFile2Path)
    
    ' 処理1：1から5までの数字で始まる文章を抽出
    ExtractedText1 = ExtractNumberedText(TextContent1)
    
    ' 処理2：キーワード1からキーワード2までの内容を抽出
    ExtractedText2 = ExtractTextBetweenKeywords(TextContent2, Keyword1, Keyword2)
    
    ' 処理3：キーワード3からキーワード4までの内容を抽出
    ExtractedText3 = ExtractTextBetweenKeywords(TextContent2, Keyword3, Keyword4)
    
    ' 処理4：キーワード5からキーワード6までの内容を抽出
    ExtractedText4 = ExtractTextBetweenKeywords(TextContent2, Keyword5, Keyword6)
    
    ' ワードファイル名を指定（必要に応じて変更）
    Dim WordFileName1 As String, WordFileName2 As String, WordFileName3 As String, WordFileName4 As String
    WordFileName1 = "Word1.docx"
    WordFileName2 = "Word2.docx"
    WordFileName3 = "Word3.docx"
    WordFileName4 = "Word4.docx"
    
    ' ワードを起動
    Set WordApp = CreateObject("Word.Application")
    WordApp.Visible = True
    
    ' ワードマクロを呼び出す
    On Error Resume Next
    WordApp.Run "ProcessExtractedTextWithFolder", FolderPath, WordFileName1, WordFileName2, WordFileName3, WordFileName4, ExtractedText1, ExtractedText2, ExtractedText3, ExtractedText4
    
    If Err.Number <> 0 Then
        MsgBox "ワードでマクロの実行中にエラーが発生しました: " & Err.Description & vbCrLf & _
               "ワードに必要なマクロがインストールされているか確認してください。", vbExclamation
    End If
    On Error GoTo 0
    
    ' オブジェクトの解放
    Set FSO = Nothing
    
    MsgBox "テキスト抽出とワード転送タスクが完了しました。" & vbCrLf & _
           "テキストファイル1: " & FSO.GetFileName(TextFile1Path) & vbCrLf & _
           "テキストファイル2: " & FSO.GetFileName(TextFile2Path), vbInformation
End Sub

Sub CollectTextFiles(folderPath As String, ByRef textFiles As Collection)
    ' フォルダ内のテキストファイルを収集する
    Dim Folder As Object
    Dim File As Object
    
    Set Folder = FSO.GetFolder(folderPath)
    
    For Each File In Folder.Files
        If LCase(FSO.GetExtensionName(File.Name)) = "txt" Then
            textFiles.Add File.Path
        End If
    Next File
End Sub

Sub DetermineTextFilesBySize(textFiles As Collection, ByRef textFile1Path As String, ByRef textFile2Path As String)
    ' ファイルサイズに基づいてテキストファイルを選択する
    Dim i As Integer
    Dim FileSizes As Collection
    Dim TempPath As String
    Dim TempSize As Long
    Dim SortedFiles As Collection
    
    ' ファイルサイズを収集
    Set FileSizes = New Collection
    For i = 1 To textFiles.Count
        TempPath = textFiles(i)
        TempSize = FSO.GetFile(TempPath).Size
        FileSizes.Add TempSize
    Next i
    
    ' 小さいファイルと大きいファイルを見つける
    Dim SmallestSize As Long, LargestSize As Long
    Dim SmallestIndex As Integer, LargestIndex As Integer
    
    SmallestSize = 9223372036854775807# ' Long型の最大値
    LargestSize = 0
    
    For i = 1 To FileSizes.Count
        If FileSizes(i) < SmallestSize Then
            SmallestSize = FileSizes(i)
            SmallestIndex = i
        End If
        
        If FileSizes(i) > LargestSize Then
            LargestSize = FileSizes(i)
            LargestIndex = i
        End If
    Next i
    
    ' ファイルパスを設定
    textFile1Path = textFiles(SmallestIndex)
    textFile2Path = textFiles(LargestIndex)
End Sub

Function ReadTextFile(filePath As String) As String
    ' テキストファイルの内容を読み込む関数（エンコーディングを自動判別）
    Dim Content As String
    Dim FileNum As Integer
    Dim ByteArray() As Byte
    Dim UTF8Identifier As Boolean
    
    ' ファイルをバイナリモードで開く
    FileNum = FreeFile
    Open filePath For Binary As #FileNum
    
    ' ファイルサイズに合わせて配列のサイズを設定
    ReDim ByteArray(LOF(FileNum) - 1)
    
    ' ファイルの内容をバイト配列に読み込む
    Get #FileNum, , ByteArray
    Close #FileNum
    
    ' UTF-8のBOMをチェック（EF BB BF）
    If UBound(ByteArray) >= 2 Then
        If ByteArray(0) = 239 And ByteArray(1) = 187 And ByteArray(2) = 191 Then
            UTF8Identifier = True
            ' BOMを除外した配列を作成
            Dim TempArray() As Byte
            ReDim TempArray(UBound(ByteArray) - 3)
            Dim i As Long
            For i = 3 To UBound(ByteArray)
                TempArray(i - 3) = ByteArray(i)
            Next i
            ByteArray = TempArray
        End If
    End If
    
    ' ADODB.Streamを使用してエンコーディングを処理
    Dim Stream As Object
    Set Stream = CreateObject("ADODB.Stream")
    
    Stream.Open
    Stream.Type = 1 ' バイナリ
    Stream.Write ByteArray
    Stream.Position = 0
    Stream.Type = 2 ' テキスト
    
    ' エンコーディングの設定
    If UTF8Identifier Then
        Stream.Charset = "UTF-8"
    Else
        ' SJISかUTF-8かを推測（簡易的な方法）
        If IsSJIS(ByteArray) Then
            Stream.Charset = "Shift-JIS"
        Else
            Stream.Charset = "UTF-8"
        End If
    End If
    
    Content = Stream.ReadText
    Stream.Close
    Set Stream = Nothing
    
    ReadTextFile = Content
End Function

Function IsSJIS(ByteArray() As Byte) As Boolean
    ' 簡易的なSJIS判定（正確ではない場合があります）
    Dim i As Long
    Dim SJIS_Count As Long
    Dim UTF8_Count As Long
    
    For i = 0 To UBound(ByteArray) - 1
        ' SJISの1バイト目の範囲チェック
        If (ByteArray(i) >= 129 And ByteArray(i) <= 159) Or (ByteArray(i) >= 224 And ByteArray(i) <= 239) Then
            If i + 1 <= UBound(ByteArray) Then
                ' SJISの2バイト目の範囲チェック
                If (ByteArray(i + 1) >= 64 And ByteArray(i + 1) <= 126) Or (ByteArray(i + 1) >= 128 And ByteArray(i + 1) <= 252) Then
                    SJIS_Count = SJIS_Count + 1
                End If
            End If
        End If
        
        ' UTF-8の特徴的なパターンをチェック
        If ByteArray(i) >= 224 And ByteArray(i) <= 239 Then  ' 3バイト文字の1バイト目
            If i + 2 <= UBound(ByteArray) Then
                If (ByteArray(i + 1) >= 128 And ByteArray(i + 1) <= 191) And _
                   (ByteArray(i + 2) >= 128 And ByteArray(i + 2) <= 191) Then
                    UTF8_Count = UTF8_Count + 1
                End If
            End If
        End If
    Next i
    
    ' SJISの特徴が多ければSJISと判断
    IsSJIS = (SJIS_Count > UTF8_Count)
End Function

Function ExtractNumberedText(textContent As String) As String
    ' 1から5までの数字で始まる文章を抽出する関数
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

Function ExtractTextBetweenKeywords(textContent As String, startKeyword As String, endKeyword As String) As String
    ' キーワード間のテキストを抽出する関数
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

