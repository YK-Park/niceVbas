' text2ppt 250308-01
' text2ppt 250126-01

' グローバル変数の宣言
Private pptApp As Object
Private pptPres As Object
Private gSaveFolderPath As String
Private gTextFileNames() As String  ' テキストファイル名を保存する配列

Public Sub ImportTextToPPT()
    On Error GoTo ErrorHandler
    
    ' PowerPointアプリケーションの初期化
    InitializePowerPoint
    
    ' フォルダパスの取得と検証
    If Not ValidateFolderPath Then Exit Sub
    
    ' PowerPointファイルの選択と開く
    If Not OpenPowerPointFile Then Exit Sub
    
    ' テキストファイルの選択と処理
    ProcessTextFiles
    
    Exit Sub

ErrorHandler:
    MsgBox "エラーが発生しました: " & Err.Description, vbCritical
    CleanupObjects
End Sub

Private Sub InitializePowerPoint()
    ' PowerPointアプリケーションの初期化と設定
    Set pptApp = CreateObject("PowerPoint.Application")
    pptApp.Visible = True
End Sub

Private Function ValidateFolderPath() As Boolean
    ' 保存先フォルダパスの検証
    gSaveFolderPath = Range("C9").Value
    ValidateFolderPath = (gSaveFolderPath <> "")
    
    If Not ValidateFolderPath Then
        MsgBox "保存先フォルダパスを入力してください。", vbExclamation
    End If
End Function

Private Function OpenPowerPointFile() As Boolean
    ' PowerPointファイルを選択して開く処理
    Dim pptPath As String
    
    With Application.FileDialog(msoFileDialogFilePicker)
        .Title = "PowerPointファイルを選択してください"
        .Filters.Clear
        .Filters.Add "PowerPointファイル", "*.pptm"
        .AllowMultiSelect = False
        
        If .Show Then
            pptPath = .SelectedItems(1)
            Set pptPres = pptApp.Presentations.Open(pptPath)
            OpenPowerPointFile = True
        End If
    End With
End Function

Private Sub ProcessTextFiles()
    ' フォルダを選択して、そのフォルダ内のすべてのテキストファイルを処理する
    Dim folderPath As String
    Dim txtFiles() As String
    
    With Application.FileDialog(msoFileDialogFolderPicker)
        .Title = "テキストファイルが含まれるフォルダを選択してください"
        .AllowMultiSelect = False
        
        If .Show Then
            folderPath = .SelectedItems(1)
            txtFiles = GetTextFilesInFolder(folderPath)
            
            ' 少なくとも1つのテキストファイルがある場合のみ処理
            If UBound(txtFiles) >= LBound(txtFiles) Then
                ProcessMultipleTextFiles txtFiles
            Else
                MsgBox "選択されたフォルダにテキストファイルがありません。", vbExclamation
            End If
        End If
    End With
End Sub

Private Function GetTextFilesInFolder(folderPath As String) As String()
    ' フォルダ内のすべてのテキストファイルを検索
    Dim result() As String
    Dim file As String
    Dim fileCount As Integer
    
    ' 初期化
    fileCount = 0
    ReDim result(0 To 0)
    
    ' フォルダ内のすべての.txtファイルを探す
    file = Dir(folderPath & "\*.txt")
    
    While file <> ""
        ReDim Preserve result(0 To fileCount)
        result(fileCount) = folderPath & "\" & file
        fileCount = fileCount + 1
        file = Dir()
    Wend
    
    GetTextFilesInFolder = result
End Function

Private Sub ProcessMultipleTextFiles(txtFiles() As String)
    Dim i As Integer
    Dim fileContent() As String
    Dim newSlide As Object
    
    ' テキストファイル名の配列を初期化
    ReDim gTextFileNames(LBound(txtFiles) To UBound(txtFiles))
    
    ' 最初のスライド以外のスライドをクリア
    ClearExtraSlides
    
    ' 各テキストファイルを処理
    For i = LBound(txtFiles) To UBound(txtFiles)
        ' ファイル名を抽出して保存
        gTextFileNames(i) = GetFileName(txtFiles(i))
        
        ' ファイル内容を読み込む
        fileContent = ReadTextFile(txtFiles(i))
        
        ' 最初のファイルは既存のスライドを使用
        If i = LBound(txtFiles) Then
            UpdateSlideContent pptPres.Slides(1), fileContent
        Else
            ' スライドを複製する（より安定的な方法）
            pptPres.Slides(1).Duplicate
            
            ' 複製されたスライドは自動的に最後に追加される
            Set newSlide = pptPres.Slides(pptPres.Slides.Count)
            
            ' 新しいスライドの内容を更新
            UpdateSlideContent newSlide, fileContent
        End If
    Next i
    
    ' プレゼンテーションを保存
    pptPres.Save
    
    MsgBox "すべてのテキストファイルが処理されました。合計 " & (UBound(txtFiles) - LBound(txtFiles) + 1) & " 枚のスライドが作成されました。"
End Sub

' テキストファイルの読み込み機能
Private Function ReadTextFile(txtPath As String) As String()
   ' ファイルを読み込んで配列として返す
   Dim txtContent As String
   Dim fileNum As Integer
   Dim cleanedContent As String
   Dim result() As String
   Dim i As Integer, j As Integer
   
   fileNum = FreeFile
   Open txtPath For Input As fileNum
   txtContent = Input$(LOF(fileNum), fileNum)
   Close fileNum
   
   ' 複数の空行を処理するために繰り返し置換を行う
   cleanedContent = txtContent
   Do While InStr(cleanedContent, vbCrLf & vbCrLf) > 0
       cleanedContent = Replace(cleanedContent, vbCrLf & vbCrLf, vbCrLf)
   Loop
   
   ' 空の行を完全に除去する
   Dim tempArray() As String
   tempArray = Split(cleanedContent, vbCrLf)
   
   ' 空でない行数をカウント
   Dim nonEmptyCount As Integer
   nonEmptyCount = 0
   For i = LBound(tempArray) To UBound(tempArray)
       If Trim(tempArray(i)) <> "" Then
           nonEmptyCount = nonEmptyCount + 1
       End If
   Next i
   
   ' 空でない行だけを新しい配列に格納
   ReDim result(0 To nonEmptyCount - 1)
   j = 0
   For i = LBound(tempArray) To UBound(tempArray)
       If Trim(tempArray(i)) <> "" Then
           result(j) = tempArray(i)
           j = j + 1
       End If
   Next i
   
   ReadTextFile = result
End Function

' ファイル名を取得する関数
Private Function GetFileName(fullPath As String) As String
    Dim pos As Integer
    pos = InStrRev(fullPath, "\")
    If pos > 0 Then
        GetFileName = Mid(fullPath, pos + 1)
    Else
        GetFileName = fullPath
    End If
End Function

Private Sub UpdateSlideContent(slide As Object, fileContent() As String)
    ' "Intro"という文字列を含む行から最後までを抽出し、テキストボックスに挿入する
    ' 特定の行 (抽出後の1, 5, 7, 9, 11行目) は太字で表示
    ' 必要な位置に空行を追加する
    
    On Error GoTo ErrorHandler
    
    Dim i As Integer, j As Integer
    Dim introIndex As Integer
    Dim extractedContent() As String
    Dim boldLineNumbers As Variant
    Dim shapeFound As Boolean
    Dim finalText As String
    
    ' 配列のバウンドチェック
    If LBound(fileContent) > UBound(fileContent) Then
        MsgBox "テキストファイルが空です。", vbExclamation
        Exit Sub
    End If
    
    ' "Intro"を含む行を検索
    introIndex = -1
    For i = LBound(fileContent) To UBound(fileContent)
        If InStr(1, fileContent(i), "Intro", vbTextCompare) > 0 Then
            introIndex = i
            Exit For
        End If
    Next i
    
    ' "Intro"が見つからない場合はすべての行を使用
    If introIndex = -1 Then
        introIndex = LBound(fileContent)
    End If
    
    ' Intro行から最後までを抽出 (配列バウンドを安全に処理)
    Dim contentSize As Integer
    contentSize = UBound(fileContent) - introIndex + 1
    
    If contentSize <= 0 Then
        MsgBox "抽出するコンテンツが見つかりません。", vbExclamation
        Exit Sub
    End If
    
    ReDim extractedContent(0 To contentSize - 1)
    j = 0
    For i = introIndex To UBound(fileContent)
        extractedContent(j) = fileContent(i)
        j = j + 1
    Next i
    
    ' Intro行の後に空行を挿入する処理
    finalText = extractedContent(0) & vbCrLf & vbCrLf & vbCrLf
    
    ' 残りの内容を追加
    For i = 1 To UBound(extractedContent)
        finalText = finalText & extractedContent(i) & vbCrLf
    Next i
    
    ' 太字にする行の行番号を設定 (1, 5, 7, 9, 11行目 - 空行を考慮した位置)
    boldLineNumbers = Array(1, 5, 7, 9, 11)
    
    ' メインのテキストボックスを更新（A）
    Dim shp As Object
    
    shapeFound = False
    For Each shp In slide.Shapes
        If shp.AlternativeText = "A" Then
            shapeFound = True
            
            ' テキスト全体を設定
            shp.TextFrame2.TextRange.Text = finalText
            
            ' 指定された行番号を太字に設定
            ApplyBoldToLines shp.TextFrame2.TextRange, boldLineNumbers
            
            Exit For
        End If
    Next shp
    
    If Not shapeFound Then
        MsgBox "代替テキスト「A」が設定されたシェイプが見つかりませんでした。", vbExclamation
    End If
    
    Exit Sub
    
ErrorHandler:
    MsgBox "スライド内容の更新中にエラーが発生しました: " & Err.Description & " (エラー番号: " & Err.Number & ")", vbCritical
End Sub

' 指定された行番号に太字を適用する関数
Private Sub ApplyBoldToLines(textRange As Object, lineNumbers As Variant)
    On Error Resume Next
    
    Dim paragraphs As Object
    Dim lineCount As Integer
    Dim lineNum As Variant
    
    Set paragraphs = textRange.Paragraphs
    lineCount = paragraphs.Count
    
    For Each lineNum In lineNumbers
        ' 行番号が有効範囲内かチェック (1ベースのインデックス)
        If lineNum >= 1 And lineNum <= lineCount Then
            paragraphs(lineNum).Font.Bold = msoTrue
        End If
    Next lineNum
    
    On Error GoTo 0
End Sub

' 値が配列に含まれているかチェックする関数
Private Function IsInArray(valueToFind As Variant, arr As Variant) As Boolean
    Dim element As Variant
    IsInArray = False
    
    For Each element In arr
        If element = valueToFind Then
            IsInArray = True
            Exit Function
        End If
    Next element
End Function

' 特定のShapeのテキストを更新する関数（メインとは別に保持）
Private Sub UpdateShapeText(slide As Object, altText As String, newText As String)
    ' 指定されたShapeのテキストを更新
    Dim shp As Object
    For Each shp In slide.Shapes
        If shp.AlternativeText = altText Then
            shp.TextFrame2.TextRange.Text = newText
            Exit For
        End If
    Next shp
End Sub

Public Sub SaveAsPDF()
   ' 各スライドを個別のPDFとして保存する機能
   On Error GoTo ErrorHandler
   
   If pptPres Is Nothing Then
       MsgBox "PowerPointファイルが開かれていません。", vbExclamation
       Exit Sub
   End If
   
   ' 保存先フォルダパスの確認
   If Right(gSaveFolderPath, 1) <> "\" Then
       gSaveFolderPath = gSaveFolderPath & "\"
   End If
   
   ' 各スライドを個別のPDFとして保存
   Dim i As Integer
   Dim pdfPath As String
   Dim slideFileName As String
   Dim totalSaved As Integer
   
   totalSaved = 0
   For i = 1 To pptPres.Slides.Count
       ' 対応するテキストファイル名があるか確認
       If i <= UBound(gTextFileNames) + 1 Then
           ' テキストファイル名から拡張子を除いた名前を使用
           slideFileName = GetFileNameWithoutExtension(gTextFileNames(i - 1))
           
           pdfPath = gSaveFolderPath & slideFileName & ".pdf"
           
           ' スライドを選択
           pptPres.Slides(i).Select
           
           ' 選択したスライドをPDFとして保存
           pptApp.ActivePresentation.ExportAsFixedFormat _
               pdfPath, _
               ppFixedFormatTypePDF, _
               ppFixedFormatIntentScreen, _
               msoFalse, _
               ppPrintHandoutHorizontalFirst, _
               ppPrintCurrentSlide, _
               msoFalse, _
               Nothing, _
               ppPrintAll, _
               msoFalse
           
           totalSaved = totalSaved + 1
       End If
   Next i
   
   If totalSaved > 0 Then
       MsgBox totalSaved & " 枚のスライドがPDFとして保存されました。" & vbCrLf & "保存先: " & gSaveFolderPath
   Else
       MsgBox "保存されたPDFはありません。テキストファイルが処理されているか確認してください。", vbExclamation
   End If
   
   Exit Sub
   
ErrorHandler:
   MsgBox "PDF保存中にエラーが発生しました: " & Err.Description, vbCritical
End Sub

' ファイル名から拡張子を除いた名前を取得する関数
Private Function GetFileNameWithoutExtension(fullFileName As String) As String
   Dim dotPos As Integer
   dotPos = InStrRev(fullFileName, ".")
   
   If dotPos > 0 Then
       GetFileNameWithoutExtension = Left(fullFileName, dotPos - 1)
   Else
       GetFileNameWithoutExtension = fullFileName
   End If
End Function

' 新しいフォルダを作成する関数
Private Function CreateNewFolder(parentPath As String) As String
   Dim folderName As String
   Dim fullPath As String
   Dim counter As Integer
   
   ' 基本フォルダ名
   folderName = "new"
   fullPath = parentPath & "\" & folderName
   
   ' すでにフォルダが存在する場合は新しい名前を生成
   counter = 1
   While Dir(fullPath, vbDirectory) <> ""
       folderName = "new" & counter
       fullPath = parentPath & "\" & folderName
       counter = counter + 1
   Wend
   
   ' フォルダを作成
   MkDir fullPath
   
   CreateNewFolder = fullPath
End Function

Public Sub ClearExtraSlides()
    ' 最初のスライド以外のすべてのスライドを削除する
    On Error GoTo ErrorHandler
    
    If pptPres Is Nothing Then
        MsgBox "PowerPointファイルが開かれていません。", vbExclamation
        Exit Sub
    End If
    
    ' スライドが1つ以上ある場合のみ処理
    If pptPres.Slides.Count > 0 Then
        ' 最初のスライド以外を削除
        While pptPres.Slides.Count > 1
            pptPres.Slides(pptPres.Slides.Count).Delete
        Wend
        
        MsgBox "最初のスライド以外のすべてのスライドが削除されました。"
    Else
        MsgBox "プレゼンテーションにスライドがありません。", vbExclamation
    End If
    
    Exit Sub
    
ErrorHandler:
    MsgBox "スライド削除中にエラーが発生しました: " & Err.Description, vbCritical
End Sub

Private Sub CleanupObjects()
    ' オブジェクトの解放処理（ただし、閉じないようにする）
    Set pptPres = Nothing
    Set pptApp = Nothing
End Sub