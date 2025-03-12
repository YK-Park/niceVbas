' text2ppt 250126-01 설문이미지작성툴

' グローバル変数の宣言
Private pptApp As Object
Private pptPres As Object
Private gTextFileName As String
Private gSaveFolderPath As String

Public Sub ImportTextToPPT()
    On Error GoTo ErrorHandler
    
    ' PowerPointアプリケーションの初期化
    InitializePowerPoint
    
    ' フォルダパスの取得と検証
    If Not ValidateFolderPath Then Exit Sub
    
    ' PowerPointファイルの選択と開く
    If Not OpenPowerPointFile Then Exit Sub
    
    ' テキストファイルの選択と処理
    ProcessTextFile
    
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

Private Sub ProcessTextFile()
    ' テキストファイルの選択と処理
    Dim txtPath As String
    Dim fileContent() As String
    
    With Application.FileDialog(msoFileDialogFilePicker)
        .Title = "テキストファイルを選択してください"
        .Filters.Clear
        .Filters.Add "テキストファイル", "*.txt"
        .AllowMultiSelect = False
        
        If .Show Then
            txtPath = .SelectedItems(1)
            gTextFileName = GetFileName(txtPath)
            fileContent = ReadTextFile(txtPath)
            UpdateSlideContent fileContent
        End If
    End With
End Sub

' テキストファイルの読み込み機能
Private Function ReadTextFile(txtPath As String) As String()
   ' ファイルを読み込んで配列として返す
   Dim txtContent As String
   Dim fileNum As Integer
   
   fileNum = FreeFile
   Open txtPath For Input As fileNum
   txtContent = Input$(LOF(fileNum), fileNum)
   Close fileNum
   
   ' 空行を削除して配列に変換
   txtContent = Replace(txtContent, vbCrLf & vbCrLf, vbCrLf)
   ReadTextFile = Split(txtContent, vbCrLf)
End Function

Private Sub UpdateSlideContent(fileContent() As String)
   ' スライド内のテキストボックスを更新
   Dim slide As Object
   Set slide = pptPres.Slides(1)
   
   ' テキストボックスの更新処理
   UpdateTextBoxA slide, fileContent   ' テキストボックスA (4-6行目)
   UpdateTextBoxB slide, fileContent   ' テキストボックスB (10行目)
   UpdateTextBoxC slide, fileContent   ' テキストボックスC (12行目)
   UpdateTextBoxD slide, fileContent   ' テキストボックスD (14行目)
   UpdateTextBoxE slide, fileContent   ' テキストボックスE (16行目)
   
   pptPres.Save
End Sub

Private Sub UpdateTextBoxA(slide As Object, content() As String)
   ' テキストボックスAの更新 (4-6行目)
   Dim combinedText As String
   combinedText = content(3) & vbCrLf & content(4) & vbCrLf & content(5)
   UpdateShapeText slide, "A", combinedText
End Sub

Private Sub UpdateTextBoxB(slide As Object, content() As String)
   ' テキストボックスBの更新 (10行目)
   UpdateShapeText slide, "B", content(9)
End Sub

Private Sub UpdateTextBoxC(slide As Object, content() As String)
   ' テキストボックスCの更新 (12行目)
   UpdateShapeText slide, "C", content(11)
End Sub

Private Sub UpdateTextBoxD(slide As Object, content() As String)
   ' テキストボックスDの更新 (14行目)
   UpdateShapeText slide, "D", content(13)
End Sub

Private Sub UpdateTextBoxE(slide As Object, content() As String)
   ' テキストボックスEの更新 (16行目)
   UpdateShapeText slide, "E", content(15)
End Sub

' テキストボックスの代わりにShapeを使用して更新
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
   ' PDFとして保存する機能
   On Error GoTo ErrorHandler
   
   If pptPres Is Nothing Then
       MsgBox "PowerPointファイルが開かれていません。", vbExclamation
       Exit Sub
   End If
   
   ' PDFファイル名の生成
   Dim pdfPath As String
   If Right(gSaveFolderPath, 1) <> "\" Then
       gSaveFolderPath = gSaveFolderPath & "\"
   End If
   
   pdfPath = gSaveFolderPath & _
             Left(gTextFileName, InStrRev(gTextFileName, ".") - 1) & ".pdf"
             
   ' PDFとして保存
   pptPres.SaveAs pdfPath, ppSaveAsPDF
   
   MsgBox "PDFを保存しました。" & vbCrLf & "保存先: " & pdfPath
   Exit Sub
   
ErrorHandler:
   MsgBox "PDF保存中にエラーが発生しました: " & Err.Description, vbCritical
End Sub

