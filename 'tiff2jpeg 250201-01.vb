'tiff2jpeg 250201-01
Option Explicit

Sub ConvertTIFFtoJPEG()
    '// WIA オブジェクトの作成
    Dim imgFile As WIA.ImageFile
    Dim imgProcess As WIA.ImageProcess
    Set imgFile = New WIA.ImageFile
    Set imgProcess = New WIA.ImageProcess
    
    '// ファイル選択ダイアログを表示
    Dim fd As FileDialog
    Set fd = Application.FileDialog(msoFileDialogFilePicker)
    
    With fd
        .Title = "TIFF ファイルを選択してください"
        .Filters.Clear
        .Filters.Add "TIFF Files", "*.tiff;*.tif"
        .AllowMultiSelect = True
        
        If .Show = False Then Exit Sub
        
        '// outputフォルダの作成
        Dim firstFile As String
        firstFile = .SelectedItems(1)
        Dim outputFolder As String
        outputFolder = CreateOutputFolder(GetFolderPath(firstFile))
        
        '// 選択された各ファイルを処理
        Dim filePath As Variant
        Dim successCount As Long
        Dim errorCount As Long
        successCount = 0
        errorCount = 0
        
        For Each filePath In .SelectedItems
            If ProcessTIFFFile(CStr(filePath), imgFile, imgProcess, outputFolder) Then
                successCount = successCount + 1
            Else
                errorCount = errorCount + 1
            End If
        Next filePath
        
        '// 処理結果を表示
        MsgBox "処理完了" & vbCrLf & _
               "成功: " & successCount & "件" & vbCrLf & _
               "失敗: " & errorCount & "件" & vbCrLf & _
               "出力フォルダ: " & outputFolder
    End With
End Sub

Private Function ProcessTIFFFile(ByVal filePath As String, ByVal imgFile As WIA.ImageFile, _
                               ByVal imgProcess As WIA.ImageProcess, ByVal outputFolder As String) As Boolean
    On Error GoTo ErrorHandler
    
    '// TIFFファイルを読み込む
    imgFile.LoadFile filePath
    
    '// ビット数を確認
    If imgFile.PixelDepth = 24 Then
        '// 24ビットの場合、8ビットに変換
        imgProcess.Filters.Add imgProcess.FilterInfos("Convert").FilterID
        imgProcess.Filters(1).Properties("FormatID").Value = wiaFormatBMP
        imgProcess.Filters(1).Properties("Quality").Value = 8
        Set imgFile = imgProcess.Apply(imgFile)
    End If
    
    '// 出力ファイル名の設定
    Dim outputPath As String
    outputPath = outputFolder & "\" & GetFileName(filePath)
    outputPath = Replace(outputPath, ".tif", ".jpg")
    outputPath = Replace(outputPath, ".tiff", ".jpg")
    
    '// JPEGとして保存（品質100%）
    imgProcess.Filters.Clear
    imgProcess.Filters.Add imgProcess.FilterInfos("Convert").FilterID
    imgProcess.Filters(1).Properties("FormatID").Value = wiaFormatJPEG
    imgProcess.Filters(1).Properties("Quality").Value = 100
    Set imgFile = imgProcess.Apply(imgFile)
    imgFile.SaveFile outputPath
    
    ProcessTIFFFile = True
    Exit Function

ErrorHandler:
    ProcessTIFFFile = False
End Function

'// フォルダパスを取得
Private Function GetFolderPath(ByVal filePath As String) As String
    GetFolderPath = Left(filePath, InStrRev(filePath, "\") - 1)
End Function

'// ファイル名を取得
Private Function GetFileName(ByVal filePath As String) As String
    GetFileName = Mid(filePath, InStrRev(filePath, "\") + 1)
End Function

'// outputフォルダを作成
Private Function CreateOutputFolder(ByVal basePath As String) As String
    Dim outputFolder As String
    outputFolder = basePath & "\output"
    
    '// フォルダが存在しない場合は作成
    If Dir(outputFolder, vbDirectory) = "" Then
        MkDir outputFolder
    End If
    
    CreateOutputFolder = outputFolder
End Function