'tiff2jpeg 250215-01
Sub ConvertTiffToJpeg()
    Dim folderPath As String
    Dim fileName As String
    Dim shellApp As Object
    Dim totalFiles As Long
    Dim processedFiles As Long
    Dim errorCount As Long
    
    'フォルダー選択ダイアログを表示
    With Application.FileDialog(msoFileDialogFolderPicker)
        .Title = "TIFF ファイルが存在するフォルダーを選択してください"
        If .Show = -1 Then
            folderPath = .SelectedItems(1)
        Else
            MsgBox "フォルダーが選択されていません。", vbExclamation
            Exit Sub
        End If
    End With
    
    '変数の初期化
    Set shellApp = CreateObject("Shell.Application")
    totalFiles = 0
    processedFiles = 0
    errorCount = 0
    
    'TIFFファイルの数をカウント
    fileName = Dir(folderPath & "\*.tif*")
    Do While fileName <> ""
        totalFiles = totalFiles + 1
        fileName = Dir()
    Loop
    
    If totalFiles = 0 Then
        MsgBox "指定されたフォルダーにTIFFファイルが見つかりません。", vbInformation
        Exit Sub
    End If
    
    'メイン処理
    fileName = Dir(folderPath & "\*.tif*")
    Do While fileName <> ""
        ProcessSingleFile folderPath, fileName, processedFiles, errorCount
        processedFiles = processedFiles + 1
        Application.StatusBar = "処理中... " & processedFiles & "/" & totalFiles & " ファイル"
        DoEvents
        fileName = Dir()
    Loop
    
    '結果表示
    MsgBox "処理が完了しました" & vbCrLf & _
           "総ファイル数: " & totalFiles & vbCrLf & _
           "処理済: " & processedFiles & vbCrLf & _
           "エラー: " & errorCount, vbInformation
           
    Application.StatusBar = False
End Sub

Private Sub ProcessSingleFile(folderPath As String, fileName As String, ByRef processedFiles As Long, ByRef errorCount As Long)
    Dim wsh As Object
    Dim fname As String
    
    On Error GoTo ErrorHandler
    
    'ファイル名から拡張子を除いた名前を取得
    fname = Left(fileName, InStrRev(fileName, ".") - 1)
    
    'WScript.Shellオブジェクトを作成
    Set wsh = CreateObject("WScript.Shell")
    
    'ペイントウィンドウをアクティブに
    wsh.AppActivate "ペイント"
    Wait 0.5
    
    'TIFFファイルを開く
    SendKeys "^o", True  'Ctrl+O
    Wait 0.5
    SendKeys folderPath & "\" & fileName, True
    SendKeys "{ENTER}", True
    Wait 0.5
    
    'グレースケールに変換
    SendKeys "%i", True  'ALTでメニューを開く
    SendKeys "g", True   'グレースケール
    Wait 0.5
    
    '同じ名前でTIFF保存
    SendKeys "^s", True  'Ctrl+S
    Wait 0.5
    SendKeys "{ENTER}", True
    
    'JPEGとして保存
    SendKeys "^+s", True 'Ctrl+Shift+S
    Wait 0.5
    SendKeys fname & ".jpg", True
    SendKeys "{TAB}", True
    SendKeys "j", True   'JPEG選択
    SendKeys "{ENTER}", True
    Wait 0.5
    
    Exit Sub

ErrorHandler:
    errorCount = errorCount + 1
    MsgBox "エラーが発生しました: " & fileName & vbCrLf & Err.Description, vbExclamation
    Resume Next
End Sub