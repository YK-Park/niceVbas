'2nd textToPpt 250215-1

Private pptApp As Object        ' PowerPointアプリケーションオブジェクト
Private pptPres As Object       ' PowerPointプレゼンテーションオブジェクト
Private gSaveFolderPath As String   ' 保存先フォルダパス
Private gTextFileName As String     ' テキストファイル名

' 定数の宣言
Private Const SECTION_START = "1A"    ' セクション開始マーカー
Private Const SECTION_END = "1B"      ' セクション終了マーカー

Public Sub ImportTextToPPT()
    On Error GoTo ErrorHandler
    
    ' PowerPointアプリケーションの初期化
    InitializePowerPoint
    
    ' フォルダパスの取得と検証
    If Not ValidateFolderPath Then Exit Sub
    
    ' PowerPointファイルの選択と開く
    If Not OpenPowerPointFile Then Exit Sub
    
    ' テキストファイルの選択と処理
    Dim selectedPath As String
    selectedPath = ProcessTextFile()
    
    If selectedPath <> "" Then
        Call insertText1(selectedPath)
    End If
    
    Exit Sub

ErrorHandler:
    MsgBox "エラーが発生しました: " & Err.Description, vbCritical
    CleanupObjects
End Sub

Private Sub InitializePowerPoint()
    On Error Resume Next
    
    If pptApp Is Nothing Then
        Set pptApp = GetObject(, "PowerPoint.Application")
        If pptApp Is Nothing Then
            Set pptApp = CreateObject("PowerPoint.Application")
        End If
    End If
    
    On Error GoTo ErrorHandler
    If Not pptApp Is Nothing Then
        pptApp.Visible = True
    End If
    Exit Sub
    
ErrorHandler:
    MsgBox "PowerPointの初期化中にエラーが発生しました: " & Err.Description, vbCritical
    End
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
    On Error GoTo ErrorHandler
    
    ' PowerPointファイルを選択して開く処理
    Dim strFile As Variant
    OpenPowerPointFile = False
    
    strFile = Application.GetOpenFilename("PowerPointファイル (*.pptm), *.pptm", , "PowerPointファイルを選択してください")
    
    If strFile <> False Then
        If pptApp Is Nothing Then
            Set pptApp = New PowerPoint.Application
            pptApp.Visible = True
        End If
        Set pptPres = pptApp.Presentations.Open(CStr(strFile))
        OpenPowerPointFile = True
    End If
    Exit Function
    
ErrorHandler:
    MsgBox "PowerPointファイルを開く際にエラーが発生しました: " & Err.Description, vbCritical
    OpenPowerPointFile = False
End Function

Private Function ProcessTextFile() As String
    ' テキストファイルの選択と処理
    Dim strFile As String
    
    strFile = Application.GetOpenFilename("テキストファイル (*.txt), *.txt", , "テキストファイルを選択してください")
    
    If strFile <> "False" Then
        gTextFileName = GetFileName(strFile)
        ProcessTextFile = strFile
    Else
        ProcessTextFile = ""
    End If
End Function

Private Function ReadTextFile(ByVal filePath As String) As String()
    ' ローカル変数の宣言
    Dim fileNum As Integer
    Dim txtLine As String
    Dim nonEmptyLines As String
    Dim result() As String
    
    On Error GoTo ErrorHandler
    
    fileNum = FreeFile
    Open filePath For Input As #fileNum
    Do Until EOF(fileNum)
        Line Input #fileNum, txtLine
        If Trim(txtLine) <> "" Then
            nonEmptyLines = nonEmptyLines & txtLine & vbCrLf
        End If
    Loop
    Close #fileNum
    
    result = Split(nonEmptyLines, vbCrLf)
    ReadTextFile = result
    Exit Function

ErrorHandler:
    If fileNum > 0 Then Close #fileNum
    MsgBox "ファイル読み込み中エラーが発生しました。" & Err.Description, vbCritical
End Function

Private Function ImportText(insertText As String, targetWord As String, textData() As String) As String
    ' ローカル変数の宣言
    Dim startLine As Long
    Dim endLine As Long
    Dim i As Long
    
    For i = 0 To UBound(textData)
        If InStr(textData(i), SECTION_START) > 0 Then
            startLine = i
            Exit For
        End If
    Next i
    
    For i = 0 To UBound(textData)
        If InStr(textData(i), SECTION_END) > 0 Then
            endLine = i
            Exit For
        End If
    Next i
    
    If startLine <= endLine And endLine <= UBound(textData) Then
        For i = startLine To endLine
            If Left(insertText, 5) = targetWord Then
                insertText = insertText & textData(i) & vbCrLf
            End If
        Next i
    End If
    
    ImportText = insertText
End Function

Private Sub insertText1(ByVal filePath As String)
    Dim insertText As String
    Dim slide As Object
    Dim textData() As String
    
    ' テキストデータの読み込み
    textData = ReadTextFile(filePath)
    
    insertText = ""
    insertText = ImportText(insertText, "no.1", textData)
    insertText = Replace(insertText, "no.1", "no.1" & vbCrLf)
    
    Set slide = pptPres.Slides(17)
    UpdateShapeText slide, "A", insertText
End Sub

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

Private Sub CleanupObjects()
    ' オブジェクトのクリーンアップ
    If Not pptPres Is Nothing Then
        pptPres.Close
        Set pptPres = Nothing
    End If
    
    If Not pptApp Is Nothing Then
        pptApp.Quit
        Set pptApp = Nothing
    End If
End Sub

Private Function GetFileName(ByVal filePath As String) As String
    ' ファイルパスからファイル名を取得
    Dim pos As Integer
    pos = InStrRev(filePath, "\")
    If pos > 0 Then
        GetFileName = Mid(filePath, pos + 1)
    Else
        GetFileName = filePath
    End If
End Function