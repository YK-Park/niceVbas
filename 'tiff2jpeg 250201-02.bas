'tiff2jpeg 250201-02
Option Explicit

Private Const wiaFormatBMP As String = "{B96B3CAB-0728-11D3-9D7B-0000F81EF32E}"
Private Const wiaFormatJPEG As String = "{B96B3CAE-0728-11D3-9D7B-0000F81EF32E}"

Option Explicit

'// WIA Format Constants
Private Const wiaFormatBMP As Long = 1
Private Const wiaFormatPNG As Long = 2
Private Const wiaFormatGIF As Long = 3
Private Const wiaFormatJPEG As Long = 4
Private Const wiaFormatTIFF As Long = 5

Sub ConvertTIFFtoJPEG()
    '// WIA オブジェクトの作成
    Dim imgFile As Object
    Dim imgProcess As Object
    
    On Error Resume Next
    Set imgProcess = GetObject(, "WIA.ImageProcess")
    If imgProcess Is Nothing Then Set imgProcess = CreateObject("WIA.ImageProcess")
    
    Set imgFile = GetObject(, "WIA.ImageFile")
    If imgFile Is Nothing Then Set imgFile = CreateObject("WIA.ImageFile")
    
    If Err.Number <> 0 Then
        MsgBox "WIAの初期化に失敗しました。(" & Err.Number & ")" & vbCrLf & _
               Err.Description & vbCrLf & _
               "Microsoft Windows Image Acquisition Libraryの参照設定を確認してください。", _
               vbExclamation
        Exit Sub
    End If
    On Error GoTo 0
    
    '// 初期化テスト
    On Error Resume Next
    Debug.Print "Testing WIA initialization..."
    imgProcess.Filters.Clear
    
    If Err.Number <> 0 Then
        MsgBox "WIAオブジェクトの初期化テストに失敗しました。(" & Err.Number & ")" & vbCrLf & _
               Err.Description, vbExclamation
        Exit Sub
    End If
    On Error GoTo 0
    
    '// フォルダ選択ダイアログの表示
    Dim fd As FileDialog
    Set fd = Application.FileDialog(msoFileDialogFolderPicker)
    
    With fd
        .Title = "TIFFファイルがあるフォルダを選択してください"
        .AllowMultiSelect = False
        
        If .Show = False Then Exit Sub
        
        Dim selectedFolder As String
        selectedFolder = .SelectedItems(1)
        Debug.Print "選択されたフォルダ: " & selectedFolder
        
        '// outputフォルダの作成
        Dim outputFolder As String
        outputFolder = CreateOutputFolder(selectedFolder)
        Debug.Print "出力フォルダ: " & outputFolder
        
        '// 選択されたフォルダ内のすべてのTIFFファイルを処理
        Dim fileName As String
        Dim successCount As Long
        Dim errorCount As Long
        successCount = 0
        errorCount = 0
        
        '// フォルダ内のすべての.tifおよび.tiffファイルを検索
        fileName = Dir(selectedFolder & "\*.tif*")
        
        Do While fileName <> ""
            Debug.Print "処理中のファイル: " & fileName
            '// TIFFファイルかどうかを確認（.tifまたは.tiffで終わるか）
            If LCase(Right(fileName, 3)) = "tif" Or LCase(Right(fileName, 4)) = "tiff" Then
                If ProcessTIFFFile(selectedFolder & "\" & fileName, imgFile, imgProcess, outputFolder) Then
                    successCount = successCount + 1
                    Debug.Print "-> 変換成功"
                Else
                    errorCount = errorCount + 1
                    Debug.Print "-> 変換失敗"
                End If
            End If
            fileName = Dir()
        Loop
        
        '// 処理結果を表示
        MsgBox "処理完了" & vbCrLf & _
               "成功: " & successCount & "件" & vbCrLf & _
               "失敗: " & errorCount & "件" & vbCrLf & _
               "出力フォルダ: " & outputFolder
    End With
    
    Set imgProcess = Nothing
    Set imgFile = Nothing
    Set CommonDialog = Nothing
End Sub

Private Function ProcessTIFFFile(ByVal filePath As String, ByVal imgFile As Object, _
                               ByVal imgProcess As Object, ByVal outputFolder As String) As Boolean
    On Error GoTo ErrorHandler
    
    Debug.Print "処理開始: " & filePath
    
    '// TIFFファイルを読み込む
    imgFile.LoadFile filePath
    Debug.Print "-> ファイル読み込み完了"
    Debug.Print "-> ビット深度: " & imgFile.PixelDepth
    
    '// ビット数を確認
    If imgFile.PixelDepth <> 8 Then
        Debug.Print "-> " & imgFile.PixelDepth & "ビットから8ビットに変換中"
        imgProcess.Filters.Clear
        imgProcess.Filters.Add imgProcess.FilterInfos("Convert").FilterID
        imgProcess.Filters(1).Properties("FormatID").Value = wiaFormatBMP
        imgProcess.Filters(1).Properties("Quality").Value = 8
        Set imgFile = imgProcess.Apply(imgFile)
    End If
    
    '// JPEG形式に変換して保存（品質100%）
    imgProcess.Filters.Clear
    imgProcess.Filters.Add imgProcess.FilterInfos("Convert").FilterID
    imgProcess.Filters(1).Properties("FormatID").Value = wiaFormatJPEG
    imgProcess.Filters(1).Properties("Quality").Value = 100
    
    '// 出力ファイル名の設定
    Dim outputPath As String
    outputPath = outputFolder & "\" & GetFileName(filePath)
    outputPath = Replace(outputPath, ".tif", ".jpg")
    outputPath = Replace(outputPath, ".tiff", ".jpg")
    Debug.Print "-> 出力パス: " & outputPath
    
    '// 変換と保存
    Set imgFile = imgProcess.Apply(imgFile)
    imgFile.SaveFile outputPath
    Debug.Print "-> JPEG保存完了"
    
    ProcessTIFFFile = True
    Exit Function

ErrorHandler:
    Debug.Print "エラー発生: " & Err.Number & " - " & Err.Description
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
        Debug.Print "出力フォルダを作成: " & outputFolder
    Else
        Debug.Print "既存の出力フォルダを使用: " & outputFolder
    End If
    
    CreateOutputFolder = outputFolder
End Function

import os
from PIL import Image

def convert_to_8bit(input_folder):
    # 폴더 내의 모든 TIFF 파일 처리
    for filename in os.listdir(input_folder):
        if filename.lower().endswith(('.tiff', '.tif')):
            input_path = os.path.join(input_folder, filename)
            output_path = os.path.join(input_folder, 'converted_' + filename)
            
            try:
                with Image.open(input_path) as img:
                    # 8비트로 변환
                    img_8bit = img.convert('L')
                    # 저장
                    img_8bit.save(output_path, 'TIFF')
                print(f'변환 완료: {filename}')
            except Exception as e:
                print(f'에러 발생 ({filename}): {str(e)}')

# 스크립트 실행
if __name__ == "__main__":
    # 현재 디렉토리 사용
    current_folder = os.getcwd()
    convert_to_8bit(current_folder)