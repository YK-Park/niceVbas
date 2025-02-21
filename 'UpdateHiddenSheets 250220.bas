'UpdateHiddenSheets 250220
폴더 내의 엑셀파일에서 숨겨진 시트의 내용을 변경하고 싶다.
1.각 파일에서 시트명 "AB"의 MN열의 값을 액티브북의 시트"AB"의 AB열의 값으로 변경한다.
2.각 파일에서 시트명 "CD"의 전체 값을 액티브북의 시트"CD"의 값으로 변경한다.

그런데 이 숨겨진 시트들은 드롭다운메뉴의 데이터입력규칙에서 참조하는 부분이다.
이렇게 바꿀 수 있는지?
바꿀 수 있다면, 데이터입력규칙에서도 에러가 나지 않을지?

대상 폴더의 위치가 어디인가요?
액티브북(기준이 되는 엑셀 파일)의 위치는 어디인가요?
폴더 내의 모든 엑셀 파일을 대상으로 하나요, 아니면 특정 파일들만 대상인가요?
데이터 유효성 검사(드롭다운 메뉴)가 사용되는 시트와 셀 범위를 알 수 있을까요?

데이터 유효성 검사와 관련해서 말씀드리면:

숨겨진 시트의 내용을 변경하더라도, 해당 시트를 참조하는 데이터 유효성 검사 규칙은 자동으로 업데이트됩니다.
단, 다음 사항들을 주의해야 합니다:

기존 데이터의 범위가 줄어들면 유효성 검사에 오류가 발생할 수 있습니다
참조하는 셀 범위가 정확히 일치해야 합니다
새로운 데이터가 기존 형식과 일치해야 합니다

1.대상폴더를 선택하도록 하고싶다. 2. 액티브 북은 다른 폴더에 있다
3.폴더 내의 모든 엑셀 파일을 대상으로 4. 1의 경우가 드롭다운 메뉴가 사용되는 셀 범위다. 그런데 데이터 수는 줄어들지만, 공백은 무시가 체크되어있으면 괜찮지 않을까?

Sub UpdateHiddenSheets()
    ' 変数の宣言
    Dim folderPath As String
    Dim fileName As String
    Dim wb As Workbook
    Dim sourceWb As Workbook
    Dim fd As FileDialog
    
    ' エラー処理の開始
    On Error GoTo ErrorHandler
    
    ' アクティブブックを参照用として保存
    Set sourceWb = ThisWorkbook
    
    ' フォルダ選択ダイアログを表示
    Set fd = Application.FileDialog(msoFileDialogFolderPicker)
    With fd
        .Title = "処理するフォルダを選択してください"
        .AllowMultiSelect = False
        If .Show = -1 Then
            folderPath = .SelectedItems(1)
        Else
            MsgBox "フォルダが選択されていません。処理を終了します。"
            Exit Sub
        End If
    End With
    
    ' 画面更新を無効化
    Application.ScreenUpdating = False
    Application.DisplayAlerts = False
    
    ' 選択されたフォルダ内のExcelファイルを処理
    fileName = Dir(folderPath & "\*.xlsx")
    
    Do While fileName <> ""
        ' 処理中のファイル名を表示
        Debug.Print "処理中: " & fileName
        
        ' ファイルを開く
        Set wb = Workbooks.Open(folderPath & "\" & fileName)
        
        ' "AB"シートの更新
        If SheetExists(wb, "AB") Then
            wb.Sheets("AB").Range("M:M").Value = _
                sourceWb.Sheets("AB").Range("A:A").Value
        End If
        
        ' "CD"シートの更新
        If SheetExists(wb, "CD") Then
            sourceWb.Sheets("CD").UsedRange.Copy
            wb.Sheets("CD").UsedRange.PasteSpecial xlPasteValues
        End If
        
        ' 変更を保存して閉じる
        wb.Close SaveChanges:=True
        
        ' 次のファイルを取得
        fileName = Dir()
    Loop
    
ExitSub:
    ' 画面更新を有効化
    Application.ScreenUpdating = True
    Application.DisplayAlerts = True
    MsgBox "処理が完了しました。"
    Exit Sub
    
ErrorHandler:
    MsgBox "エラーが発生しました: " & Err.Description
    Resume ExitSub
End Sub

' シートの存在確認用関数
Function SheetExists(wb As Workbook, sheetName As String) As Boolean
    Dim ws As Worksheet
    
    On Error Resume Next
    Set ws = wb.Sheets(sheetName)
    On Error GoTo 0
    
    SheetExists = Not ws Is Nothing
End Function