'insertWordIntoExcel 250226-01
Sub ImportQuestionFromWord()
    ' オブジェクトの宣言
    Dim targetSheet As Worksheet
    Dim textContent As String
    Dim lines() As String
    Dim skipNextLine As Boolean
    Dim i As Long
    Dim j As Long
    
    ' 追加の変数宣言
    Dim targetRow As Long
    Dim currentLine As String
    Dim startLine As Long
    Dim endLine As Long
    Dim fullContent As String
    Dim allLines() As String
    Dim lineCount As Long
    Dim longText As String
    Dim answer As String
    Dim startLineOfLongText As Long
    Dim endLineOfLongText As Long
    Dim text1 As String, text2 As String, text3 As String, text4 As String
    Dim patternProcessed As Boolean ' 패턴 처리 완료 여부를 추적하는 변수
    
    ' エラーハンドリングの設定
    On Error GoTo ErrorHandler
    
    ' 画面更新を無効化してパフォーマンスを向上
    Application.ScreenUpdating = False
    
    Debug.Print "処理開始"
        
    ' Wordファイルの選択ダイアログを表示
    Dim filePath As String
    With Application.FileDialog(msoFileDialogFilePicker)
        .title = "問題が入っているWordファイルを選択してください"
        .Filters.Clear
        .Filters.Add "Word文書", "*.docx; *.doc"
        If .Show = -1 Then
            filePath = .SelectedItems(1)
            Debug.Print "選択されたファイル: " & filePath
        Else
            Debug.Print "ファイル選択がキャンセルされました"
            Exit Sub
        End If
    End With
    
    Set targetSheet = ActiveSheet
    
    Debug.Print "対象シートを現在のアクティブシートに設定完了"

    Dim lastRow As Long
    lastRow = targetSheet.Cells(targetSheet.Rows.Count, "A").End(xlUp).Row
    
    ' Wordファイルを開いてテキストとして読み込む
    Debug.Print "Wordアプリケーション作成開始"
    Set wdApp = New Word.Application
    Debug.Print "Wordアプリケーション作成完了"

    Debug.Print "Wordファイルを開く: " & filePath
    Set wdDoc = wdApp.Documents.Open(filePath)
    Debug.Print "Wordファイル開く完了"

    ' 全テキストを取得
    fullContent = wdDoc.Content.Text

    fullContent = Replace(fullContent, Chr(12), "")
    fullContent = Replace(fullContent, Chr(14), "")
        
    ' テキストを行に分割して全行を一時配列に格納
    allLines = Split(Replace(fullContent, Chr(13), vbNewLine), vbNewLine)

    ' A1の検索と配列の動的設定
    Debug.Print "行数：　" & UBound(allLines) + 1
    Debug.Print "行の処理"

    ' エラー処理を追加
    If UBound(allLines) < 0 Then
        MsgBox "テキストの分割に失敗しました。", vbCritical
        GoTo CleanUp
    End If
    
    ' まず、A1の数をカウント
    Dim a1Count As Integer
    a1Count = 0
    
    For i = 0 To UBound(allLines)
        ' 行の内容をトリムしてから検査
        Dim tempLine As String
        tempLine = Trim(allLines(i))
        
        If Left(tempLine, 2) = "A1" Then
            a1Count = a1Count + 1
        End If
    Next i
    
    If a1Count = 0 Then
        MsgBox "A1が見つかりませんでした。", vbExclamation
        GoTo CleanUp
    End If
    
    Debug.Print "A1が" & a1Count & "個見つかりました"
    
    ' A1の位置を保存する配列を動的に設定
    Dim a1Positions() As Long
    ReDim a1Positions(a1Count - 1) ' 0から始まるインデックスのため-1
    
    ' 配列サイズが設定されたことを確認
    Debug.Print "配列サイズを" & a1Count & "に設定しました"
    
    ' 改めてA1位置を検索してデバッグ情報を出力
    Dim foundCount As Integer
    foundCount = 0
    
    For i = 0 To UBound(allLines)
        tempLine = Trim(allLines(i))
        
        If Left(tempLine, 2) = "A1" Then
            a1Positions(foundCount) = i
            Debug.Print (foundCount + 1) & "番目のA1が見つかりました: " & i & "行目, 内容: " & tempLine
            foundCount = foundCount + 1
        End If
    Next i
    
    ' 処理開始行
    targetRow = 3 ' 시작 행 설정
    
    ' すべてのA1パターンを処理
    For patternIndex = 0 To a1Count - 1  ' すべてのA1パターンを処理
        patternProcessed = False ' 패턴 처리 시작 시 초기화
        startLine = a1Positions(patternIndex)
        
        ' 最後のパターンは文書の最後まで
        If patternIndex = a1Count - 1 Then
            endLine = UBound(allLines)
        Else
            endLine = a1Positions(patternIndex + 1) - 1
        End If
        
        ' 処理範囲の行数を計算
        lineCount = endLine - startLine + 1
        
        ' 新しい配列を作成して必要な範囲のみコピー
        ReDim lines(lineCount - 1)
        
        ' 指定範囲の行を新しい配列にコピー
        For j = 0 To lineCount - 1
            lines(j) = allLines(startLine + j)
        Next j
        
        ' 連続する改行を1つに置換
        textContent = Join(lines, vbNewLine)
        Do While InStr(textContent, vbNewLine & vbNewLine & vbNewLine) > 0
            textContent = Replace(textContent, vbNewLine & vbNewLine & vbNewLine, vbNewLine & vbNewLine)
        Loop
        
        Debug.Print "パターン" & (patternIndex + 1) & "のテキスト前処理完了"
        Debug.Print "処理対象行数: " & lineCount
        
        ' 処理のための変数を初期化
        longText = ""
        
        ' 各行を処理
        For i = 0 To UBound(lines)
            If skipNextLine Then
                skipNextLine = False
                GoTo NextLine
            End If

            currentLine = Trim(lines(i))
            
            ' 空行をスキップ
            If currentLine = "" Then
                GoTo NextLine
            End If

            ' 問題番号の行を検出
            If Left(currentLine, 2) = "A1" Then
                ' 前の処理内容をクリア
                longText = ""
                
                ' デバッグ情報を出力
                Debug.Print "問題番号: " & currentLine & ", 行インデックス: " & i & ", 配列サイズ: " & UBound(lines)
                
                ' エラーハンドリングを追加
                On Error Resume Next
                
                ' 問題番号を記録 - クリーンな文字列のみを使用
                Dim cleanedLine As String
                cleanedLine = Application.WorksheetFunction.Clean(currentLine)
                
                ' 問題発生箇所
                Debug.Print "セルに書き込む前"
                targetSheet.Cells(targetRow, COL_AAAA).Value = cleanedLine
                
                ' エラーチェック
                If Err.Number <> 0 Then
                    Debug.Print "エラー発生: " & Err.Number & " - " & Err.Description
                    Err.Clear
                End If
                Debug.Print "セルに書き込み後"
                
                On Error GoTo ErrorHandler
                
                ' 次の番号を記録 - 範囲チェックを厳密に行う
                If i + 2 <= UBound(lines) Then
                    Debug.Print "次の行の内容: " & lines(i + 2)
                    targetSheet.Cells(targetRow, COL_BBBB).Value = lines(i + 2)
                Else
                    Debug.Print "次の行が範囲外です"
                End If
                
                ' 長文開始位置を設定 - 範囲チェックと安全な既定値
                startLineOfLongText = i + 4
                If startLineOfLongText > UBound(lines) Then
                    startLineOfLongText = i + 3
                    If startLineOfLongText > UBound(lines) Then
                        startLineOfLongText = i + 2
                    End If
                End If
                
                Debug.Print "長文開始位置: " & startLineOfLongText
                
                ' もし長文開始位置が配列の最大を超える場合は処理をスキップ
                If startLineOfLongText > UBound(lines) Then
                    Debug.Print "警告: 長文開始位置が配列サイズを超えています"
                    GoTo NextLine
                End If
            
            ElseIf Left(currentLine, 4) = "(11)" Then
                ' (11)が見つかったら、ここまでを長文として扱う
                endLineOfLongText = i - 1
                
                ' 長文を結合
                longText = ""
                For j = startLineOfLongText To endLineOfLongText
                    If j <= UBound(lines) Then
                        longText = longText & lines(j) & vbCrLf
                    End If
                Next j
                
                targetSheet.Cells(targetRow, COL_EEEE).Value = longText
                
                ' 選択肢1の処理
                text1 = ""
                If i + 1 <= UBound(lines) And Left(Trim(lines(i + 1)), 1) = "1" Then
                    text1 = Trim(Mid(lines(i + 1), 2))
                    targetSheet.Cells(targetRow, COL_1EEEE).Value = text1
                End If
                
                ' 選択肢2の処理
                text2 = ""
                If i + 2 <= UBound(lines) And Left(Trim(lines(i + 2)), 1) = "2" Then
                    text2 = Trim(Mid(lines(i + 2), 2))
                    targetSheet.Cells(targetRow, COL_2EEEE).Value = text2
                End If
                
                ' 選択肢3の処理
                text3 = ""
                If i + 3 <= UBound(lines) And Left(Trim(lines(i + 3)), 1) = "3" Then
                    text3 = Trim(Mid(lines(i + 3), 2))
                    targetSheet.Cells(targetRow, COL_3EEEE).Value = text3
                End If
                
                ' 選択肢4の処理
                text4 = ""
                If i + 4 <= UBound(lines) And Left(Trim(lines(i + 4)), 1) = "4" Then
                    text4 = Trim(Mid(lines(i + 4), 2))
                    targetSheet.Cells(targetRow, COL_4EEEE).Value = text4
                End If
                
                ' 解答の処理
                If i + 5 <= UBound(lines) And Left(Trim(lines(i + 5)), 7) = "Answer:" Then
                    answer = Right(Trim(lines(i + 5)), 1)
                    targetSheet.Cells(targetRow, COL_CCCC).Value = answer
                End If

            ElseIf Left(currentLine, 4) = "(12)" Then
                ' 選択肢1の処理
                text1 = ""
                If i + 1 <= UBound(lines) And Left(Trim(lines(i + 1)), 1) = "1" Then
                    text1 = Trim(Mid(lines(i + 1), 2))
                    targetSheet.Cells(targetRow, COL_11EEEE).Value = text1
                End If
                
                ' 選択肢2の処理
                text2 = ""
                If i + 2 <= UBound(lines) And Left(Trim(lines(i + 2)), 1) = "2" Then
                    text2 = Trim(Mid(lines(i + 2), 2))
                    targetSheet.Cells(targetRow, COL_12EEEE).Value = text2
                End If
                
                ' 選択肢3の処理
                text3 = ""
                If i + 3 <= UBound(lines) And Left(Trim(lines(i + 3)), 1) = "3" Then
                    text3 = Trim(Mid(lines(i + 3), 2))
                    targetSheet.Cells(targetRow, COL_13EEEE).Value = text3
                End If
                
                ' 選択肢4の処理
                text4 = ""
                If i + 4 <= UBound(lines) And Left(Trim(lines(i + 4)), 1) = "4" Then
                    text4 = Trim(Mid(lines(i + 4), 2))
                    targetSheet.Cells(targetRow, COL_14EEEE).Value = text4
                End If
                
                ' 解答の処理
                If i + 5 <= UBound(lines) And Left(Trim(lines(i + 5)), 7) = "Answer:" Then
                    answer = Right(Trim(lines(i + 5)), 1)
                    targetSheet.Cells(targetRow, COL_12CCCC).Value = answer
                End If

            ElseIf Left(currentLine, 4) = "(13)" Then
                ' 選択肢1の処理
                text1 = ""
                If i + 1 <= UBound(lines) And Left(Trim(lines(i + 1)), 1) = "1" Then
                    text1 = Trim(Mid(lines(i + 1), 2))
                    targetSheet.Cells(targetRow, COL_21EEEE).Value = text1
                End If
                
                ' 選択肢2の処理
                text2 = ""
                If i + 2 <= UBound(lines) And Left(Trim(lines(i + 2)), 1) = "2" Then
                    text2 = Trim(Mid(lines(i + 2), 2))
                    targetSheet.Cells(targetRow, COL_22EEEE).Value = text2
                End If
                
                ' 選択肢3の処理
                text3 = ""
                If i + 3 <= UBound(lines) And Left(Trim(lines(i + 3)), 1) = "3" Then
                    text3 = Trim(Mid(lines(i + 3), 2))
                    targetSheet.Cells(targetRow, COL_23EEEE).Value = text3
                End If
                
                ' 選択肢4の処理
                text4 = ""
                If i + 4 <= UBound(lines) And Left(Trim(lines(i + 4)), 1) = "4" Then
                    text4 = Trim(Mid(lines(i + 4), 2))
                    targetSheet.Cells(targetRow, COL_24EEEE).Value = text4
                End If
                
                ' 解答の処理
                If i + 5 <= UBound(lines) And Left(Trim(lines(i + 5)), 7) = "Answer:" Then
                    answer = Right(Trim(lines(i + 5)), 1)
                    targetSheet.Cells(targetRow, COL_13CCCC).Value = answer
                End If
                
                ' パターン処理完了フラグを設定
                patternProcessed = True
            End If
            
NextLine:
        Next i
        
        ' パターン処理が完了したら次の行に進む
        If patternProcessed Then
            targetRow = targetRow + 1 ' パターン処理完了時のみ行を進める
            Debug.Print "パターン" & (patternIndex + 1) & "が完全に処理されたため、次の行へ: " & targetRow
        Else
            Debug.Print "パターン" & (patternIndex + 1) & "の処理が不完全なため、行は進めません: " & targetRow
        End If
        
        ' このパターンの処理が終了したら次のパターンに進む
        Debug.Print "パターン" & (patternIndex + 1) & "の処理完了"
    Next patternIndex
    
    Debug.Print "全テキスト処理完了"
    
CleanUp:
    ' Wordオブジェクトの解放
    If Not wdDoc Is Nothing Then
        wdDoc.Close SaveChanges:=False
        Set wdDoc = Nothing
    End If
    
    If Not wdApp Is Nothing Then
        wdApp.Quit
        Set wdApp = Nothing
    End If
    
    Debug.Print "Wordオブジェクト解放完了"
    
    ' 画面更新を再度有効化
    Application.ScreenUpdating = True
    Application.DisplayAlerts = True
    
    Debug.Print "処理完了"
    Exit Sub

ErrorHandler:
    ' エラー発生時の処理
    Dim errMsg As String
    errMsg = "エラーが発生しました:" & vbCrLf & _
            "エラー番号: " & Err.Number & vbCrLf & _
            "エラーの説明: " & Err.Description & vbCrLf & _
            "エラー発生箇所: " & Err.Source
    
    Debug.Print errMsg
    MsgBox errMsg
    
    ' クリーンアップルーチンへ
    GoTo CleanUp
End Sub