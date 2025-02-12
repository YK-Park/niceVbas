'word target tool 250212-02
Public Sub SearchRelatedWords()
    '*** メイン処理を行うサブルーチン ***
    '*** Sheets(4)のA2に入力された単語を基に類似単語を検索し、C列以降に結果を表示する ***
    
    Dim wsTarget As Worksheet
    Dim wsList As Worksheet
    Dim targetWord As String
    Dim lastRow As Long
    Dim i As Long
    Dim resultRow As Long
    
    Application.ScreenUpdating = False
    
    On Error GoTo ErrorHandler
    
    Set wsTarget = ThisWorkbook.Sheets(4)
    Set wsList = ThisWorkbook.Sheets("単語リスト")
    
    '*** ヘッダーの設定（存在しない場合のみ） ***
    If wsTarget.Range("A1").Value = "" Then
        wsTarget.Range("A1").Value = "検索単語"
        wsTarget.Range("C1").Value = "級番号"
        wsTarget.Range("D1").Value = "ユニーク番号"
        wsTarget.Range("E1").Value = "級"
        wsTarget.Range("F1").Value = "単語"
        wsTarget.Range("G1").Value = "品詞"
        wsTarget.Range("H1").Value = "出題区分"
        
        '*** ヘッダー行の書式設定 ***
        With wsTarget.Range("A1:H1")
            .Interior.Color = RGB(220, 230, 241)  '*** 薄い青色 ***
            .Font.Bold = True
            .Borders.LineStyle = xlContinuous
            .Font.Name = "メイリオ"
            .Font.Size = 11
            .EntireColumn.AutoFit
        End With
    End If
    
    '*** 基準単語の取得 ***
    targetWord = LCase(Trim(wsTarget.Range("A2").Value))
    If targetWord = "" Then
        MsgBox "A2セルに検索する単語を入力してください。", vbExclamation
        Exit Sub
    End If
    
    '*** 結果エリアのクリア（ヘッダー行を除く） ***
    wsTarget.Range("C2:H" & wsTarget.Rows.Count).ClearContents
    
    lastRow = wsList.Cells(wsList.Rows.Count, "D").End(xlUp).Row
    resultRow = 2  '*** ヘッダー行の次から開始 ***
    
    '*** 第一段階：類似単語の検索 ***
    For i = 2 To lastRow
        Dim currentWord As String
        currentWord = LCase(Trim(wsList.Cells(i, "D").Value))
        
        If currentWord <> "" And currentWord <> targetWord Then
            If CompareWords(targetWord, currentWord) Then
                '*** 結果の書き込み ***
                wsTarget.Cells(resultRow, "C").Value = wsList.Cells(i, "A").Value  '級番号
                wsTarget.Cells(resultRow, "D").Value = wsList.Cells(i, "B").Value  'ユニーク番号
                wsTarget.Cells(resultRow, "E").Value = wsList.Cells(i, "C").Value  '級
                wsTarget.Cells(resultRow, "F").Value = wsList.Cells(i, "D").Value  '単語
                wsTarget.Cells(resultRow, "G").Value = wsList.Cells(i, "E").Value  '品詞
                wsTarget.Cells(resultRow, "H").Value = wsList.Cells(i, "F").Value  '出題区分
                resultRow = resultRow + 1
            End If
        End If
    Next i
    
    '*** 第二段階：派生語の除去 ***
    If resultRow > 2 Then
        FilterDerivatives wsTarget, 2  '*** ヘッダー行を除いて処理 ***
        
        '*** 結果の書式設定 ***
        With wsTarget.Range("C1:H" & resultRow - 1)
            .Borders.LineStyle = xlContinuous
            .Font.Name = "メイリオ"
            .Font.Size = 11
            .EntireColumn.AutoFit
        End With
        
        '*** ヘッダー行の書式設定 ***
        With wsTarget.Range("C1:H1")
            .Interior.Color = RGB(220, 230, 241)  '*** 薄い青色 ***
            .Font.Bold = True
        End With
    End If
    
ExitSub:
    Application.ScreenUpdating = True
    
    '*** 検索結果の報告 ***
    If resultRow > 2 Then
        MsgBox (resultRow - 2) & "件の類似単語が見つかりました。", vbInformation
    Else
        MsgBox "該当する単語は見つかりませんでした。", vbInformation
    End If
    Exit Sub
    
ErrorHandler:
    MsgBox "エラーが発生しました: " & Err.Description, vbCritical
    Resume ExitSub
End Sub

Private Function CompareWords(ByVal baseWord As String, ByVal compareWord As String) As Boolean
    '*** 単語比較関数 ***
    '*** 基準となる単語(baseWord)より短い単語は全て含める ***
    '*** 基準となる単語の派生語は除外 ***
    '*** 似ている単語（studio/study など）は含める ***
    
    '*** 両方の単語を小文字に変換して空白を削除 ***
    baseWord = LCase(Trim(baseWord))
    compareWord = LCase(Trim(compareWord))
    
    '*** 比較する単語が基準語より短い場合は常にTrue ***
    If Len(compareWord) < Len(baseWord) Then
        CompareWords = True
        Exit Function
    End If
    
    '*** 基準語が比較語の中に完全に含まれている場合は派生語として除外 ***
    If InStr(compareWord, baseWord) > 0 Then
        CompareWords = False
        Exit Function
    End If
    
    '*** それ以外の単語は含める ***
    CompareWords = True
End Function

'*** 結果から派生語を除去する関数 ***
Private Sub FilterDerivatives(ByVal ws As Worksheet, ByVal startRow As Long)
    Dim lastRow As Long
    Dim i As Long, j As Long
    Dim word1 As String, word2 As String
    
    lastRow = ws.Cells(ws.Rows.Count, "J").End(xlUp).Row
    
    '*** 重複チェック用の配列 ***
    Dim processed() As Boolean
    ReDim processed(startRow To lastRow) As Boolean
    
    For i = startRow To lastRow
        If Not processed(i) Then
            word1 = LCase(Trim(ws.Cells(i, "M").Value))
            
            For j = i + 1 To lastRow
                word2 = LCase(Trim(ws.Cells(j, "M").Value))
                
                '*** 派生語関係をチェック ***
                If InStr(word2, word1) > 0 Or InStr(word1, word2) > 0 Then
                    '*** 短い方を残し、長い方を削除 ***
                    If Len(word1) > Len(word2) Then
                        ws.Rows(i).Delete
                        Exit For
                    Else
                        ws.Rows(j).Delete
                        lastRow = lastRow - 1
                        j = j - 1
                    End If
                End If
            Next j
        End If
    Next i
End Sub