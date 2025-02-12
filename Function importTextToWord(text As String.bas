Function importTextToWord(text As String, startNum As Integer, startLineNum As Integer, _
                        endNum As Integer, endLineNum As Integer) As String
    text = ""
    
    ' 開始行を検索
    For i = 0 To UBound(lines)
        If InStr(lines(i), "(" & startNum & ")") > 0 Then
            startLine = i + startLineNum 
            Exit For
        End If
    Next i
    
    ' 終了行を検索
    For i = 0 To UBound(lines)
        If InStr(lines(i), "(" & endNum & ")") > 0 Then
            endLine = i + endLineNum 
            Exit For
        End If
    Next i

    ' テキストを構築
    If startLine <= endLine And endLine <= UBound(lines) Then
        For i = startLine To endLine
            text = text & lines(i) & vbCrLf
        Next i
    End If

    importTextToWord = text
End Function

Function formatSpeakers(ByVal insertText As String) As String
    Dim lines() As String
    Dim resultText As String
    Dim firstSpeakerPosition As Long
    
    lines = Split(insertText, vbCrLf)
    resultText = ""
    firstSpeakerPosition = -1
    
    For i = 0 To UBound(lines)
        Dim currentLine As String
        Dim colonPos As Long
        currentLine = lines(i)
        colonPos = InStr(currentLine, ":")
        
        If colonPos > 0 Then
            ' 문장이 (로 시작하면 문제 번호로 간주
            If Left(Trim(currentLine), 1) = "(" Then
                ' 첫 번째 화자의 위치 저장 (탭 포함)
                If firstSpeakerPosition = -1 Then
                    firstSpeakerPosition = InStr(currentLine, vbTab) + Len(vbTab)
                End If
                
                ' (1) 다음의 탭 유지
                If InStr(currentLine, vbTab) = 0 Then
                    currentLine = Left(currentLine, InStr(currentLine, ")")) & vbTab & Mid(currentLine, InStr(currentLine, ")") + 1)
                End If
            Else
                ' 두 번째 이후 화자 정렬
                If firstSpeakerPosition > 0 Then
                    Dim speakerPart As String
                    speakerPart = Left(currentLine, colonPos - 1)
                    currentLine = Space(firstSpeakerPosition - 1) & Trim(speakerPart) & ":" & Mid(currentLine, colonPos + 1)
                End If
            End If
        End If
        
        resultText = resultText & currentLine
        If i < UBound(lines) Then
            resultText = resultText & vbCrLf
        End If
    Next i
    
    formatSpeakers = resultText
End Function

Sub ApplyBoldToSpeakers(rng As Range)
    If rng Is Nothing Then Exit Sub
    
    Dim para As Paragraph
    For Each para In rng.Paragraphs
        Dim paraText As String
        Dim colonPos As Long
        paraText = para.Range.Text
        colonPos = InStr(paraText, ":")
        
        If colonPos > 0 Then
            ' 문장이 (로 시작하면 괄호 이후부터 볼드 처리
            If Left(Trim(paraText), 1) = "(" Then
                Dim bracketEnd As Long
                bracketEnd = InStr(paraText, ")") + 1
                
                ' 괄호 다음 탭 이후부터 콜론까지 볼드체
                With para.Range
                    .SetRange para.Range.Start + bracketEnd + 1, para.Range.Start + colonPos - 1
                    .Bold = True
                End With
            Else
                ' 문장 시작부터 콜론까지 볼드체
                With para.Range
                    .SetRange para.Range.Start, para.Range.Start + colonPos - 1
                    .Bold = True
                End With
            End If
        End If
    Next para
End Sub

Function replaceThings(ByVal originalText As String) As String
    Dim resultText As String
    resultText = originalText

    ' 選択肢の空白を調整
    resultText = Replace(resultText, "1　", "1" & Space(2))
    resultText = Replace(resultText, "2　", "2" & Space(2))
    resultText = Replace(resultText, "3　", "3" & Space(2))
    resultText = Replace(resultText, "4　", "4" & Space(2))
    
    replaceThings = resultText
End Function

Sub Section1()
    On Error GoTo ErrorHandler
    
    Dim insertText As String
    Dim doc As Document
    Dim rng As Range
    
    ' テキストを読み込み
    insertText = ""
    insertText = importTextToWord(insertText, 1, 0, 5, -2)

    ' フォーマットを適用
    insertText = Replace(insertText, vbTab & "3　", vbCrLf & "3　")
    insertText = replaceThings(insertText)

     insertText = formatSpeakers(insertText)

    Set doc = ActiveDocument
    Set rng = Selection.Range

    ' 基本フォーマットを適用
    With rng
        .Style = doc.Styles("標準")
        .ParagraphFormat.Reset
        .ParagraphFormat.LineSpacingRule = 5
        .ParagraphFormat.LineSpacing = 15
        
        .Font.Size = 18
        .Font.Color = vbBlack
        .Font.Bold = False
        .Font.Name = "Arial"

        .ParagraphFormat.Alignment = wdAlignParagraphJustify
        
        .text = insertTextB4PageBreaks(insertText, 2)

        .ParagraphFormat.CharacterUnitLeftIndent = 0
        .ParagraphFormat.CharacterUnitFirstLineIndent = 0
        .ParagraphFormat.LeftIndent = 1.5 * CM_TO_PONITS
        .ParagraphFormat.FirstLineIndent = -1.5 * CM_TO_PONITS
    End With
    
    ' 話者のフォーマットを適用
    Call ApplySpeakerBold(rng)

Cleanup:
    Set doc = Nothing
    Set rng = Nothing
    Exit Sub

ErrorHandler:
    MsgBox "エラーが発生しました: " & Err.Description
    Resume Cleanup
End Sub

Sub Listening1()
    
    insertText = ""

    For i = 0 To UBound(lines)
        If InStr(lines(i), "No　.　") > 0 Then
            NoCount = NoCount + 1
            If NoCount <= 10 Then
                lines(i) = Replace(lines(i), vbTabm vbCrLf)
                insertText = insertText & lines(i) & vbCrLf

                NoLine = i
                For j = NoLine To NoLine + 4
                    If j <= UBound(lines) Then
                        Dim lineContent As String
                        lineContent = Trim(Replace(lines(j), vbTab, ""))
                        insertText = insertText & lineContent & vbCrLf
                    End If
                Next j
            Else
                Exit For
            End If
        End If
    Next i
   
   ' アクティブなドキュメントを設定
    Set doc = ActiveDocument
   
    ' 改ページを上から数えて9番目の位置を特定
    Dim foundPageBreaks As Integer
    foundPageBreaks = 0
    Dim rng As Range
    Set rng = doc.Range(0, 0)
    Do While rng.Find.Execute(FindText:=Chr(12))
        foundPageBreaks = foundPageBreaks + 1
        If foundPageBreaks = 9 Then
            Exit Do
        End If
    Loop

    ' 9つ目の改ページの前にテキストを挿入
    If foundPageBreaks = 9 Then
        rng.Collapse direction:=wdCollapseStart
        rng.InsertBefore insertText
    Else
        MsgBox "文書に9つ目の改ページが見つかりませんでした。"
    End If

Cleanup:
    ' オブジェクトのクリーンアップ
    Set dlg = Nothing
    Set doc = Nothing
    Set rng = Nothing
    Exit Sub

ErrorHandler:
    MsgBox "エラーが発生しました: " & Err.Description
    Resume Cleanup
    
End Sub


Function DebugSpecialChars(ByVal text As String) As String
    Dim result As String
    Dim i As Long
    Dim char As String
    
    For i = 1 To Len(text)
        char = Mid(text, i, 1)
        Select Case AscW(char)
            Case 9    ' Tab
                result = result & "[TAB]"
            Case 10   ' Line Feed (LF)
                result = result & "[LF]"
            Case 13   ' Carriage Return (CR)
                result = result & "[CR]"
            Case 32   ' Space (Half-width)ㅣㄹ
                result = result & "[SP]"
            Case 12288 ' Full-width Space (全角スペース)
                result = result & "[全SP]"
            Case Else
                result = result & char
        End Select
    Next i
    
    ' Debug.Print로 출력
    Debug.Print "==== Text Length: " & Len(text) & " chars ===="
    Debug.Print result
    Debug.Print "====================================="
    
    DebugSpecialChars = result
End Function

Private Sub DebugSelectedText()
    ' 選択されているテキストの各文字のコードを確認
    Dim selectedText As String
    Dim i As Long
    Dim charCode As Long
    
    ' 選択されているテキストを取得
    selectedText = Selection.Text
    
    Debug.Print "=== 選択テキストの文字コード分析 ==="
    Debug.Print "全長: " & Len(selectedText) & " 文字"
    
    For i = 1 To Len(selectedText)
        charCode = AscW(Mid(selectedText, i, 1))
        Debug.Print "位置 " & i & ": " & _
                   "文字[" & Mid(selectedText, i, 1) & "] " & _
                   "コード[" & charCode & "]"
    Next i
    
    ' 段落の書式情報も確認
    With Selection.ParagraphFormat
        Debug.Print "=== 段落書式 ==="
        Debug.Print "左インデント: " & .LeftIndent
        Debug.Print "タブ設定数: " & .TabStops.Count
        
        ' タブ位置の確認
        If .TabStops.Count > 0 Then
            Dim ts As TabStop
            Debug.Print "=== タブ位置 ==="
            For Each ts In .TabStops
                Debug.Print "位置: " & ts.Position & ", 種類: " & ts.Alignment
            Next ts
        End If
    End With
    
    Debug.Print "==============================="
End Sub

Public Sub QuickDebug()
    If Selection.Type = wdSelectionIP Or Selection.Type = wdSelectionNormal Then
        Call DebugSelectedText
    Else
        MsgBox "テキストが選択されていません。"
    End If
End Sub