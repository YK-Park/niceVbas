'登録用データSJIS変換マクロ 241213-02
Option Explicit

' SJISに対応できない文字をチェックして抽出するサブルーチン
Sub CheckNonSJISCharacters()
    Dim wsData As Worksheet
    Dim resultWs As Worksheet
    Dim usedRange As Range
    Dim cell As Range
    Dim resultRow As Long
    Dim nonSJISChar As String
    Dim i As Long
    Dim lastRow As Long
    Dim lastCol As Long
    
    Application.ScreenUpdating = False
    On Error Resume Next
    
    ' シートの設定
    Set wsData = ThisWorkbook.Worksheets("Data")
    Set resultWs = ThisWorkbook.Worksheets("result")
    
    ' シートが存在しない場合のエラー処理
    If wsData Is Nothing Then
        MsgBox "シート'Data'が見つかりません。", vbCritical
        Exit Sub
    End If
    
    If resultWs Is Nothing Then
        MsgBox "シート'result'が見つかりません。", vbCritical
        Exit Sub
    End If

    ' 既存のボタンを削除してからシートをクリア
    On Error Resume Next
    resultWs.Buttons.Delete
    On Error GoTo 0
    
    ' resultシートの2行目以降をクリア
    If resultWs.Cells(2, 1).Value <> "" Then
        resultWs.Rows("2:" & resultWs.Rows.Count).Clear
    End If
    
    ' 実際に使用されている範囲を取得
    lastRow = wsData.Cells(wsData.Rows.Count, "B").End(xlUp).Row
    lastCol = wsData.Cells(1, wsData.Columns.Count).End(xlToLeft).Column
    
    If lastRow < 2 Then Exit Sub  ' データがない場合は終了
    
    Set usedRange = wsData.Range(wsData.Cells(2, 1), wsData.Cells(lastRow, lastCol))
    
    resultRow = 2  ' ヘッダー行の次から開始
    
    ' 使用範囲内の各セルをチェック
    For Each cell In usedRange
        If Not IsEmpty(cell) Then
            For i = 1 To Len(cell.Value)
                nonSJISChar = Mid(cell.Value, i, 1)
                
                ' 文字のエンコーディング情報を取得
                Dim charInfo As String
                charInfo = CheckCharacterEncoding(nonSJISChar)
                
                ' SJIS非対応の文字のみを結果に出力
                If InStr(charInfo, "SJIS非対応") > 0 Then
                    With resultWs
                        .Cells(resultRow, 1).Value = resultRow - 1
                        .Cells(resultRow, 2).Value = wsData.Cells(cell.Row, 2).Value  ' ID
                        .Cells(resultRow, 3).Value = cell.Address
                        .Cells(resultRow, 4).Value = nonSJISChar
                        .Cells(resultRow, 5).Value = charInfo
                        .Cells(resultRow, 6).Value = GetReplacementSuggestion(nonSJISChar)
                    End With
                    resultRow = resultRow + 1
                End If
            Next i
        End If
        
        If (cell.Row Mod 100) = 0 Then DoEvents
    Next cell
    
    ' 結果シートの列幅を自動調整
    resultWs.Columns("A:G").AutoFit

    ' 変換ボタンを追加
    Call AddConvertButtons

    ' resultシートをアクティブにする
    resultWs.Activate
    resultWs.Range("A1").Select

    Set usedRange = Nothing
    Set wsData = Nothing
    Set resultWs = Nothing
    
    Application.ScreenUpdating = True
    MsgBox "チェック完了。結果はシート 'result' に出力されました。", vbInformation
End Sub

' 文字がSJISに変換可能かどうかをチェックする関数
Function CanConvertToSJIS(character As String) As Boolean
    On Error Resume Next
    Dim test As String
    Dim charCode As Integer
    
    ' ASCIIの範囲（0-127）をチェック
    charCode = AscW(character)
    If charCode >= 0 And charCode <= 127 Then
        CanConvertToSJIS = True
        Exit Function
    End If
    
    ' SJISへの変換を試みる
    test = StrConv(character, vbFromUnicode, 932)
    CanConvertToSJIS = (Err.Number = 0)
    
    Err.Clear
    On Error GoTo 0
End Function

' 文字の変換推奨を提供する関数
Function GetReplacementSuggestion(character As String) As String
    Dim charCode As Long
    charCode = AscW(character)
    
    Select Case charCode
        ' A系
        Case &H100: GetReplacementSuggestion = "A"  ' Ā -> A
        Case &H101: GetReplacementSuggestion = "a"  ' ā -> a
        Case &H102: GetReplacementSuggestion = "A"  ' Ă -> A
        Case &H103: GetReplacementSuggestion = "a"  ' ă -> a
        Case &H104: GetReplacementSuggestion = "A"  ' Ą -> A
        Case &H105: GetReplacementSuggestion = "a"  ' ą -> a
        
        ' E系
        Case &H112: GetReplacementSuggestion = "E"  ' Ē -> E
        Case &H113: GetReplacementSuggestion = "e"  ' ē -> e
        Case &H114: GetReplacementSuggestion = "E"  ' Ĕ -> E
        Case &H115: GetReplacementSuggestion = "e"  ' ĕ -> e
        Case &H116: GetReplacementSuggestion = "E"  ' Ė -> E
        Case &H117: GetReplacementSuggestion = "e"  ' ė -> e
        Case &H118: GetReplacementSuggestion = "E"  ' Ę -> E
        Case &H119: GetReplacementSuggestion = "e"  ' ę -> e
        Case &H11A: GetReplacementSuggestion = "E"  ' Ě -> E
        Case &H11B: GetReplacementSuggestion = "e"  ' ě -> e
        
        ' I系
        Case &H128: GetReplacementSuggestion = "I"  ' Ĩ -> I
        Case &H129: GetReplacementSuggestion = "i"  ' ĩ -> i
        Case &H12A: GetReplacementSuggestion = "I"  ' Ī -> I
        Case &H12B: GetReplacementSuggestion = "i"  ' ī -> i
        Case &H12C: GetReplacementSuggestion = "I"  ' Ĭ -> I
        Case &H12D: GetReplacementSuggestion = "i"  ' ĭ -> i
        Case &H12E: GetReplacementSuggestion = "I"  ' Į -> I
        Case &H12F: GetReplacementSuggestion = "i"  ' į -> i
        
        ' O系
        Case &H14C: GetReplacementSuggestion = "O"  ' Ō -> O
        Case &H14D: GetReplacementSuggestion = "o"  ' ō -> o
        Case &H14E: GetReplacementSuggestion = "O"  ' Ŏ -> O
        Case &H14F: GetReplacementSuggestion = "o"  ' ŏ -> o
        Case &H150: GetReplacementSuggestion = "O"  ' Ő -> O
        Case &H151: GetReplacementSuggestion = "o"  ' ő -> o
        
        ' U系
        Case &H168: GetReplacementSuggestion = "U"  ' Ũ -> U
        Case &H169: GetReplacementSuggestion = "u"  ' ũ -> u
        Case &H16A: GetReplacementSuggestion = "U"  ' Ū -> U
        Case &H16B: GetReplacementSuggestion = "u"  ' ū -> u
        Case &H16C: GetReplacementSuggestion = "U"  ' Ŭ -> U
        Case &H16D: GetReplacementSuggestion = "u"  ' ŭ -> u
        
        ' その他の文字
        Case &H160: GetReplacementSuggestion = "S"  ' Š -> S
        Case &H161: GetReplacementSuggestion = "s"  ' š -> s
        Case &H17D: GetReplacementSuggestion = "Z"  ' Ž -> Z
        Case &H17E: GetReplacementSuggestion = "z"  ' ž -> z
        Case &H178: GetReplacementSuggestion = "Y"  ' Ÿ -> Y
        
        ' 数字・記号類
        Case &H2460: GetReplacementSuggestion = "1"  ' ① -> 1
        Case &H2461: GetReplacementSuggestion = "2"  ' ② -> 2
        Case &H2462: GetReplacementSuggestion = "3"  ' ③ -> 3
        Case &H2463: GetReplacementSuggestion = "4"  ' ④ -> 4
        Case &H2464: GetReplacementSuggestion = "5"  ' ⑤ -> 5
        Case &H2192: GetReplacementSuggestion = "->"  ' → -> ->
        Case &H21D2: GetReplacementSuggestion = "=>"  ' ⇒ -> =>
        Case &H2266: GetReplacementSuggestion = "<="  ' ≦ -> <=
        Case &H2267: GetReplacementSuggestion = ">="  ' ≧ -> >=
        
        Case Else: GetReplacementSuggestion = ""
    End Select
End Function

' 文字のエンコーディング情報をチェックする関数
Function CheckCharacterEncoding(character As String) As String
    Dim resultInfo As String
    Dim charCode As Long
    Dim suggestion As String
    
    charCode = AscW(character)
    resultInfo = "Unicode U+" & Right$("0000" & Hex(charCode), 4)
    
    suggestion = GetReplacementSuggestion(character)
    If suggestion <> "" Then
        If suggestion <> character Then
            resultInfo = resultInfo & " (同等のSJIS文字: " & suggestion & ")"
        End If
    End If
    
    If CanConvertToSJIS(character) Then
        resultInfo = resultInfo & " (SJIS対応可)"
    Else
        resultInfo = resultInfo & " (SJIS非対応)"
    End If
    
    CheckCharacterEncoding = resultInfo
End Function

' 非SJIS文字を変換するサブルーチン
Sub ConvertNonSJISCharacters()
    Dim ws As Worksheet
    Dim usedRange As Range
    Dim cell As Range
    Dim convertedText As String
    Dim lastRow As Long
    Dim lastCol As Long
    
    Application.ScreenUpdating = False
    On Error Resume Next
    
    Set ws = ActiveSheet
    
    lastRow = ws.Cells(ws.Rows.Count, "A").End(xlUp).Row
    lastCol = ws.Cells(1, ws.Columns.Count).End(xlToLeft).Column
    
    If lastRow < 1 Then Exit Sub
    
    Set usedRange = ws.Range(ws.Cells(1, 1), ws.Cells(lastRow, lastCol))
    
    For Each cell In usedRange
        If Not IsEmpty(cell) Then
            convertedText = ConvertToSJIS(CStr(cell.Value))
            If convertedText <> cell.Value Then
                cell.Value = convertedText
            End If
        End If
        
        If (cell.Row Mod 100) = 0 Then DoEvents
    Next cell
    
    Set usedRange = Nothing
    Set ws = Nothing
    
    Application.ScreenUpdating = True
    MsgBox "変換が完了しました。", vbInformation
End Sub

' 非SJIS文字を対応する文字に変換する関数
Private Function ConvertToSJIS(ByVal text As String) As String
    Dim result As String
    Dim i As Long
    Dim currentChar As String
    
    result = text
    
    For i = 1 To Len(text)
        currentChar = Mid(text, i, 1)
        
        If Not CanConvertToSJIS(currentChar) Then
            Dim suggestion As String
            suggestion = GetReplacementSuggestion(currentChar)
            
            If suggestion <> "" Then
                result = Replace(result, currentChar, suggestion)
            End If
        End If
    Next i
    
    ConvertToSJIS = result
End Function

' 変換ボタンを配置するサブルーチン
Sub AddConvertButtons()
    Dim resultWs As Worksheet
    Dim lastRow As Long
    Dim i As Long
    Dim btn As Button
    
    Set resultWs = ThisWorkbook.Worksheets("result")
    lastRow = resultWs.Cells(resultWs.Rows.Count, "A").End(xlUp).Row
    
    On Error Resume Next
    resultWs.Buttons.Delete
    On Error GoTo 0
    
    For i = 2 To lastRow
        Set btn = resultWs.Buttons.Add(resultWs.Cells(i, 7).Left, _
                                     resultWs.Cells(i, 7).Top, _
                                     resultWs.Cells(i, 7).Width - 5, _
                                     resultWs.Cells(i, 7).Height - 2)
        With btn
            .OnAction = "ConvertSelectedCharacter"
            .Caption = "変換"
            .Name = "ConvertBtn_" & i
        End With
    Next i
End Sub

' 変換ボタンのクリックイベントハンドラ
Sub ConvertSelectedCharacter()
    Dim btn As Button
    Dim rowNum As Long
    Dim resultWs As Worksheet
    Dim dataWs As Worksheet
    Dim targetAddress As String
    Dim targetChar As String
    Dim suggestion As String
    
    Set btn = ActiveSheet.Buttons(Application.Caller)
    rowNum = CLng(Split(btn.Name, "_")(1))
    
    Set resultWs = ThisWorkbook.Worksheets("result")
    Set dataWs = ThisWorkbook.Worksheets("Data")
    
    ' 変換対象の情報を取得
    targetAddress = resultWs.Cells(rowNum, 3).Value
    targetChar = resultWs.Cells(rowNum, 4).Value
    suggestion = resultWs.Cells(rowNum, 6).Value
    
    ' 変換の実行
    If suggestion <> "" Then
        ' 正しいワークシートを参照して変換を実行
        dataWs.Range(targetAddress).Value = Replace(dataWs.Range(targetAddress).Value, targetChar, suggestion)
         
        ' 変換済みの表示を更新
        btn.Caption = "完了"
        btn.Enabled = False
    End If
End Sub