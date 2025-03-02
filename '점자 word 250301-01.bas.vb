'점자 word 250301-01
Option Explicit

Const keyword1 As String = "A1"
Const keyword2 As String = "B1"

Public allText As String
Public insertText As String
Public fileNum As Integer
Public byteData() As Byte
Public lines As Variant
Public allLines As Variant
Public fileContent As String
Public doc As Document
Public startLines As Long
Public endLines As Long
Public textLines As Long
Public i As Long
Public j As Long
Public k As Long
Public count As Long
Public extractedLines() As String
Public extractedText As String
Public ResultText As String
Public textContent As String

Sub ProcessWordFile2WithPath(wordFilePath As String, textFilePath As String)
    Set doc = ActiveDocument
    
    textContent = ReadTextFile(textFilePath)
    
    ExtractTextBetweenKeywords textContent, keyword1, keyword2
    
    Call Section1
    Call Section2
End Sub

Function ReadTextFile(filePath As String) As String
    ' ADODBストリームを使用してテキストファイルを読み込む
    Dim stream As Object
    Set stream = CreateObject("ADODB.Stream")

    On Error Resume Next
    stream.Charset = "Shift-JIS"  ' まずShift-JISで試す
    stream.Open
    stream.LoadFromFile filePath

    ReadTextFile = stream.ReadText   ' 関数名を正しく修正
    stream.Close
    Set stream = Nothing
End Function

Function ExtractTextBetweenKeywords(textContent As String, startKeyword As String, endKeyword As String) As String
    Dim StartPos As Long, EndPos As Long
    Dim Line As Variant   ' 追加：Lineを宣言

    ' デバッグ情報
    Debug.Print "テキストの長さ: " & Len(textContent) & " 文字"
    Debug.Print "検索キーワード: " & startKeyword & " から " & endKeyword

    ' キーワード間のテキストを抽出
    StartPos = InStr(1, textContent, startKeyword)
    If StartPos > 0 Then
        Debug.Print "開始キーワード '" & startKeyword & "' が見つかりました。位置: " & StartPos

        StartPos = StartPos + Len(startKeyword)
        EndPos = InStr(StartPos, textContent, endKeyword)

        If EndPos > 0 Then
            Debug.Print "終了キーワード '" & endKeyword & "' が見つかりました。位置: " & EndPos
            extractedText = Mid(textContent, StartPos, EndPos - StartPos)
        Else
            Debug.Print "終了キーワード '" & endKeyword & "' が見つかりませんでした。残りのすべてのテキストを抽出します。"
            extractedText = Mid(textContent, StartPos)
        End If

        ' グローバル変数に行配列を設定
        lines = Split(textContent, vbCrLf)
        If UBound(lines) = 0 Then lines = Split(textContent, vbLf)  ' CRのみの場合に対応

        ResultText = ""   ' 初期化
        For Each Line In lines
            If Len(Line) > 0 Then
                ResultText = ResultText & Line & vbCrLf
            End If
        Next Line
    End If

    ExtractTextBetweenKeywords = extractedText
End Function

Sub Section1()
    ' 抽出された行を処理する例
    Dim textLine As Long   ' 追加：textLineを宣言
    
    insertText = ""
    count = 0   ' 数値で初期化
    extractedText = ""

    For i = LBound(lines) To UBound(lines)
        If InStr(lines(i), "●") > 0 Then
            count = count + 1
            If count = 3 Then
                textLine = i + 2
                Exit For   ' Exit Subではなく、For文を抜ける
            End If
        End If
    Next i
    
    ' textLineが範囲内かチェック
    If textLine >= LBound(lines) And textLine <= UBound(lines) Then
        extractedText = lines(textLine) & vbCrLf
    Else
        extractedText = ""
        MsgBox "指定された行が見つかりませんでした"
        Exit Sub
    End If

    insertText = extractedText

    Dim foundPageBreaks As Integer
    foundPageBreaks = 0
    Dim rng As Range
    Set rng = doc.Range(0, 0)   ' 代入演算子を追加
    
    With rng.Find
        .Text = Chr(12)  ' 改ページコード
        .Forward = True
        .Wrap = wdFindStop
        While .Execute
            foundPageBreaks = foundPageBreaks + 1
            If foundPageBreaks = 9 Then
                Exit While
            End If
            rng.Collapse wdCollapseEnd
        Wend
    End With

    ' 9つ目の改ページの前にテキストを挿入
    If foundPageBreaks = 9 Then
        rng.Collapse Direction:=wdCollapseStart
        rng.InsertBefore insertText
    Else
        MsgBox "文書に9つ目の改ページが見つかりませんでした。"
    End If
End Sub

' Section2の実装（ダミー）
Sub Section2()
    ' ここにSection2のコードを実装してください
    Debug.Print "Section2 実行"
End Sub