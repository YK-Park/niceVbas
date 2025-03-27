' 環算表 250327-04
Public Function ExtractRegistrationParts(regNum As String, Optional allowedPrefixes As String = "") As Object
    Dim parts As Object
    Set parts = CreateObject("Scripting.Dictionary")
    
    ' 기본값으로 초기화
    parts("aValue") = ""
    parts("bValue") = ""
    parts("fValue") = ""
    parts("gValue") = ""
    parts("isValid") = False
    
    ' 입력 검증 - 최소 길이 체크
    If Len(regNum) < 20 Then
        ' 長さが不足している場合はすぐに戻る
        Set ExtractRegistrationParts = parts
        Exit Function
    End If
    
    ' 처음 세 글자 추출 및 검증
    Dim prefix As String
    prefix = Left(regNum, 3)
    
    ' 허용된 접두사와 비교 (allowedPrefixes가 비어있으면 모든 접두사 허용)
    If allowedPrefixes <> "" Then
        ' 쉼표로 구분된 접두사 목록에서 검색
        If InStr(1, allowedPrefixes & ",", prefix & ",") = 0 Then
            ' 허용된 접두사가 아니면 바로 반환
            Set ExtractRegistrationParts = parts
            Exit Function
        End If
    End If
    
    ' 패턴 추출
    Dim tempAValue As String
    tempAValue = Mid(regNum, 6, 4)
    
    ' aValue가 4자리 숫자인지 확인
    If Len(tempAValue) <> 4 Or Not IsNumeric(tempAValue) Then
        ' 有効でない形式ならすぐに戻る
        Set ExtractRegistrationParts = parts
        Exit Function
    End If
    
    ' aValue가 유효하므로 저장
    parts("aValue") = tempAValue
    
    ' bValue 검사 (2자리 숫자인지)
    Dim tempBValue As String
    tempBValue = Mid(regNum, 10, 2)
    
    If Len(tempBValue) <> 2 Or Not IsNumeric(tempBValue) Then
        ' 有効でない形式
        Set ExtractRegistrationParts = parts
        Exit Function
    End If
    
    parts("bValue") = tempBValue
    
    ' fValue 추출
    parts("fValue") = Mid(regNum, 12, 7)
    
    ' gValue는 단일 문자이므로 간단히 추출
    parts("gValue") = Mid(regNum, 19, 1)
    
    ' 모든 검사를 통과했으므로 유효함
    parts("isValid") = True
    
    ' 접두사 정보도 저장 (추후 참조를 위해)
    parts("prefix") = prefix
    
    Set ExtractRegistrationParts = parts
End Function


' 함수 사용 예시
Public Sub ProcessFilesForSyukei(csvFilePath As String, xlsxFilePath As String, resultFilePath As String)
    ' ... 기존 코드 ...
    
    ' 허용된 접두사 정의 (쉼표로 구분)
    Dim allowedPrefixes As String
    allowedPrefixes = "ABC,DEF,XYZ"  ' 처리할 접두사만 지정
    
    ' 유효한 등록번호만 저장할 Dictionary
    Dim validRegistrations As Object
    Set validRegistrations = CreateObject("Scripting.Dictionary")
    
    ' 카운터 초기화
    Dim validCount As Long, invalidCount As Long, skippedByPrefixCount As Long
    validCount = 0
    invalidCount = 0
    skippedByPrefixCount = 0
    
    ' CSV 데이터 처리
    Dim regKey As Variant
    For Each regKey In registrationNumbers.Keys
        Dim regNum As String
        regNum = CStr(regKey)
        
        ' 접두사로 빠른 필터링 (선택적)
        If Left(regNum, 3) <> "ABC" And Left(regNum, 3) <> "DEF" And Left(regNum, 3) <> "XYZ" Then
            skippedByPrefixCount = skippedByPrefixCount + 1
            ' 다음 항목으로 넘어감
            GoTo NextRegistration
        End If
        
        ' 등록번호 처리 - 접두사 검증 포함
        Dim regParts As Object
        Set regParts = ExtractRegistrationParts(regNum, allowedPrefixes)
        
        If regParts("isValid") Then
            ' 유효한 등록번호만 저장
            validRegistrations.Add regNum, regParts
            validCount = validCount + 1
        Else
            invalidCount = invalidCount + 1
        End If
        
NextRegistration:
    Next regKey
    
    ' 진행 상황 업데이트
    Application.StatusBar = "유효: " & validCount & ", 무효: " & invalidCount & ", 접두사 스킵: " & skippedByPrefixCount
    
    ' ... 나머지 코드 ...
End Sub